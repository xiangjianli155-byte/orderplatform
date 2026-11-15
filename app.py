from __future__ import annotations

import hashlib
import json
import re
from datetime import date
from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd
import plotly.express as px
import streamlit as st
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


APP_TITLE = "订单分析"
NO_COLUMN = "— 不使用 —"

# 常见列名会自动建议匹配；用户可以在界面中改选任何列。
FIELD_CONFIG: list[dict[str, Any]] = [
    {"key": "customer", "label": "客户", "help": "客户名称或客户号", "aliases": ["customer name", "customer", "customer no", "客户名称", "客户", "客户编号"]},
    {"key": "order", "label": "订单 / PO", "help": "订单号或客户采购订单号", "aliases": ["order no", "order number", "sales order", "customer po", "po number", "订单号", "采购订单", "客户po"]},
    {"key": "line", "label": "订单行", "help": "PO 行号或订单行号", "aliases": ["customer po line", "po line", "order line", "line number", "行号", "订单行"]},
    {"key": "item", "label": "物料号", "help": "料号、SKU 或 Item", "aliases": ["part no", "part number", "item number", "sku", "material", "料号", "物料号", "产品编号"]},
    {"key": "description", "label": "物料描述", "help": "产品名称或描述", "aliases": ["description", "item name", "part description", "产品描述", "物料描述", "品名"]},
    {"key": "ordered_qty", "label": "原始订单数量", "help": "用于计算未结量的备选字段", "aliases": ["order qty", "ordered qty", "qty ordered", "订单数量", "订购数量"]},
    {"key": "open_qty", "label": "未结数量", "help": "优先使用源表已有的 Qty Open", "aliases": ["qty open", "open qty", "open quantity", "remaining qty", "未结数量", "未交数量", "开放数量"]},
    {"key": "shipped_qty", "label": "已发货数量", "help": "与原始订单数量搭配计算未结量", "aliases": ["qty shipped", "shipped qty", "shipped quantity", "已发货数量", "发货数量", "出货数量"]},
    {"key": "backorder_qty", "label": "欠交数量", "help": "作为独立指标展示", "aliases": ["qty backordered", "backorder qty", "backordered qty", "欠交数量", "缺货数量", "欠货数量"]},
    {"key": "reserved_qty", "label": "预留数量", "help": "作为独立指标展示", "aliases": ["qty reserved", "qty resvered", "reserved qty", "预留数量", "分配数量"]},
    {"key": "picked_qty", "label": "已拣数量", "help": "作为独立指标展示", "aliases": ["qty picked", "picked qty", "已拣数量", "拣货数量"]},
    {"key": "due_date", "label": "承诺 / 修订交期", "help": "用于逾期和近期到期提醒", "aliases": ["revised due date", "due date", "promise date", "requested date", "承诺交期", "修订交期", "交期", "需求日期"]},
    {"key": "actual_date", "label": "实际发货 / 交付日期", "help": "显示在明细中；计算准时率前需确认数据粒度", "aliases": ["actual delivery date", "actual ship date", "ship date", "delivery date", "实际交付日期", "实际发货日期", "出货日期"]},
    {"key": "open_value", "label": "未结金额", "help": "若源表有金额，可汇总金额风险", "aliases": ["total value open", "open value", "open amount", "未结金额", "未交金额"]},
    {"key": "carrier", "label": "承运商", "help": "可选筛选维度", "aliases": ["carrier code", "carrier", "承运商", "运输商"]},
    {"key": "destination", "label": "目的地", "help": "可选筛选维度", "aliases": ["dest location", "destination", "ship to", "目的地", "收货地点"]},
]


def normalize_name(value: Any) -> str:
    text = str(value).strip().casefold()
    return re.sub(r"[\s_\-./\\()[\]（）]+", "", text)


def unique_column_names(columns: Any) -> list[str]:
    result: list[str] = []
    counts: dict[str, int] = {}
    for index, value in enumerate(columns):
        base = str(value).strip() if value is not None and not pd.isna(value) else ""
        if not base or base.casefold().startswith("unnamed:"):
            base = f"列 {index + 1}"
        counts[base] = counts.get(base, 0) + 1
        result.append(base if counts[base] == 1 else f"{base} ({counts[base]})")
    return result


def suggest_column(columns: list[str], aliases: list[str]) -> str | None:
    normalized_columns = {column: normalize_name(column) for column in columns}
    normalized_aliases = [normalize_name(alias) for alias in aliases]

    for alias in normalized_aliases:
        for column, normalized in normalized_columns.items():
            if normalized == alias:
                return column

    for alias in normalized_aliases:
        if len(alias) < 4:
            continue
        for column, normalized in normalized_columns.items():
            if alias in normalized or normalized in alias:
                return column
    return None


@st.cache_data(show_spinner=False)
def get_sheet_names(file_bytes: bytes) -> tuple[str, ...]:
    excel_file = pd.ExcelFile(BytesIO(file_bytes), engine="calamine")
    return tuple(excel_file.sheet_names)


@st.cache_data(show_spinner=False)
def read_preview(file_bytes: bytes, sheet_name: str, rows: int = 30) -> pd.DataFrame:
    return pd.read_excel(
        BytesIO(file_bytes),
        sheet_name=sheet_name,
        header=None,
        nrows=rows,
        engine="calamine",
    )


@st.cache_data(show_spinner=False)
def read_table(file_bytes: bytes, sheet_name: str, header_row_zero_based: int) -> pd.DataFrame:
    frame = pd.read_excel(
        BytesIO(file_bytes),
        sheet_name=sheet_name,
        header=header_row_zero_based,
        dtype=object,
        engine="calamine",
    )
    frame.columns = unique_column_names(frame.columns)
    return frame.dropna(how="all")


def guess_header_row(preview: pd.DataFrame) -> int:
    if preview.empty:
        return 0
    scores: list[tuple[float, int]] = []
    for index, row in preview.iterrows():
        non_empty = int(row.notna().sum())
        text_cells = sum(isinstance(value, str) and bool(value.strip()) for value in row.tolist())
        scores.append((non_empty + 0.2 * text_cells, int(index)))
    return max(scores)[1] if scores else 0


def numeric_series(series: pd.Series) -> pd.Series:
    cleaned = series.astype("string").str.strip()
    cleaned = cleaned.str.replace(",", "", regex=False).str.replace("$", "", regex=False)
    cleaned = cleaned.str.replace("¥", "", regex=False).str.replace("￥", "", regex=False)
    cleaned = cleaned.str.replace(r"^\((.*)\)$", r"-\1", regex=True)
    return pd.to_numeric(cleaned, errors="coerce")


def date_series(series: pd.Series) -> pd.Series:
    numeric = pd.to_numeric(series, errors="coerce")
    valid = numeric.dropna()
    if not valid.empty and valid.between(20000, 80000).mean() >= 0.8:
        return pd.to_datetime(numeric, unit="D", origin="1899-12-30", errors="coerce")
    return pd.to_datetime(series, errors="coerce", format="mixed")


def source_series(frame: pd.DataFrame, column: str | None, default: Any = pd.NA) -> pd.Series:
    if column and column in frame.columns:
        return frame[column]
    return pd.Series(default, index=frame.index)


def build_analysis_frame(
    source: pd.DataFrame,
    mapping: dict[str, str | None],
    header_row_excel: int,
) -> tuple[pd.DataFrame | None, str, dict[str, int]]:
    source = source.copy()
    excel_rows = source.index + header_row_excel + 1

    open_column = mapping.get("open_qty")
    ordered_column = mapping.get("ordered_qty")
    shipped_column = mapping.get("shipped_qty")
    if open_column:
        open_qty = numeric_series(source_series(source, open_column))
        quantity_note = f"使用源字段“{open_column}”作为未结数量。"
    elif ordered_column and shipped_column:
        ordered = numeric_series(source_series(source, ordered_column))
        shipped = numeric_series(source_series(source, shipped_column))
        open_qty = (ordered - shipped).clip(lower=0)
        quantity_note = f"按“{ordered_column} - {shipped_column}”计算未结数量，并将负值按 0 计。"
    else:
        return None, "请映射未结数量，或同时映射原始订单数量和已发货数量。", {}

    due = date_series(source_series(source, mapping.get("due_date")))
    table = pd.DataFrame(index=source.index)
    table["Excel行"] = excel_rows
    table["客户"] = source_series(source, mapping.get("customer"), "未指定").fillna("未指定").astype(str)
    table["订单号"] = source_series(source, mapping.get("order"), "").fillna("").astype(str)
    table["订单行"] = source_series(source, mapping.get("line"), "").fillna("").astype(str)
    table["物料号"] = source_series(source, mapping.get("item"), "").fillna("").astype(str)
    table["物料描述"] = source_series(source, mapping.get("description"), "").fillna("").astype(str)
    table["承诺交期"] = due
    table["实际日期"] = date_series(source_series(source, mapping.get("actual_date")))
    table["未结数量"] = open_qty
    table["欠交数量"] = numeric_series(source_series(source, mapping.get("backorder_qty")))
    table["预留数量"] = numeric_series(source_series(source, mapping.get("reserved_qty")))
    table["已拣数量"] = numeric_series(source_series(source, mapping.get("picked_qty")))
    table["已发货数量"] = numeric_series(source_series(source, mapping.get("shipped_qty")))
    table["未结金额"] = numeric_series(source_series(source, mapping.get("open_value")))
    table["承运商"] = source_series(source, mapping.get("carrier"), "").fillna("").astype(str)
    table["目的地"] = source_series(source, mapping.get("destination"), "").fillna("").astype(str)

    stats: dict[str, int] = {
        "rows": int(len(table)),
        "missing_open_qty": int(table["未结数量"].isna().sum()),
        "missing_due_date": int(table["承诺交期"].isna().sum()),
        "negative_open_qty": int((table["未结数量"] < 0).fillna(False).sum()),
        "overfulfilled_rows": 0,
        "duplicate_keys": 0,
        "quantity_mismatch": 0,
    }
    if ordered_column and shipped_column:
        ordered = numeric_series(source_series(source, ordered_column))
        shipped = numeric_series(source_series(source, shipped_column))
        stats["overfulfilled_rows"] = int(((shipped > ordered) & ordered.notna() & shipped.notna()).sum())

    key_columns: list[str] = []
    for key, output_name in [("order", "订单号"), ("line", "订单行"), ("item", "物料号")]:
        if mapping.get(key):
            key_columns.append(output_name)
    if len(key_columns) >= 2:
        key_frame = table[key_columns].replace("", pd.NA)
        complete_key = key_frame.notna().all(axis=1)
        stats["duplicate_keys"] = int(table.loc[complete_key].duplicated(key_columns, keep=False).sum())

    if open_column and all(mapping.get(key) for key in ("backorder_qty", "reserved_qty", "picked_qty")):
        components = ["欠交数量", "预留数量", "已拣数量"]
        complete = table[["未结数量", *components]].notna().all(axis=1)
        component_sum = table[components].sum(axis=1)
        mismatch = (table["未结数量"] - component_sum).abs() > 1e-9
        stats["quantity_mismatch"] = int((complete & mismatch).sum())
    return table, quantity_note, stats


def weekly_open_qty(frame: pd.DataFrame) -> pd.Series:
    valid = frame.dropna(subset=["承诺交期"]).copy()
    if valid.empty:
        return pd.Series(dtype="float64")
    week_start = valid["承诺交期"] - pd.to_timedelta(valid["承诺交期"].dt.weekday, unit="D")
    valid["交期周"] = week_start.dt.strftime("%Y-%m-%d")
    return valid.groupby("交期周")["未结数量"].sum().sort_index()


def format_number(value: Any) -> str:
    if pd.isna(value):
        return "—"
    number = float(value)
    return f"{number:,.0f}" if number.is_integer() else f"{number:,.2f}"


def add_text(
    slide: Any,
    text: str,
    left: float,
    top: float,
    width: float,
    height: float,
    size: int = 16,
    color: RGBColor = RGBColor(35, 41, 49),
    bold: bool = False,
    align: PP_ALIGN = PP_ALIGN.LEFT,
) -> Any:
    shape = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    frame = shape.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.margin_left = Inches(0.04)
    frame.margin_right = Inches(0.04)
    frame.margin_top = Inches(0.02)
    frame.margin_bottom = Inches(0.02)
    frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    paragraph = frame.paragraphs[0]
    paragraph.text = str(text)
    paragraph.alignment = align
    paragraph.font.name = "Microsoft YaHei"
    paragraph.font.size = Pt(size)
    paragraph.font.bold = bold
    paragraph.font.color.rgb = color
    return shape


def add_slide_header(slide: Any, title: str, subtitle: str) -> None:
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.12))
    bar.fill.solid()
    bar.fill.fore_color.rgb = RGBColor(185, 35, 46)
    bar.line.fill.background()
    add_text(slide, title, 0.55, 0.30, 12.1, 0.42, 23, bold=True)
    add_text(slide, subtitle, 0.58, 0.80, 12.0, 0.30, 10, color=RGBColor(100, 108, 118))


def build_pptx(
    frame: pd.DataFrame,
    source_name: str,
    sheet_name: str,
    report_date: date,
    quantity_note: str,
    quality_stats: dict[str, int],
) -> bytes:
    presentation = Presentation()
    presentation.slide_width = Inches(13.333)
    presentation.slide_height = Inches(7.5)
    blank = presentation.slide_layouts[6]
    dark = RGBColor(35, 41, 49)
    muted = RGBColor(100, 108, 118)
    accent = RGBColor(185, 35, 46)

    slide = presentation.slides.add_slide(blank)
    band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.17))
    band.fill.solid()
    band.fill.fore_color.rgb = accent
    band.line.fill.background()
    add_text(slide, "供应链规划分析", 0.75, 1.45, 11.8, 0.7, 30, dark, True)
    add_text(slide, "客户订单与交付", 0.78, 2.25, 11.6, 0.45, 20, accent, True)
    add_text(slide, f"报告日期：{report_date:%Y-%m-%d}", 0.8, 3.15, 11.4, 0.35, 15, dark)
    add_text(slide, f"来源：{source_name}　|　工作表：{sheet_name}", 0.8, 3.62, 11.4, 0.35, 11, muted)
    add_text(slide, "按本次筛选结果生成", 0.8, 6.55, 11.4, 0.28, 10, muted)

    slide = presentation.slides.add_slide(blank)
    add_slide_header(slide, "订单履约概览", f"报告截至 {report_date:%Y-%m-%d} · 当前筛选范围")
    open_qty = frame["未结数量"].fillna(0).sum()
    overdue = frame["状态"] == "逾期"
    due_soon = frame["状态"] == "7天内到期"
    backorder = frame["欠交数量"].sum(min_count=1)
    metrics = [
        ("未结订单行", f"{int((frame['未结数量'].fillna(0) > 0).sum()):,}"),
        ("未结数量", format_number(open_qty)),
        ("逾期未结行", f"{int(overdue.sum()):,}"),
        ("7天内到期行", f"{int(due_soon.sum()):,}"),
    ]
    for index, (label, value) in enumerate(metrics):
        x = 0.6 + index * 3.15
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.5), Inches(2.85), Inches(1.25))
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(246, 247, 249)
        shape.line.color.rgb = RGBColor(225, 228, 232)
        add_text(slide, label, x + 0.15, 1.68, 2.55, 0.28, 11, muted)
        add_text(slide, value, x + 0.15, 2.0, 2.55, 0.48, 23, dark, True)
    add_text(slide, f"欠交数量（独立指标）：{format_number(backorder)}", 0.7, 3.15, 11.8, 0.35, 15, accent, True)
    add_text(slide, quantity_note, 0.72, 3.62, 11.8, 0.55, 11, muted)

    slide = presentation.slides.add_slide(blank)
    add_slide_header(slide, "未结数量按交期周分布", "柱形图为 PowerPoint 原生图表，可在 PowerPoint 中编辑")
    weekly = weekly_open_qty(frame)
    if not weekly.empty:
        weekly = weekly.tail(16)
        chart_data = CategoryChartData()
        chart_data.categories = list(weekly.index)
        chart_data.add_series("未结数量", tuple(float(value) for value in weekly.values))
        chart_frame = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            Inches(0.75),
            Inches(1.45),
            Inches(11.8),
            Inches(5.35),
            chart_data,
        )
        chart = chart_frame.chart
        chart.has_legend = False
        chart.has_title = False
        chart.category_axis.tick_labels.font.name = "Microsoft YaHei"
        chart.category_axis.tick_labels.font.size = Pt(9)
        chart.value_axis.tick_labels.font.name = "Microsoft YaHei"
        chart.value_axis.tick_labels.font.size = Pt(9)
        chart.value_axis.has_major_gridlines = True
    else:
        add_text(slide, "没有可用的承诺交期，无法生成交期趋势图。", 1.0, 2.5, 11.0, 0.5, 16, muted)

    slide = presentation.slides.add_slide(blank)
    add_slide_header(slide, "优先处理订单行", "按逾期状态、交期和未结数量排列")
    columns = ["客户", "订单号", "订单行", "物料号", "承诺交期", "未结数量", "欠交数量"]
    top = frame.sort_values(
        ["状态优先级", "承诺交期", "未结数量"],
        ascending=[True, True, False],
        na_position="last",
    ).head(12)
    table = slide.shapes.add_table(
        max(2, len(top) + 1),
        len(columns),
        Inches(0.55),
        Inches(1.42),
        Inches(12.2),
        Inches(5.25),
    ).table
    widths = [1.8, 1.55, 0.85, 1.65, 1.35, 1.25, 1.25]
    for index, width in enumerate(widths):
        table.columns[index].width = Inches(width)
    for col_idx, name in enumerate(columns):
        cell = table.cell(0, col_idx)
        cell.text = name
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(242, 244, 247)
        for paragraph in cell.text_frame.paragraphs:
            paragraph.font.name = "Microsoft YaHei"
            paragraph.font.size = Pt(9)
            paragraph.font.bold = True
            paragraph.font.color.rgb = dark
    if top.empty:
        table.cell(1, 0).text = "当前筛选下没有订单行"
    else:
        for row_idx, (_, row) in enumerate(top.iterrows(), start=1):
            for col_idx, name in enumerate(columns):
                value = row.get(name, "")
                if name == "承诺交期":
                    value = value.strftime("%Y-%m-%d") if pd.notna(value) else "未提供"
                elif name in ("未结数量", "欠交数量"):
                    value = format_number(value)
                cell = table.cell(row_idx, col_idx)
                cell.text = str(value)[:48]
                for paragraph in cell.text_frame.paragraphs:
                    paragraph.font.name = "Microsoft YaHei"
                    paragraph.font.size = Pt(8)
                    paragraph.font.color.rgb = dark
    add_text(slide, "完整订单行明细可从看板下载为 CSV。", 0.62, 6.82, 11.8, 0.25, 9, muted)

    slide = presentation.slides.add_slide(blank)
    add_slide_header(slide, "数据口径与质量检查", "用于说明本次分析的范围和可用字段")
    add_text(slide, "数量口径", 0.7, 1.45, 2.2, 0.35, 15, accent, True)
    add_text(slide, quantity_note, 0.72, 1.85, 11.7, 0.58, 12, dark)
    checks = [
        f"源数据行数：{quality_stats.get('rows', 0):,}",
        f"缺少未结数量：{quality_stats.get('missing_open_qty', 0):,} 行",
        f"缺少交期：{quality_stats.get('missing_due_date', 0):,} 行",
        f"重复订单行键：{quality_stats.get('duplicate_keys', 0):,} 行",
        f"开放量与预留 / 已拣 / 欠交合计不一致：{quality_stats.get('quantity_mismatch', 0):,} 行",
    ]
    add_text(slide, "检查结果", 0.7, 2.65, 2.2, 0.35, 15, accent, True)
    for index, text in enumerate(checks):
        add_text(slide, "• " + text, 0.82, 3.08 + index * 0.52, 11.5, 0.32, 12, dark)
    add_text(slide, f"来源文件：{source_name}　|　工作表：{sheet_name}", 0.72, 6.38, 11.8, 0.3, 10, muted)

    output = BytesIO()
    presentation.save(output)
    return output.getvalue()


def numeric_columns(frame: pd.DataFrame) -> dict[str, pd.Series]:
    result: dict[str, pd.Series] = {}
    for column in frame.columns:
        values = numeric_series(frame[column])
        if values.notna().any():
            result[column] = values
    return result


def aggregate_generic(
    frame: pd.DataFrame,
    dimension: str | None,
    measure: str | None,
    aggregation: str,
    numeric_values: dict[str, pd.Series],
) -> pd.DataFrame:
    working = pd.DataFrame(index=frame.index)
    if dimension:
        working["分类"] = frame[dimension].fillna("（空白）").astype(str)
    else:
        working["分类"] = "全部数据"

    if aggregation == "计数" or not measure:
        result = working.groupby("分类", dropna=False).size().rename("值").reset_index()
    else:
        working["指标"] = numeric_values[measure]
        method = "mean" if aggregation == "平均值" else "sum"
        result = working.groupby("分类", dropna=False)["指标"].agg(method).rename("值").reset_index()
        result = result.dropna(subset=["值"])
    return result.sort_values("值", ascending=False).head(20).reset_index(drop=True)


def build_generic_pptx(
    source_name: str,
    sheet_name: str,
    report_date: date,
    dimension: str | None,
    measure: str | None,
    aggregation: str,
    grouped: pd.DataFrame,
    row_count: int,
    metric_total: float,
    missing_dimension: int,
    missing_measure: int,
) -> bytes:
    presentation = Presentation()
    presentation.slide_width = Inches(13.333)
    presentation.slide_height = Inches(7.5)
    blank = presentation.slide_layouts[6]
    dark = RGBColor(35, 41, 49)
    muted = RGBColor(100, 108, 118)
    accent = RGBColor(185, 35, 46)
    metric_label = "记录数" if aggregation == "计数" or not measure else f"{aggregation}：{measure}"
    dimension_label = dimension or "全部数据"

    slide = presentation.slides.add_slide(blank)
    band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.17))
    band.fill.solid()
    band.fill.fore_color.rgb = accent
    band.line.fill.background()
    add_text(slide, "供应链数据分析", 0.75, 1.45, 11.8, 0.7, 30, dark, True)
    add_text(slide, f"分析维度：{dimension_label}", 0.78, 2.25, 11.6, 0.45, 20, accent, True)
    add_text(slide, f"报告日期：{report_date:%Y-%m-%d}", 0.8, 3.15, 11.4, 0.35, 15, dark)
    add_text(slide, f"来源：{source_name}　|　工作表：{sheet_name}", 0.8, 3.62, 11.4, 0.35, 11, muted)
    add_text(slide, "按当前筛选范围生成", 0.8, 6.55, 11.4, 0.28, 10, muted)

    slide = presentation.slides.add_slide(blank)
    add_slide_header(slide, "数据概览", f"汇总方式：{metric_label}　|　分组字段：{dimension_label}")
    add_text(slide, "数据行数", 0.75, 1.55, 3.0, 0.3, 12, muted)
    add_text(slide, f"{row_count:,}", 0.75, 1.9, 3.0, 0.6, 26, dark, True)
    add_text(slide, "分组数量", 4.2, 1.55, 3.0, 0.3, 12, muted)
    add_text(slide, f"{len(grouped):,}", 4.2, 1.9, 3.0, 0.6, 26, dark, True)
    add_text(slide, metric_label, 7.6, 1.55, 4.7, 0.3, 12, muted)
    add_text(slide, format_number(metric_total), 7.6, 1.9, 4.7, 0.6, 26, accent, True)
    if not grouped.empty:
        chart_data = CategoryChartData()
        chart_data.categories = [str(value)[:28] for value in grouped["分类"]]
        chart_data.add_series(metric_label, tuple(float(value) for value in grouped["值"]))
        chart_frame = slide.shapes.add_chart(
            XL_CHART_TYPE.BAR_CLUSTERED,
            Inches(0.75),
            Inches(2.85),
            Inches(11.8),
            Inches(4.0),
            chart_data,
        )
        chart = chart_frame.chart
        chart.has_legend = False
        chart.has_title = False
        chart.category_axis.tick_labels.font.name = "Microsoft YaHei"
        chart.category_axis.tick_labels.font.size = Pt(9)
        chart.value_axis.tick_labels.font.name = "Microsoft YaHei"
        chart.value_axis.tick_labels.font.size = Pt(9)
        chart.value_axis.has_major_gridlines = True
    else:
        add_text(slide, "当前筛选范围内没有可汇总的数据。", 1.0, 3.7, 11.0, 0.5, 16, muted)

    slide = presentation.slides.add_slide(blank)
    add_slide_header(slide, "分组结果与数据检查", "展示当前分组下的前 20 个结果")
    add_text(slide, f"缺少分组值：{missing_dimension:,} 行", 0.72, 1.12, 5.5, 0.3, 10, muted)
    add_text(slide, f"缺少数值：{missing_measure:,} 行", 6.1, 1.12, 5.5, 0.3, 10, muted)
    row_count_table = min(16, len(grouped) + 1)
    table = slide.shapes.add_table(
        row_count_table,
        2,
        Inches(0.7),
        Inches(1.6),
        Inches(8.5),
        Inches(5.2),
    ).table
    table.columns[0].width = Inches(5.5)
    table.columns[1].width = Inches(3.0)
    for col_idx, name in enumerate([dimension_label, metric_label]):
        cell = table.cell(0, col_idx)
        cell.text = name
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(242, 244, 247)
        for paragraph in cell.text_frame.paragraphs:
            paragraph.font.name = "Microsoft YaHei"
            paragraph.font.size = Pt(10)
            paragraph.font.bold = True
            paragraph.font.color.rgb = dark
    for row_idx, row in enumerate(grouped.head(row_count_table - 1).itertuples(index=False), start=1):
        table.cell(row_idx, 0).text = str(row[0])[:60]
        table.cell(row_idx, 1).text = format_number(row[1])
        for col_idx in range(2):
            for paragraph in table.cell(row_idx, col_idx).text_frame.paragraphs:
                paragraph.font.name = "Microsoft YaHei"
                paragraph.font.size = Pt(9)
                paragraph.font.color.rgb = dark
    add_text(slide, f"来源：{source_name}　|　{sheet_name}　|　{report_date:%Y-%m-%d}", 0.72, 6.95, 11.8, 0.22, 9, muted)

    output = BytesIO()
    presentation.save(output)
    return output.getvalue()


def render_empty_state() -> None:
    st.title("订单分析")
    st.write("上传 Excel 后选择分析模式。系统会列出工作表、建议表头，并允许你按当前工作表设置分析字段。")
    left, right = st.columns([1.35, 1])
    with left:
        st.subheader("一次导入，重复使用")
        st.markdown(
            "- 支持 .xlsx、.xls、.xlsm、.xlsb 工作簿及多个工作表。\n"
            "- 字段名称不同也可以手动匹配到客户、订单、料号、数量和交期。\n"
            "- 可下载字段映射配置，供相同格式的文件复用。"
        )
    with right:
        st.subheader("从模板开始")
        st.markdown("客户订单与交付 · 通用表格分析")
        st.info("通用模式可按任意分类列汇总任意数值列，并导出 PowerPoint。")


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="📊", layout="wide")
    st.markdown(
        """
        <style>
        .block-container {padding-top: 1.2rem; padding-bottom: 2rem; max-width: 1500px;}
        [data-testid="stMetric"] {background: rgba(128,128,128,.06); padding: 14px 16px; border-radius: 10px;}
        div[data-testid="stDataFrame"] {border: 1px solid rgba(128,128,128,.15); border-radius: 8px;}
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.sidebar.title("供应链工作台")
    st.sidebar.caption("通用 Excel 导入 · 供应链分析 · PPT 汇报")
    analysis_mode = st.sidebar.selectbox(
        "分析模式",
        ["客户订单与交付", "通用表格分析"],
        help="订单模式按业务字段计算交期风险；通用模式适合其他结构化 Excel。",
    )
    uploaded = st.sidebar.file_uploader(
        "上传 Excel 工作簿",
        type=["xlsx", "xls", "xlsm", "xlsb"],
        help="本版本每次分析一份工作簿，可在工作簿中选择任意工作表。",
    )
    profile_upload = None
    if analysis_mode == "客户订单与交付":
        profile_upload = st.sidebar.file_uploader(
            "载入字段映射 JSON（可选）",
            type=["json"],
            help="用于复用之前下载的字段映射配置。",
            key="mapping_profile_upload",
        )
    saved_mapping: dict[str, str] = {}
    if profile_upload is not None:
        try:
            profile_content = json.loads(profile_upload.getvalue().decode("utf-8"))
            candidate_mapping = profile_content.get("mapping", {})
            if isinstance(candidate_mapping, dict):
                saved_mapping = candidate_mapping
            else:
                st.sidebar.error("映射 JSON 中的 mapping 需要是对象。")
        except (UnicodeDecodeError, json.JSONDecodeError, AttributeError):
            st.sidebar.error("映射 JSON 无法读取。")
    if uploaded is None:
        render_empty_state()
        st.stop()

    file_bytes = uploaded.getvalue()
    file_hash = hashlib.sha1(file_bytes).hexdigest()[:10]
    try:
        sheet_names = get_sheet_names(file_bytes)
    except Exception as exc:
        st.error(f"无法读取工作簿：{exc}")
        st.info("请确认文件未加密、格式有效，并且安装了 requirements.txt 中的依赖。")
        st.stop()

    st.sidebar.caption(f"文件：{uploaded.name}")
    if not sheet_names:
        st.error("该工作簿没有可读取的工作表。")
        st.stop()
    sheet_name = st.sidebar.selectbox("选择工作表", options=list(sheet_names))
    try:
        preview = read_preview(file_bytes, sheet_name)
    except Exception as exc:
        st.error(f"无法读取工作表：{exc}")
        st.stop()

    suggested_header = guess_header_row(preview)
    max_header_row = max(30, len(preview))
    header_row_excel = st.sidebar.number_input(
        "表头所在行（Excel 行号）",
        min_value=1,
        max_value=max_header_row,
        value=min(suggested_header + 1, max_header_row),
        step=1,
        help="如果列名不在第一行，可在这里调整。",
        key=f"header_row_{file_hash}_{sheet_name}",
    )
    report_date = st.sidebar.date_input("报告截至日期", value=date.today())
    try:
        source = read_table(file_bytes, sheet_name, int(header_row_excel) - 1)
    except Exception as exc:
        st.error(f"无法将该行作为表头读取：{exc}")
        st.stop()
    if source.empty or len(source.columns) == 0:
        st.warning("当前工作表或表头行没有读到数据。请更换工作表或调整表头行。")
        st.stop()

    st.title("订单分析")
    st.caption(f"来源：{uploaded.name}　|　工作表：{sheet_name}　|　数据行：{len(source):,}　|　报告日期：{report_date:%Y-%m-%d}")
    map_tab, dashboard_tab, ppt_tab = st.tabs(["① 字段映射", "② 分析看板", "③ PPT 报告"])

    if analysis_mode == "通用表格分析":
        numeric_values = numeric_columns(source)
        source_columns = list(source.columns)
        generic_source = source.copy()
        source_row_label = "Excel行"
        while source_row_label in generic_source.columns:
            source_row_label += "_"
        generic_source.insert(0, source_row_label, source.index + int(header_row_excel) + 1)
        generic_filtered = generic_source.copy()
        generic_grouped = pd.DataFrame(columns=["分类", "值"])
        generic_dimension: str | None = None
        generic_measure: str | None = None
        generic_aggregation = "计数"
        generic_metric_total = float(len(source))
        generic_missing_dimension = 0
        generic_missing_measure = 0

        with map_tab:
            st.subheader("当前工作表结构")
            st.write("此模式不要求固定的订单字段。所有列均可作为分类维度；能转换为数字的列可作为汇总指标。")
            schema_rows = []
            for column in source_columns:
                values = numeric_series(source[column])
                schema_rows.append({
                    "列名": column,
                    "非空行数": int(source[column].notna().sum()),
                    "唯一值数": int(source[column].nunique(dropna=True)),
                    "数值可识别行数": int(values.notna().sum()),
                })
            st.dataframe(pd.DataFrame(schema_rows), use_container_width=True, hide_index=True)
            with st.expander("预览当前工作表"):
                st.dataframe(source.head(30), use_container_width=True, hide_index=True)

        with dashboard_tab:
            st.subheader("选择一个分类维度和一个指标")
            choices = [NO_COLUMN, *source_columns]
            numeric_choices = [NO_COLUMN, *numeric_values.keys()]
            controls = st.columns(4)
            selected_dimension = controls[0].selectbox("按什么分类", choices, index=1 if source_columns else 0)
            selected_measure = controls[1].selectbox("汇总哪列数字", numeric_choices, index=1 if numeric_values else 0)
            selected_aggregation = controls[2].selectbox("汇总方式", ["求和", "平均值", "计数"], index=0 if numeric_values else 2)
            search_text = controls[3].text_input("搜索行内容")
            generic_dimension = None if selected_dimension == NO_COLUMN else selected_dimension
            generic_measure = None if selected_measure == NO_COLUMN else selected_measure
            generic_aggregation = selected_aggregation

            if search_text.strip():
                term = re.escape(search_text.strip().casefold())
                searchable = generic_source.astype("string").fillna("").agg(" ".join, axis=1).str.casefold()
                generic_filtered = generic_source[searchable.str.contains(term, na=False)]

            if generic_dimension:
                generic_missing_dimension = int(generic_filtered[generic_dimension].isna().sum())
            if generic_measure:
                selected_values = numeric_values[generic_measure].reindex(generic_filtered.index)
                generic_missing_measure = int(selected_values.isna().sum())
                if generic_aggregation == "求和":
                    metric_value = selected_values.sum(min_count=1)
                    generic_metric_total = float(metric_value) if pd.notna(metric_value) else 0.0
                elif generic_aggregation == "平均值":
                    metric_value = selected_values.mean()
                    generic_metric_total = float(metric_value) if pd.notna(metric_value) else 0.0
                else:
                    generic_metric_total = float(len(generic_filtered))
            else:
                generic_metric_total = float(len(generic_filtered))

            generic_grouped = aggregate_generic(
                generic_filtered,
                generic_dimension,
                generic_measure,
                generic_aggregation,
                numeric_values,
            )
            generic_label = "记录数" if generic_aggregation == "计数" or not generic_measure else f"{generic_aggregation}：{generic_measure}"
            metric_cols = st.columns(3)
            metric_cols[0].metric("数据行数", f"{len(generic_filtered):,}")
            metric_cols[1].metric("分类数量", f"{generic_filtered[generic_dimension].nunique(dropna=True):,}" if generic_dimension else "1")
            metric_cols[2].metric(generic_label, format_number(generic_metric_total))

            chart_col, result_col = st.columns([1.35, 1])
            with chart_col:
                st.subheader("分类汇总")
                if generic_grouped.empty:
                    st.info("当前筛选范围内没有可汇总的值。")
                else:
                    fig = px.bar(
                        generic_grouped,
                        x="分类",
                        y="值",
                        color_discrete_sequence=["#b9232e"],
                    )
                    fig.update_layout(margin=dict(l=10, r=10, t=15, b=10), height=340, xaxis_title=generic_dimension or "", yaxis_title=generic_label, showlegend=False)
                    st.plotly_chart(fig, use_container_width=True)
            with result_col:
                st.subheader("数据检查")
                st.write(f"当前筛选行数：**{len(generic_filtered):,}**")
                st.write(f"分类字段空白：**{generic_missing_dimension:,}** 行")
                st.write(f"指标字段空白或非数字：**{generic_missing_measure:,}** 行")
                st.caption("图表展示数值最高的 20 个分类。")

            st.subheader("原始明细")
            st.dataframe(generic_filtered.head(500), use_container_width=True, hide_index=True)
            st.download_button(
                "下载当前筛选明细 CSV",
                data=generic_filtered.to_csv(index=False).encode("utf-8-sig"),
                file_name="通用分析明细.csv",
                mime="text/csv",
            )
            st.download_button(
                "下载分类汇总 CSV",
                data=generic_grouped.to_csv(index=False).encode("utf-8-sig"),
                file_name="分类汇总.csv",
                mime="text/csv",
            )

        with ppt_tab:
            st.subheader("导出当前通用分析")
            st.write("PPT 包含数据概览、可编辑分类图表、前 20 个分类结果和缺失值检查。")
            st.caption(f"当前纳入 PPT 的行数：{len(generic_filtered):,}　|　分组字段：{generic_dimension or '全部数据'}")
            generic_signature = hashlib.sha1(
                (
                    file_hash + sheet_name + str(report_date) + str(generic_dimension)
                    + str(generic_measure) + generic_aggregation
                    + ",".join(generic_filtered.index.astype(str))
                ).encode("utf-8")
            ).hexdigest()
            if st.button("生成通用分析 PPTX", type="primary"):
                try:
                    st.session_state["generic_ppt_bytes"] = build_generic_pptx(
                        uploaded.name,
                        sheet_name,
                        report_date,
                        generic_dimension,
                        generic_measure,
                        generic_aggregation,
                        generic_grouped,
                        len(generic_filtered),
                        generic_metric_total,
                        generic_missing_dimension,
                        generic_missing_measure,
                    )
                    st.session_state["generic_ppt_signature"] = generic_signature
                except Exception as exc:
                    st.error(f"生成 PPT 时出错：{exc}")
            if st.session_state.get("generic_ppt_signature") == generic_signature and st.session_state.get("generic_ppt_bytes"):
                safe_stem = re.sub(r"[^\w\-]+", "_", Path(uploaded.name).stem).strip("_") or "excel_analysis"
                st.download_button(
                    "下载可编辑 PowerPoint",
                    data=st.session_state["generic_ppt_bytes"],
                    file_name=f"{safe_stem}_通用分析_{report_date:%Y%m%d}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )
        return

    with map_tab:
        st.subheader("把当前 Excel 列映射到业务字段")
        st.write("系统会先按常见英文和中文列名自动建议。请核对数量和交期口径；不同文件可以使用不同映射。")
        columns = list(source.columns)
        mapping: dict[str, str | None] = {}
        mapping_columns = st.columns(3)
        profile_hash = hashlib.sha1(profile_upload.getvalue()).hexdigest()[:8] if profile_upload is not None else "auto"
        for index, field in enumerate(FIELD_CONFIG):
            saved_column = saved_mapping.get(field["key"])
            suggested = saved_column if saved_column in columns else suggest_column(columns, field["aliases"])
            choices = [NO_COLUMN, *columns]
            selected_index = choices.index(suggested) if suggested in choices else 0
            widget_key = f"map_{file_hash}_{profile_hash}_{sheet_name}_{int(header_row_excel)}_{field['key']}"
            with mapping_columns[index % 3]:
                selected = st.selectbox(
                    field["label"],
                    choices,
                    index=selected_index,
                    key=widget_key,
                    help=field["help"],
                )
                mapping[field["key"]] = None if selected == NO_COLUMN else selected

        analysis, quantity_note, quality_stats = build_analysis_frame(source, mapping, int(header_row_excel))
        if analysis is None:
            st.warning(quantity_note)
        else:
            st.success(quantity_note)
            st.caption("数量字段按各自原始口径单独展示。若源系统的预留、已拣、欠交数量彼此重叠，不应将其相加。")
        with st.expander("预览当前工作表"):
            st.dataframe(source.head(20), use_container_width=True, hide_index=True)

        profile = {
            "template": "customer_order_delivery",
            "source_file": uploaded.name,
            "sheet_name": sheet_name,
            "header_row_excel": int(header_row_excel),
            "mapping": mapping,
        }
        st.download_button(
            "下载字段映射配置 JSON",
            data=json.dumps(profile, ensure_ascii=False, indent=2).encode("utf-8"),
            file_name=f"{Path(uploaded.name).stem}_字段映射.json",
            mime="application/json",
            help="保存后可以作为下一步自动加载映射配置的基础。",
        )

    if analysis is None:
        with dashboard_tab:
            st.info("先在“字段映射”中选择未结数量，或同时选择原始订单数量和已发货数量。")
        with ppt_tab:
            st.info("字段映射完成后，才能生成有数据的 PPT 报告。")
        return

    analysis["未结数量"] = pd.to_numeric(analysis["未结数量"], errors="coerce")
    analysis["承诺交期"] = pd.to_datetime(analysis["承诺交期"], errors="coerce")
    as_of = pd.Timestamp(report_date)
    open_mask = analysis["未结数量"].fillna(0) > 0
    due = analysis["承诺交期"]
    analysis["状态"] = "未到期"
    missing_qty_mask = analysis["未结数量"].isna()
    analysis.loc[missing_qty_mask, "状态"] = "缺少数量"
    analysis.loc[~missing_qty_mask & (analysis["未结数量"] <= 0), "状态"] = "无未结数量"
    analysis.loc[open_mask & due.isna(), "状态"] = "缺少交期"
    analysis.loc[open_mask & due.notna() & (due < as_of), "状态"] = "逾期"
    analysis.loc[open_mask & due.notna() & (due >= as_of) & (due <= as_of + pd.Timedelta(days=7)), "状态"] = "7天内到期"
    analysis["状态优先级"] = analysis["状态"].map({"逾期": 0, "7天内到期": 1, "缺少交期": 2, "缺少数量": 3, "未到期": 4, "无未结数量": 5}).fillna(6)

    filtered = analysis.copy()
    with dashboard_tab:
        filter_cols = st.columns([1.25, 1.25, 1.25, 1])
        customer_values = sorted(value for value in filtered["客户"].dropna().astype(str).unique() if value.strip())
        customer_choice = filter_cols[0].selectbox("客户", ["全部客户", *customer_values])
        status_values = ["逾期"]
        status_choice = filter_cols[1].multiselect("交付状态", status_values, default=["逾期"])
        carrier_values = sorted(value for value in filtered["承运商"].dropna().astype(str).unique() if value.strip())
        carrier_choice = filter_cols[2].selectbox("承运商", ["全部承运商", *carrier_values])
        search = filter_cols[3].text_input("订单 / 料号搜索")

        if customer_choice != "全部客户":
            filtered = filtered[filtered["客户"].astype(str) == customer_choice]
        if status_choice:
            filtered = filtered[filtered["状态"].isin(status_choice)]
        else:
            filtered = filtered.iloc[0:0]
        if carrier_choice != "全部承运商":
            filtered = filtered[filtered["承运商"].astype(str) == carrier_choice]
        if search.strip():
            term = search.strip().casefold()
            searchable = filtered[["订单号", "订单行", "物料号", "物料描述", "客户"]].astype(str).agg(" ".join, axis=1).str.casefold()
            filtered = filtered[searchable.str.contains(re.escape(term), na=False)]

        st.caption(f"当前筛选：{len(filtered):,} 行　|　报告截至 {report_date:%Y-%m-%d}")
        overdue_mask = filtered["状态"] == "逾期"
        next7_mask = filtered["状态"] == "7天内到期"
        metric_cols = st.columns(4)
        metric_cols[0].metric("未结订单行", f"{int((filtered['未结数量'].fillna(0) > 0).sum()):,}")
        metric_cols[1].metric("未结数量", format_number(filtered["未结数量"].fillna(0).sum()))
        metric_cols[2].metric("逾期未结数量", format_number(filtered.loc[overdue_mask, "未结数量"].fillna(0).sum()))
        metric_cols[3].metric("7天内到期数量", format_number(filtered.loc[next7_mask, "未结数量"].fillna(0).sum()))
        if mapping.get("backorder_qty"):
            st.caption(f"欠交数量（独立统计）：{format_number(filtered['欠交数量'].sum(min_count=1))}")
        if mapping.get("open_value"):
            st.caption(f"未结金额（源字段）：{format_number(filtered['未结金额'].sum(min_count=1))}")

        chart_col, quality_col = st.columns([1.55, 1])
        with chart_col:
            st.subheader("按交期周查看未结数量")
            weekly = weekly_open_qty(filtered)
            if weekly.empty:
                st.info("没有可用于绘图的交期。请检查承诺交期字段映射。")
            else:
                chart_data = weekly.tail(16).rename("未结数量").reset_index()
                chart_data.columns = ["交期周", "未结数量"]
                fig = px.bar(chart_data, x="交期周", y="未结数量", color_discrete_sequence=["#b9232e"])
                fig.update_layout(margin=dict(l=10, r=10, t=15, b=10), height=330, xaxis_title="", yaxis_title="数量", showlegend=False)
                st.plotly_chart(fig, use_container_width=True)
        with quality_col:
            st.subheader("数据质量提示")
            st.write(f"当前表格行数：**{quality_stats.get('rows', 0):,}**")
            st.write(f"缺少交期：**{quality_stats.get('missing_due_date', 0):,}** 行")
            st.write(f"未结数量为空：**{quality_stats.get('missing_open_qty', 0):,}** 行")
            if quality_stats.get("duplicate_keys", 0):
                st.warning(f"可能重复的订单行键：{quality_stats['duplicate_keys']:,} 行。请核对订单号、行号和物料号的组合。")
            if quality_stats.get("quantity_mismatch", 0):
                st.warning(
                    f"{quality_stats['quantity_mismatch']:,} 行中，未结量不等于预留、已拣和欠交数量之和。"
                    "这些状态可能重叠，应用分别展示。"
                )
            if quality_stats.get("overfulfilled_rows", 0):
                st.warning(f"{quality_stats['overfulfilled_rows']:,} 行的已发货数量大于原始订单数量。")

        st.subheader("订单行待办")
        field_columns = [
            ("customer", "客户"),
            ("order", "订单号"),
            ("line", "订单行"),
            ("item", "物料号"),
            ("description", "物料描述"),
            ("due_date", "承诺交期"),
            ("actual_date", "实际日期"),
            ("open_qty", "未结数量"),
            ("backorder_qty", "欠交数量"),
            ("shipped_qty", "已发货数量"),
            ("carrier", "承运商"),
            ("destination", "目的地"),
        ]
        display_columns = ["状态"] + [
            analysis_column
            for field_key, analysis_column in field_columns
            if mapping.get(field_key) and analysis_column in filtered.columns
        ] + ["Excel行"]
        table = filtered.sort_values(
            ["状态优先级", "承诺交期", "未结数量"],
            ascending=[True, True, False],
            na_position="last",
        )[display_columns]
        header_names = {"状态": "Status", "Excel行": "Excel Row"}
        header_names.update(
            {
                analysis_column: mapping[field_key]
                for field_key, analysis_column in field_columns
                if mapping.get(field_key) and analysis_column in table.columns
            }
        )
        table = table.rename(columns=header_names)
        st.dataframe(table, use_container_width=True, hide_index=True)
        csv_bytes = table.to_csv(index=False).encode("utf-8-sig")
        st.download_button("下载当前筛选明细 CSV", data=csv_bytes, file_name="供应链待办明细.csv", mime="text/csv")

    with ppt_tab:
        st.subheader("导出当前筛选结果")
        st.write("PPT 将包含履约概览、交期趋势、优先订单行和数据口径说明。PowerPoint 图表使用原生图表，可继续编辑。")
        st.caption(f"当前纳入 PPT 的订单行：{len(filtered):,}　|　报告日期：{report_date:%Y-%m-%d}")
        ppt_signature = hashlib.sha1(
            (
                file_hash
                + sheet_name
                + str(report_date)
                + json.dumps(mapping, sort_keys=True, ensure_ascii=False)
                + ",".join(filtered["Excel行"].astype(str))
            ).encode("utf-8")
        ).hexdigest()
        if st.button("生成 PPTX", type="primary"):
            try:
                st.session_state["ppt_bytes"] = build_pptx(filtered, uploaded.name, sheet_name, report_date, quantity_note, quality_stats)
                st.session_state["ppt_signature"] = ppt_signature
            except Exception as exc:
                st.error(f"生成 PPT 时出错：{exc}")
        if st.session_state.get("ppt_signature") == ppt_signature and st.session_state.get("ppt_bytes"):
            safe_stem = re.sub(r"[^\w\-]+", "_", Path(uploaded.name).stem).strip("_") or "supply_chain"
            st.download_button(
                "下载可编辑 PowerPoint",
                data=st.session_state["ppt_bytes"],
                file_name=f"{safe_stem}_供应链分析_{report_date:%Y%m%d}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )


if __name__ == "__main__":
    main()

