# 订单分析

一个可部署到 GitHub Pages 的浏览器端供应链订单分析应用。界面使用 HTML、CSS 和 JavaScript；Excel 文件在用户的浏览器中解析，不会由应用发送到服务器。

## GitHub Pages 部署

本仓库的网页入口位于根目录：`index.html`。

1. 打开仓库 **Settings → Pages**。
2. 在 **Build and deployment** 中选择 **Deploy from a branch**。
3. 选择 `main` 分支和 `/(root)` 文件夹，然后保存。
4. 等 Pages 完成发布后访问：<https://xiangjianli155-byte.github.io/orderplatform/>。

提交到 `main` 的更新会自动触发 Pages 重新发布。

## 功能

- 上传 `.xlsx`、`.xls`、`.xlsm`、`.xlsb` 或 `.csv`，并选择工作表和表头行。
- 自动建议常见客户、订单、料号、数量和交期字段；支持手动映射及 JSON 映射配置导入/导出。
- 订单模式按报告日期计算未结订单状态；交付状态筛选当前只提供“逾期”。
- 按客户、承运商和订单/料号搜索；查看未结量、逾期未结量、交期周趋势和数据质量检查。
- 订单待办表沿用 Excel 原始列名，附带应用生成的 `Status` 和 `Excel Row`。
- 通用模式支持按分类字段汇总数字列，可求和、平均或计数。
- 下载筛选明细和分类汇总 CSV；生成含可编辑原生图表的 PowerPoint 报告。

网页使用 SheetJS、Plotly.js 和 PptxGenJS 浏览器版本，这些库由 CDN 加载；使用时需要网络连接。应用不重算 Excel 公式，也不会执行宏。若工作簿单元格使用公式，请先在 Excel 中打开并保存，使公式结果可供读取。

## 本地运行 Python 版本（可选）

仓库仍保留原有 Streamlit 版本。需要 Python 3.10 或以上版本：

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m streamlit run app.py
```

若 PowerShell 策略不允许激活虚拟环境，可像上面一样直接使用 `.venv\Scripts\python.exe`。浏览器端 GitHub Pages 版本不需要安装 Python 依赖。

## 分析口径

- “未结数量”优先使用源表字段；若没有，则按“原始订单数量 - 已发货数量”计算，并将负值按 0 处理。
- 逾期条件为未结数量大于 0，且承诺/修订交期早于报告截至日期。
- 欠交、预留和已拣数量作为独立字段展示。若 ERP 口径未确认互不重叠，不应将它们相加或用于反推未结数量。
- 只有未结数量和交期时，可以识别逾期未结订单；要确认实际迟交或计算准时交付率，需要可靠的实际发运/交付数据。

