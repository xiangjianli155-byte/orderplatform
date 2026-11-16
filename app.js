/* Browser-only Excel analysis dashboard for GitHub Pages. */
"use strict";

const NO_COLUMN = "— 不使用 —";
const FIELD_CONFIG = [
  { key: "customer", label: "客户", help: "客户名称或客户号", aliases: ["customer name", "customer", "customer no", "客户名称", "客户", "客户编号"] },
  { key: "order", label: "订单 / PO", help: "订单号或客户采购订单号", aliases: ["order no", "order number", "sales order", "customer po", "po number", "订单号", "采购订单", "客户po"] },
  { key: "line", label: "订单行", help: "PO 行号或订单行号", aliases: ["customer po line", "po line", "order line", "line number", "行号", "订单行"] },
  { key: "item", label: "物料号", help: "料号、SKU 或 Item", aliases: ["part no", "part number", "item number", "sku", "material", "料号", "物料号", "产品编号"] },
  { key: "description", label: "物料描述", help: "产品名称或描述", aliases: ["description", "item name", "part description", "产品描述", "物料描述", "品名"] },
  { key: "ordered_qty", label: "原始订单数量", help: "用于计算未结量的备选字段", aliases: ["order qty", "ordered qty", "qty ordered", "订单数量", "订购数量"] },
  { key: "open_qty", label: "未结数量", help: "优先使用源表已有的 Qty Open", aliases: ["qty open", "open qty", "open quantity", "remaining qty", "未结数量", "未交数量", "开放数量"] },
  { key: "shipped_qty", label: "已发货数量", help: "与原始订单数量搭配计算未结量", aliases: ["qty shipped", "shipped qty", "shipped quantity", "已发货数量", "发货数量", "出货数量"] },
  { key: "backorder_qty", label: "欠交数量", help: "作为独立指标展示", aliases: ["qty backordered", "backorder qty", "backordered qty", "欠交数量", "缺货数量", "欠货数量"] },
  { key: "reserved_qty", label: "预留数量", help: "作为独立指标展示", aliases: ["qty reserved", "qty resvered", "reserved qty", "预留数量", "分配数量"] },
  { key: "picked_qty", label: "已拣数量", help: "作为独立指标展示", aliases: ["qty picked", "picked qty", "已拣数量", "拣货数量"] },
  { key: "due_date", label: "承诺 / 修订交期", help: "用于逾期和近期到期提醒", aliases: ["revised due date", "due date", "promise date", "requested date", "承诺交期", "修订交期", "交期", "需求日期"] },
  { key: "actual_date", label: "实际发货 / 交付日期", help: "用于明细查看；计算准时率前需确认数据粒度", aliases: ["actual delivery date", "actual ship date", "ship date", "delivery date", "实际交付日期", "实际发货日期", "出货日期"] },
  { key: "open_value", label: "未结金额", help: "若源表有金额，可汇总金额风险", aliases: ["total value open", "open value", "open amount", "未结金额", "未交金额"] },
  { key: "carrier", label: "承运商", help: "可选筛选维度", aliases: ["carrier code", "carrier", "承运商", "运输商"] },
  { key: "destination", label: "目的地", help: "可选筛选维度", aliases: ["dest location", "destination", "ship to", "目的地", "收货地点"] },
];

const state = {
  workbook: null,
  fileName: "",
  sheetName: "",
  sheetRows: [],
  headers: [],
  sourceRows: [],
  headerRow: 1,
  reportDate: "",
  mode: "orders",
  mapping: {},
  profileMapping: {},
  orderRows: [],
  quantityNote: "",
  quality: {},
  orderFiltered: [],
  genericFiltered: [],
  genericGrouped: [],
  genericDimension: null,
  genericMeasure: null,
  genericAggregation: "求和",
};

const $ = (id) => document.getElementById(id);
const localToday = () => {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;
};

function setMessage(id, text, type = "") {
  const box = $(id);
  box.textContent = text;
  box.className = `message${type ? ` ${type}` : ""}`;
  box.classList.toggle("hidden", !text);
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (character) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  })[character]);
}

function normalizeName(value) {
  return String(value ?? "").trim().toLocaleLowerCase().replace(/[\s_.\-/\\()（）]+/g, "").replace(/[\[\]]/g, "");
}

function suggestColumn(columns, aliases) {
  const normalizedColumns = columns.map((name) => [name, normalizeName(name)]);
  const normalizedAliases = aliases.map(normalizeName);
  for (const alias of normalizedAliases) {
    const match = normalizedColumns.find(([, normalized]) => normalized === alias);
    if (match) return match[0];
  }
  for (const alias of normalizedAliases) {
    if (alias.length < 4) continue;
    const match = normalizedColumns.find(([, normalized]) => normalized.includes(alias) || alias.includes(normalized));
    if (match) return match[0];
  }
  return null;
}

function uniqueColumnNames(values) {
  const counts = new Map();
  return values.map((value, index) => {
    let base = value == null ? "" : String(value).trim();
    if (!base || base.toLocaleLowerCase().startsWith("unnamed:")) base = `列 ${index + 1}`;
    const count = (counts.get(base) || 0) + 1;
    counts.set(base, count);
    return count === 1 ? base : `${base} (${count})`;
  });
}

function numericValue(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (value == null || value instanceof Date) return null;
  let text = String(value).trim();
  if (!text) return null;
  const isNegative = /^\(.*\)$/.test(text);
  text = text.replace(/^\((.*)\)$/, "$1").replace(/[,$¥￥€£\s]/g, "");
  const result = Number(text);
  if (!Number.isFinite(result)) return null;
  return isNegative ? -result : result;
}

function dateValue(value) {
  if (value == null || value === "") return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return formatLocalDate(value);
  if (typeof value === "number" && value >= 20000 && value <= 80000 && window.XLSX?.SSF) {
    const parts = XLSX.SSF.parse_date_code(value);
    if (parts) return `${parts.y}-${String(parts.m).padStart(2, "0")}-${String(parts.d).padStart(2, "0")}`;
  }
  const text = String(value).trim();
  const iso = text.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (iso) return `${iso[1]}-${String(iso[2]).padStart(2, "0")}-${String(iso[3]).padStart(2, "0")}`;
  const parsed = new Date(text);
  if (Number.isNaN(parsed.getTime())) return null;
  return formatLocalDate(parsed);
}

function formatLocalDate(value) {
  return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, "0")}-${String(value.getDate()).padStart(2, "0")}`;
}

function formatNumber(value) {
  if (value == null || !Number.isFinite(Number(value))) return "—";
  const number = Number(value);
  return new Intl.NumberFormat("en-US", { maximumFractionDigits: Number.isInteger(number) ? 0 : 2 }).format(number);
}

function guessHeaderRow(rows) {
  if (!rows.length) return 1;
  let bestIndex = 0;
  let bestScore = -1;
  rows.slice(0, 30).forEach((row, index) => {
    const nonEmpty = row.filter((value) => value != null && String(value).trim() !== "").length;
    const textCells = row.filter((value) => typeof value === "string" && value.trim()).length;
    const score = nonEmpty + 0.2 * textCells;
    if (score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
  });
  return bestIndex + 1;
}

function readCurrentTable() {
  const headerIndex = Math.max(0, Math.min(state.sheetRows.length - 1, Number(state.headerRow) - 1));
  const headerCells = state.sheetRows[headerIndex] || [];
  state.headers = uniqueColumnNames(headerCells);
  state.sourceRows = state.sheetRows.slice(headerIndex + 1).map((cells, index) => {
    const record = { __excelRow: headerIndex + index + 2 };
    state.headers.forEach((header, columnIndex) => { record[header] = cells?.[columnIndex] ?? null; });
    return record;
  }).filter((record) => state.headers.some((header) => record[header] != null && String(record[header]).trim() !== ""));
  state.mapping = {};
  FIELD_CONFIG.forEach((field) => {
    const saved = state.profileMapping?.[field.key];
    state.mapping[field.key] = state.headers.includes(saved) ? saved : suggestColumn(state.headers, field.aliases);
  });
  renderWorkspace();
}

async function loadWorkbook(file) {
  if (!window.XLSX) throw new Error("Excel parser library did not load. Check your internet connection and reload the page.");
  setMessage("loadMessage", "正在读取工作簿…");
  try {
    const bytes = await file.arrayBuffer();
    const workbook = XLSX.read(bytes, { cellDates: false, dense: false });
    if (!workbook.SheetNames?.length) throw new Error("The workbook has no readable worksheets.");
    state.workbook = workbook;
    state.fileName = file.name;
    state.sheetName = workbook.SheetNames[0];
    state.profileMapping = {};
    $("fileSummary").textContent = file.name;
    $("fileSummary").classList.remove("hidden");
    const sheetOptions = workbook.SheetNames.map((name) => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`).join("");
    $("sheetSelect").innerHTML = sheetOptions;
    $("sheetSelect").value = state.sheetName;
    $("workspace").classList.remove("hidden");
    $("uploadCard").classList.add("hidden");
    $("reportDateInput").value = state.reportDate || localToday();
    setMessage("loadMessage", "");
    loadCurrentSheet(true);
  } catch (error) {
    state.workbook = null;
    $("workspace").classList.add("hidden");
    $("uploadCard").classList.remove("hidden");
    setMessage("loadMessage", `无法读取工作簿：${error.message || error}`, "error");
  }
}

function loadCurrentSheet(guessHeader = false) {
  const sheet = state.workbook?.Sheets[state.sheetName];
  if (!sheet) return;
  state.sheetRows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null, blankrows: true });
  if (guessHeader) state.headerRow = guessHeaderRow(state.sheetRows);
  const maxHeader = Math.max(30, Math.min(100, state.sheetRows.length || 30));
  $("headerRowInput").max = String(maxHeader);
  $("headerRowInput").value = String(Math.min(state.headerRow, maxHeader));
  readCurrentTable();
}

function getMappingFromUi() {
  const mapping = {};
  FIELD_CONFIG.forEach(({ key }) => {
    const select = document.getElementById(`map-${key}`);
    mapping[key] = select && select.value !== NO_COLUMN ? select.value : null;
  });
  state.mapping = mapping;
  return mapping;
}

function renderMappingPanel() {
  const ordersMode = state.mode === "orders";
  $("orderMappingPanel").classList.toggle("hidden", !ordersMode);
  $("downloadMappingButton").classList.toggle("hidden", !ordersMode);
  $("genericSchemaPanel").classList.toggle("hidden", ordersMode);
  $("profileField").classList.toggle("hidden", !ordersMode);
  $("mappingTitle").textContent = ordersMode ? "核对订单分析字段" : "当前工作表结构";
  $("mappingDescription").textContent = ordersMode
    ? "核对系统建议的业务字段。列名可以不同，每份工作簿都可以单独映射。"
    : "通用模式不要求固定字段；所有列都可以作为分类维度，能识别为数字的列可以汇总。";
  if (ordersMode) {
    $("orderMappingPanel").innerHTML = FIELD_CONFIG.map((field) => {
      const candidate = state.mapping[field.key];
      const options = [NO_COLUMN, ...state.headers];
      const selected = options.includes(candidate) ? candidate : NO_COLUMN;
      return `<div class="mapping-item"><label for="map-${field.key}">${escapeHtml(field.label)}</label><select id="map-${field.key}">${options.map((option) => `<option value="${escapeHtml(option)}"${option === selected ? " selected" : ""}>${escapeHtml(option)}</option>`).join("")}</select><small>${escapeHtml(field.help)}</small></div>`;
    }).join("");
    FIELD_CONFIG.forEach(({ key }) => {
      document.getElementById(`map-${key}`).addEventListener("change", () => {
        getMappingFromUi();
        updateOrderAnalysis();
      });
    });
    updateOrderAnalysis();
  } else {
    const columns = state.headers.map((header) => {
      const values = state.sourceRows.map((row) => row[header]);
      const nonEmpty = values.filter((value) => value != null && String(value).trim() !== "").length;
      const unique = new Set(values.filter((value) => value != null).map((value) => String(value))).size;
      const numeric = values.filter((value) => numericValue(value) != null).length;
      return `<article class="schema-card"><strong>${escapeHtml(header)}</strong><span>非空行：${formatNumber(nonEmpty)}<br>唯一值：${formatNumber(unique)}<br>可识别数值：${formatNumber(numeric)}</span></article>`;
    }).join("");
    $("genericSchemaPanel").innerHTML = `<div class="schema-grid">${columns || "<div class='empty-table'>没有可显示的列</div>"}</div>`;
    renderSourcePreview();
  }
}

function renderSourcePreview() {
  const rows = state.sheetRows.slice(Math.max(0, state.headerRow - 1), Math.max(0, state.headerRow - 1) + 30);
  const headers = rows.map((row) => row.map((cell) => cell == null ? "" : String(cell)));
  const maxColumns = Math.min(20, Math.max(0, ...headers.map((row) => row.length)));
  if (!headers.length || !maxColumns) {
    $("sourcePreview").innerHTML = `<div class="empty-table">没有预览数据</div>`;
    return;
  }
  const previewHeaders = uniqueColumnNames(headers[0]);
  const previewRows = headers.slice(1).filter((row) => row.some((cell) => cell !== "")).map((row) => previewHeaders.slice(0, maxColumns).map((_, index) => row[index] ?? ""));
  $("sourcePreview").innerHTML = buildHtmlTable(previewHeaders.slice(0, maxColumns), previewRows.slice(0, 29), 29);
}

function buildOrderAnalysis() {
  const mapping = getMappingFromUi();
  const openColumn = mapping.open_qty;
  const orderedColumn = mapping.ordered_qty;
  const shippedColumn = mapping.shipped_qty;
  if (!openColumn && !(orderedColumn && shippedColumn)) {
    return { rows: [], note: "请映射未结数量，或同时映射原始订单数量和已发货数量。", stats: {} };
  }
  const asOf = $("reportDateInput").value || localToday();
  const rows = state.sourceRows.map((source) => {
    let openQty;
    let quantityNote;
    if (openColumn) {
      openQty = numericValue(source[openColumn]);
      quantityNote = `使用源字段“${openColumn}”作为未结数量。`;
    } else {
      const ordered = numericValue(source[orderedColumn]);
      const shipped = numericValue(source[shippedColumn]);
      openQty = ordered == null || shipped == null ? null : Math.max(0, ordered - shipped);
      quantityNote = `按“${orderedColumn} - ${shippedColumn}”计算未结数量，并将负值按 0 计。`;
    }
    const dueDate = mapping.due_date ? dateValue(source[mapping.due_date]) : null;
    const actualDate = mapping.actual_date ? dateValue(source[mapping.actual_date]) : null;
    const row = {
      source,
      excelRow: source.__excelRow,
      customer: mapping.customer ? String(source[mapping.customer] ?? "未指定") : "未指定",
      order: mapping.order ? String(source[mapping.order] ?? "") : "",
      line: mapping.line ? String(source[mapping.line] ?? "") : "",
      item: mapping.item ? String(source[mapping.item] ?? "") : "",
      description: mapping.description ? String(source[mapping.description] ?? "") : "",
      dueDate,
      actualDate,
      openQty,
      backorderQty: mapping.backorder_qty ? numericValue(source[mapping.backorder_qty]) : null,
      reservedQty: mapping.reserved_qty ? numericValue(source[mapping.reserved_qty]) : null,
      pickedQty: mapping.picked_qty ? numericValue(source[mapping.picked_qty]) : null,
      shippedQty: mapping.shipped_qty ? numericValue(source[mapping.shipped_qty]) : null,
      openValue: mapping.open_value ? numericValue(source[mapping.open_value]) : null,
      carrier: mapping.carrier ? String(source[mapping.carrier] ?? "") : "",
      destination: mapping.destination ? String(source[mapping.destination] ?? "") : "",
      status: "未到期",
    };
    if (openQty == null) row.status = "缺少数量";
    else if (openQty <= 0) row.status = "无未结数量";
    else if (!dueDate) row.status = "缺少交期";
    else if (dueDate < asOf) row.status = "逾期";
    else if (dueDate <= addDays(asOf, 7)) row.status = "7天内到期";
    row.priority = ({ "逾期": 0, "7天内到期": 1, "缺少交期": 2, "缺少数量": 3, "未到期": 4, "无未结数量": 5 })[row.status] ?? 6;
    row.quantityNote = quantityNote;
    return row;
  });
  const stats = {
    rows: rows.length,
    missing_open_qty: rows.filter((row) => row.openQty == null).length,
    missing_due_date: rows.filter((row) => !row.dueDate).length,
    negative_open_qty: rows.filter((row) => row.openQty != null && row.openQty < 0).length,
    overfulfilled_rows: 0,
    duplicate_keys: 0,
    quantity_mismatch: 0,
  };
  if (orderedColumn && shippedColumn) {
    stats.overfulfilled_rows = rows.filter((row) => {
      const ordered = numericValue(row.source[orderedColumn]);
      const shipped = numericValue(row.source[shippedColumn]);
      return ordered != null && shipped != null && shipped > ordered;
    }).length;
  }
  const keyFields = ["order", "line", "item"].filter((key) => mapping[key]);
  if (keyFields.length >= 2) {
    const keyCount = new Map();
    rows.forEach((row) => {
      const values = keyFields.map((key) => String(row[key] ?? "").trim());
      if (values.every(Boolean)) {
        const key = JSON.stringify(values);
        keyCount.set(key, (keyCount.get(key) || 0) + 1);
      }
    });
    rows.forEach((row) => {
      const values = keyFields.map((key) => String(row[key] ?? "").trim());
      if (values.every(Boolean) && keyCount.get(JSON.stringify(values)) > 1) stats.duplicate_keys += 1;
    });
  }
  if (openColumn && mapping.backorder_qty && mapping.reserved_qty && mapping.picked_qty) {
    stats.quantity_mismatch = rows.filter((row) => row.openQty != null && row.backorderQty != null && row.reservedQty != null && row.pickedQty != null
      && Math.abs(row.openQty - row.backorderQty - row.reservedQty - row.pickedQty) > 1e-9).length;
  }
  return { rows, note: rows[0]?.quantityNote || (openColumn ? `使用源字段“${openColumn}”作为未结数量。` : `按“${orderedColumn} - ${shippedColumn}”计算未结数量，并将负值按 0 计。`), stats };
}

function addDays(dateText, days) {
  if (!dateText) return "";
  const [year, month, day] = dateText.split("-").map(Number);
  const date = new Date(Date.UTC(year, month - 1, day + days));
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-${String(date.getUTCDate()).padStart(2, "0")}`;
}

function updateOrderAnalysis() {
  if (state.mode !== "orders") return;
  const built = buildOrderAnalysis();
  state.orderRows = built.rows;
  state.quantityNote = built.note;
  state.quality = built.stats;
  const valid = Boolean(state.orderRows.length && state.mapping.open_qty || state.orderRows.length && state.mapping.ordered_qty && state.mapping.shipped_qty);
  if (!valid) {
    setMessage("analysisNotice", built.note, "error");
    state.orderFiltered = [];
    renderOrderDashboard();
    return;
  }
  setMessage("analysisNotice", built.note, "success");
  renderOrderDashboard();
}

function orderFilterRows() {
  const customer = $("customerFilter").value || "__all__";
  const carrier = $("carrierFilter").value || "__all__";
  const search = $("orderSearch").value.trim().toLocaleLowerCase();
  return state.orderRows.filter((row) => {
    if (row.status !== "逾期") return false;
    if (customer !== "__all__" && row.customer !== customer) return false;
    if (carrier !== "__all__" && row.carrier !== carrier) return false;
    if (search) {
      const searchable = [row.order, row.line, row.item, row.description, row.customer].join(" ").toLocaleLowerCase();
      if (!searchable.includes(search)) return false;
    }
    return true;
  }).sort((a, b) => (a.dueDate || "9999").localeCompare(b.dueDate || "9999") || (b.openQty || 0) - (a.openQty || 0));
}

function renderOrderDashboard() {
  $("orderDashboard").classList.toggle("hidden", state.mode !== "orders");
  $("genericDashboard").classList.toggle("hidden", state.mode !== "generic");
  if (state.mode !== "orders") return;
  const customerOptions = [...new Set(state.orderRows.map((row) => row.customer).filter((value) => value && value.trim()))].sort((a, b) => a.localeCompare(b));
  const carrierOptions = [...new Set(state.orderRows.map((row) => row.carrier).filter((value) => value && value.trim()))].sort((a, b) => a.localeCompare(b));
  updateSelectOptions($("customerFilter"), [["__all__", "全部客户"], ...customerOptions.map((value) => [value, value])], "__all__");
  updateSelectOptions($("carrierFilter"), [["__all__", "全部承运商"], ...carrierOptions.map((value) => [value, value])], "__all__");
  $("dashboardCaption").textContent = `当前筛选 ${state.orderRows.filter((row) => row.status === "逾期").length.toLocaleString()} 条逾期行　|　报告截至 ${$("reportDateInput").value || localToday()}`;
  state.orderFiltered = orderFilterRows();
  const filtered = state.orderFiltered;
  const openRows = filtered.filter((row) => row.openQty != null && row.openQty > 0);
  const openQty = openRows.reduce((sum, row) => sum + row.openQty, 0);
  const overdueQty = filtered.filter((row) => row.status === "逾期").reduce((sum, row) => sum + (row.openQty || 0), 0);
  const dueSoonQty = filtered.filter((row) => row.status === "7天内到期").reduce((sum, row) => sum + (row.openQty || 0), 0);
  $("metricOpenLines").textContent = formatNumber(openRows.length);
  $("metricOpenQty").textContent = formatNumber(openQty);
  $("metricOverdueQty").textContent = formatNumber(overdueQty);
  $("metricDueSoonQty").textContent = formatNumber(dueSoonQty);
  $("tableResultCount").textContent = `${filtered.length.toLocaleString()} 条记录`;
  renderOrderChart(filtered);
  renderQualitySummary();
  const mappedFields = [
    ["customer", "customer"], ["order", "order"], ["line", "line"], ["item", "item"], ["description", "description"],
    ["due_date", "dueDate"], ["actual_date", "actualDate"], ["open_qty", "openQty"], ["backorder_qty", "backorderQty"],
    ["shipped_qty", "shippedQty"], ["carrier", "carrier"], ["destination", "destination"],
  ].filter(([field]) => state.mapping[field]);
  const columns = [{ name: "Status", get: (row) => row.status }, ...mappedFields.map(([field, property]) => ({
    name: state.mapping[field],
    get: (row) => property === "customer" || property === "order" || property === "line" || property === "item" || property === "description" || property === "carrier"
      ? row[property]
      : property === "destination" ? row.destination
        : property === "dueDate" || property === "actualDate" ? row[property] || ""
          : row[property] == null ? "" : formatNumber(row[property]),
  })), { name: "Excel Row", get: (row) => row.excelRow }];
  renderDataTable($("orderTable"), columns.map((column) => column.name), filtered.map((row) => columns.map((column) => column.get(row))), 500);
}

function updateSelectOptions(select, options, preferredValue) {
  const current = select.value || preferredValue;
  select.innerHTML = options.map(([value, label]) => `<option value="${escapeHtml(value)}">${escapeHtml(label)}</option>`).join("");
  const values = options.map(([value]) => value);
  select.value = values.includes(current) ? current : preferredValue;
}

function renderOrderChart(rows) {
  if (!window.Plotly) {
    $("orderChart").innerHTML = `<div class="empty-table">Chart library did not load. Check your internet connection.</div>`;
    return;
  }
  const byWeek = new Map();
  rows.forEach((row) => {
    if (!row.dueDate) return;
    const monday = weekStart(row.dueDate);
    byWeek.set(monday, (byWeek.get(monday) || 0) + (row.openQty || 0));
  });
  const weeks = [...byWeek.keys()].sort().slice(-16);
  if (!weeks.length) {
    Plotly.purge("orderChart");
    $("orderChart").innerHTML = `<div class="empty-table">没有可用于绘图的交期。</div>`;
    return;
  }
  Plotly.react("orderChart", [{ x: weeks, y: weeks.map((week) => byWeek.get(week)), type: "bar", marker: { color: "#b9232e" }, hovertemplate: "%{x}<br>未结数量：%{y:,.0f}<extra></extra>" }], {
    margin: { l: 54, r: 12, t: 12, b: 58 }, height: 300, paper_bgcolor: "#ffffff", plot_bgcolor: "#ffffff",
    font: { family: "Segoe UI, Microsoft YaHei, sans-serif", size: 11, color: "#47515d" },
    xaxis: { title: "", tickangle: -35, gridcolor: "#f0f2f5" }, yaxis: { title: "数量", rangemode: "tozero", gridcolor: "#e9edf1" },
    showlegend: false,
  }, { responsive: true, displayModeBar: false });
}

function weekStart(dateText) {
  const [year, month, day] = dateText.split("-").map(Number);
  const date = new Date(Date.UTC(year, month - 1, day));
  const dayIndex = (date.getUTCDay() + 6) % 7;
  date.setUTCDate(date.getUTCDate() - dayIndex);
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-${String(date.getUTCDate()).padStart(2, "0")}`;
}

function renderQualitySummary() {
  const stats = state.quality || {};
  const entries = [
    ["当前工作表行数", stats.rows || 0, false],
    ["缺少交期", stats.missing_due_date || 0, (stats.missing_due_date || 0) > 0],
    ["未结数量为空", stats.missing_open_qty || 0, (stats.missing_open_qty || 0) > 0],
    ["可能重复的订单行", stats.duplicate_keys || 0, (stats.duplicate_keys || 0) > 0],
  ];
  let content = entries.map(([label, value, warning]) => `<div class="quality-row${warning ? " warning" : ""}"><span>${escapeHtml(label)}</span><strong>${formatNumber(value)}</strong></div>`).join("");
  if ((stats.quantity_mismatch || 0) > 0) content += `<div class="quality-note">${formatNumber(stats.quantity_mismatch)} 行中，Qty Open 与预留、已拣、欠交数量之和不一致。这些数量可能重叠，已分别统计。</div>`;
  if ((stats.overfulfilled_rows || 0) > 0) content += `<div class="quality-note">${formatNumber(stats.overfulfilled_rows)} 行的已发货数量大于原始订单数量，请核对源数据。</div>`;
  if (state.mapping.backorder_qty) content += `<div class="quality-row"><span>欠交数量（独立指标）</span><strong>${formatNumber(state.orderFiltered.reduce((sum, row) => sum + (row.backorderQty || 0), 0))}</strong></div>`;
  if (state.mapping.open_value) content += `<div class="quality-row"><span>未结金额（源字段）</span><strong>${formatNumber(state.orderFiltered.reduce((sum, row) => sum + (row.openValue || 0), 0))}</strong></div>`;
  $("qualitySummary").innerHTML = content;
}

function renderDataTable(container, headers, rows, maxRows = Infinity) {
  if (!headers.length || !rows.length) {
    container.innerHTML = `<div class="empty-table">当前筛选范围内没有记录</div>`;
    return;
  }
  const limited = rows.slice(0, maxRows);
  const head = headers.map((name) => `<th title="${escapeHtml(name)}">${escapeHtml(name)}</th>`).join("");
  const body = limited.map((row) => `<tr>${row.map((value, index) => {
    const status = headers[index] === "Status" && value === "逾期";
    const content = status ? `<span class="status-pill">${escapeHtml(value)}</span>` : escapeHtml(value);
    return `<td title="${escapeHtml(value)}">${content}</td>`;
  }).join("")}</tr>`).join("");
  const extra = rows.length > limited.length ? `<caption>显示前 ${limited.length.toLocaleString()} 行；下载可获取全部筛选结果。</caption>` : "";
  container.innerHTML = `<table class="data-table">${extra}<thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>`;
}

function downloadCsv(headers, rows, filename) {
  const quote = (value) => `"${String(value ?? "").replace(/"/g, '""')}"`;
  const csv = `\uFEFF${[headers, ...rows].map((row) => row.map(quote).join(",")).join("\r\n")}`;
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function numericColumnMap(rows, headers) {
  const result = {};
  headers.forEach((header) => {
    const values = rows.map((row) => numericValue(row[header]));
    if (values.some((value) => value != null)) result[header] = values;
  });
  return result;
}

function getGenericSearchRows() {
  const search = $("genericSearch").value.trim().toLocaleLowerCase();
  if (!search) return [...state.sourceRows];
  return state.sourceRows.filter((row) => state.headers.map((header) => String(row[header] ?? "")).join(" ").toLocaleLowerCase().includes(search));
}

function setupGenericControls() {
  const numericMap = numericColumnMap(state.sourceRows, state.headers);
  const dimensionSelect = $("genericDimension");
  const measureSelect = $("genericMeasure");
  const dimensionOptions = [[NO_COLUMN, NO_COLUMN], ...state.headers.map((header) => [header, header])];
  const measureOptions = [[NO_COLUMN, NO_COLUMN], ...Object.keys(numericMap).map((header) => [header, header])];
  const previousDimension = state.genericDimension || state.headers[0] || NO_COLUMN;
  const previousMeasure = state.genericMeasure || Object.keys(numericMap)[0] || NO_COLUMN;
  updateSelectOptions(dimensionSelect, dimensionOptions, dimensionOptions.some(([value]) => value === previousDimension) ? previousDimension : NO_COLUMN);
  updateSelectOptions(measureSelect, measureOptions, measureOptions.some(([value]) => value === previousMeasure) ? previousMeasure : NO_COLUMN);
  dimensionSelect.onchange = () => { state.genericDimension = dimensionSelect.value === NO_COLUMN ? null : dimensionSelect.value; renderGenericDashboard(); };
  measureSelect.onchange = () => { state.genericMeasure = measureSelect.value === NO_COLUMN ? null : measureSelect.value; renderGenericDashboard(); };
}

function renderGenericDashboard() {
  if (state.mode !== "generic") return;
  const dimension = $("genericDimension").value === NO_COLUMN ? null : $("genericDimension").value;
  const measure = $("genericMeasure").value === NO_COLUMN ? null : $("genericMeasure").value;
  const aggregation = $("genericAggregation").value;
  const rows = getGenericSearchRows();
  state.genericFiltered = rows;
  const numericMap = numericColumnMap(state.sourceRows, state.headers);
  const groupMap = new Map();
  const values = measure ? numericMap[measure] || [] : [];
  let missingMeasure = 0;
  let missingDimension = 0;
  rows.forEach((row) => {
    const categoryValue = dimension ? row[dimension] : "全部数据";
    const category = categoryValue == null || String(categoryValue).trim() === "" ? "（空白）" : String(categoryValue);
    if (dimension && (categoryValue == null || String(categoryValue).trim() === "")) missingDimension += 1;
    const numeric = measure ? numericValue(row[measure]) : null;
    if (measure && numeric == null) missingMeasure += 1;
    if (!groupMap.has(category)) groupMap.set(category, { sum: 0, count: 0, rows: 0 });
    const bucket = groupMap.get(category);
    bucket.rows += 1;
    if (numeric != null) { bucket.sum += numeric; bucket.count += 1; }
  });
  const isCount = aggregation === "计数" || !measure;
  const grouped = [...groupMap.entries()].map(([category, bucket]) => ({
    category,
    value: isCount ? bucket.rows : aggregation === "平均值" ? (bucket.count ? bucket.sum / bucket.count : null) : bucket.count ? bucket.sum : null,
  })).filter((group) => group.value != null).sort((a, b) => b.value - a.value).slice(0, 20);
  state.genericGrouped = grouped;
  let metricValue;
  if (isCount) metricValue = rows.length;
  else {
    const selectedValues = rows.map((row) => numericValue(row[measure])).filter((value) => value != null);
    metricValue = aggregation === "平均值" ? (selectedValues.length ? selectedValues.reduce((sum, value) => sum + value, 0) / selectedValues.length : 0) : selectedValues.reduce((sum, value) => sum + value, 0);
  }
  $("genericRowsMetric").textContent = formatNumber(rows.length);
  $("genericCategoriesMetric").textContent = formatNumber(dimension ? new Set(rows.map((row) => row[dimension]).filter((value) => value != null).map(String)).size : 1);
  $("genericMetricLabel").textContent = isCount ? "记录数" : `${aggregation}：${measure}`;
  $("genericValueMetric").textContent = formatNumber(metricValue);
  $("genericTableCount").textContent = `${rows.length.toLocaleString()} 条筛选记录`;
  $("genericQuality").innerHTML = `<div class="quality-row"><span>当前筛选行数</span><strong>${formatNumber(rows.length)}</strong></div><div class="quality-row${missingDimension ? " warning" : ""}"><span>分类字段空白</span><strong>${formatNumber(missingDimension)}</strong></div><div class="quality-row${missingMeasure ? " warning" : ""}"><span>数值为空或非数字</span><strong>${formatNumber(missingMeasure)}</strong></div>`;
  renderGenericChart(grouped, dimension, isCount ? "记录数" : `${aggregation}：${measure}`);
  const detailHeaders = ["Excel行", ...state.headers];
  const detailRows = rows.slice(0, 500).map((row) => [row.__excelRow, ...state.headers.map((header) => row[header] == null ? "" : String(row[header]))]);
  renderDataTable($("genericTable"), detailHeaders, detailRows, 500);
  renderDataTable($("genericGroupedTable"), [dimension || "分类", isCount ? "记录数" : `${aggregation}：${measure}`], grouped.map((item) => [item.category, formatNumber(item.value)]));
}

function renderGenericChart(grouped, dimension, metricLabel) {
  if (!window.Plotly) return;
  if (!grouped.length) {
    Plotly.purge("genericChart");
    $("genericChart").innerHTML = `<div class="empty-table">当前筛选范围内没有可汇总的数据。</div>`;
    return;
  }
  Plotly.react("genericChart", [{ x: grouped.map((item) => item.category), y: grouped.map((item) => item.value), type: "bar", marker: { color: "#b9232e" }, hovertemplate: "%{x}<br>%{y:,.2f}<extra></extra>" }], {
    margin: { l: 58, r: 12, t: 12, b: 70 }, height: 310, paper_bgcolor: "#ffffff", plot_bgcolor: "#ffffff",
    font: { family: "Segoe UI, Microsoft YaHei, sans-serif", size: 11, color: "#47515d" },
    xaxis: { title: dimension || "", tickangle: -35, gridcolor: "#f0f2f5" }, yaxis: { title: metricLabel, rangemode: "tozero", gridcolor: "#e9edf1" }, showlegend: false,
  }, { responsive: true, displayModeBar: false });
}

function buildHtmlTable(headers, rows, maxRows = Infinity) {
  if (!headers.length || !rows.length) return `<div class="empty-table">没有预览数据</div>`;
  const limited = rows.slice(0, maxRows);
  return `<table class="data-table"><thead><tr>${headers.map((header) => `<th title="${escapeHtml(header)}">${escapeHtml(header)}</th>`).join("")}</tr></thead><tbody>${limited.map((row) => `<tr>${row.map((value) => `<td title="${escapeHtml(value)}">${escapeHtml(value)}</td>`).join("")}</tr>`).join("")}</tbody></table>`;
}

function setActiveView(viewId) {
  document.querySelectorAll(".tab-button").forEach((button) => button.classList.toggle("active", button.dataset.view === viewId));
  document.querySelectorAll(".view-panel").forEach((panel) => panel.classList.toggle("hidden", panel.id !== viewId));
  if (viewId === "dashboardView") {
    if (state.mode === "orders") renderOrderDashboard();
    else renderGenericDashboard();
  }
}

function renderWorkspace() {
  $("fileSummary").textContent = `${state.fileName}　|　${state.sheetName}　|　${state.sourceRows.length.toLocaleString()} 行`;
  $("fileSummary").classList.remove("hidden");
  renderSourcePreview();
  $("modeSelect").value = state.mode;
  renderMappingPanel();
  if (state.mode === "generic") {
    setupGenericControls();
    renderGenericDashboard();
  }
  renderReportSettings();
}

function renderReportSettings() {
  const ordersMode = state.mode === "orders";
  $("reportTitle").textContent = ordersMode ? "订单履约报告" : "通用分析报告";
  $("reportDescription").textContent = ordersMode
    ? "包含履约概览、交期趋势、优先订单行和数据质量说明。"
    : "包含数据概览、可编辑分类图表、前 20 个分类结果和缺失值检查。";
}

function createPptx() {
  if (typeof window.PptxGenJS !== "function") throw new Error("PowerPoint library did not load. Check your internet connection and reload the page.");
  const pptx = new window.PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "Order Analysis";
  pptx.subject = "Supply chain planning report";
  pptx.title = "订单分析报告";
  pptx.company = "";
  pptx.lang = "zh-CN";
  pptx.theme = { headFontFace: "Microsoft YaHei", bodyFontFace: "Microsoft YaHei", lang: "zh-CN" };
  return pptx;
}

function addPptBand(slide, pptx) {
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 0.17, line: { color: "B9232E", transparency: 100 }, fill: { color: "B9232E" } });
}

function addPptHeader(slide, pptx, title, subtitle = "") {
  addPptBand(slide, pptx);
  slide.addText(title, { x: 0.65, y: 0.42, w: 12, h: 0.45, fontFace: "Microsoft YaHei", fontSize: 20, bold: true, color: "202833", margin: 0 });
  if (subtitle) slide.addText(subtitle, { x: 0.67, y: 0.9, w: 12, h: 0.3, fontFace: "Microsoft YaHei", fontSize: 9, color: "66717F", margin: 0 });
}

function pptMetricCard(slide, pptx, x, label, value) {
  slide.addShape(pptx.ShapeType.roundRect, { x, y: 1.45, w: 2.85, h: 1.12, rectRadius: 0.08, line: { color: "E1E5EA", pt: 0.6 }, fill: { color: "F7F8FA" } });
  slide.addText(label, { x: x + 0.16, y: 1.64, w: 2.55, h: 0.22, fontFace: "Microsoft YaHei", fontSize: 9, color: "66717F", margin: 0 });
  slide.addText(value, { x: x + 0.16, y: 1.94, w: 2.55, h: 0.42, fontFace: "Microsoft YaHei", fontSize: 20, bold: true, color: "202833", margin: 0 });
}

function pptChart(slide, pptx, entries, opts = {}) {
  if (!entries.length) {
    slide.addText(opts.emptyText || "没有可用数据生成图表。", { x: 1, y: 2.8, w: 10.8, h: 0.5, fontFace: "Microsoft YaHei", fontSize: 14, color: "66717F", align: "center", margin: 0.1 });
    return;
  }
  slide.addChart(pptx.ChartType.bar, [{ name: opts.seriesName || "数量", labels: entries.map((item) => String(item.label)), values: entries.map((item) => Number(item.value) || 0) }], {
    x: 0.68, y: 1.45, w: 11.95, h: 5.45, catAxisLabelFontFace: "Microsoft YaHei", catAxisLabelFontSize: 9,
    valAxisLabelFontFace: "Microsoft YaHei", valAxisLabelFontSize: 9, catAxisLabelRotate: -35, showLegend: false,
    showTitle: false, chartColors: ["B9232E"], valGridLine: { color: "E5E8EC", width: 1 },
    showValue: false, ...opts.chartOptions,
  });
}

function pptTable(slide, rows, opts = {}) {
  if (!rows.length) return;
  const styled = rows.map((row, rowIndex) => row.map((value) => rowIndex === 0
    ? { text: String(value ?? ""), options: { bold: true, color: "35404C", fill: "F1F3F5" } }
    : String(value ?? "")));
  slide.addTable(styled, {
    x: 0.6, y: 1.45, w: 12.1, h: 5.4, border: { type: "solid", color: "D7DDE4", pt: 0.5 },
    fontFace: "Microsoft YaHei", fontSize: 8, color: "27313C", margin: 0.06, valign: "mid", rowH: 0.28, ...opts,
  });
}

function weeklyEntries(rows) {
  const byWeek = new Map();
  rows.forEach((row) => {
    if (!row.dueDate) return;
    const week = weekStart(row.dueDate);
    byWeek.set(week, (byWeek.get(week) || 0) + (row.openQty || 0));
  });
  return [...byWeek.entries()].sort(([a], [b]) => a.localeCompare(b)).slice(-16).map(([label, value]) => ({ label, value }));
}

function orderPptTableRows() {
  const fieldList = [
    ["customer", "客户"], ["order", "订单号"], ["line", "订单行"], ["item", "物料号"],
    ["due_date", "承诺交期"], ["open_qty", "未结数量"], ["backorder_qty", "欠交数量"],
  ].filter(([field]) => state.mapping[field]);
  const headers = fieldList.map(([field]) => state.mapping[field]);
  const rows = [...state.orderFiltered].sort((a, b) => a.priority - b.priority || (a.dueDate || "9999").localeCompare(b.dueDate || "9999") || (b.openQty || 0) - (a.openQty || 0)).slice(0, 12);
  return [headers, ...rows.map((row) => fieldList.map(([field]) => {
    if (field === "due_date") return row.dueDate || "未提供";
    if (field === "open_qty") return formatNumber(row.openQty);
    if (field === "backorder_qty") return formatNumber(row.backorderQty);
    const property = ({ customer: "customer", order: "order", line: "line", item: "item" })[field];
    return row[property] || "";
  }))];
}

async function generateOrderPpt() {
  const pptx = createPptx();
  const reportDate = $("reportDateInput").value || localToday();
  const sourceName = state.fileName;
  const sheetName = state.sheetName;
  const rows = state.orderFiltered;
  const openRows = rows.filter((row) => row.openQty != null && row.openQty > 0);
  const totalOpen = openRows.reduce((sum, row) => sum + row.openQty, 0);
  const overdueQty = rows.filter((row) => row.status === "逾期").reduce((sum, row) => sum + (row.openQty || 0), 0);
  const dueSoon = rows.filter((row) => row.status === "7天内到期").length;

  let slide = pptx.addSlide();
  addPptBand(slide, pptx);
  slide.addText("订单分析报告", { x: 0.8, y: 1.55, w: 11.8, h: 0.65, fontFace: "Microsoft YaHei", fontSize: 30, bold: true, color: "202833", margin: 0 });
  slide.addText("客户订单与交付", { x: 0.82, y: 2.3, w: 11.6, h: 0.42, fontFace: "Microsoft YaHei", fontSize: 19, bold: true, color: "B9232E", margin: 0 });
  slide.addText(`报告日期：${reportDate}`, { x: 0.84, y: 3.15, w: 11.4, h: 0.3, fontFace: "Microsoft YaHei", fontSize: 14, color: "27313C", margin: 0 });
  slide.addText(`来源：${sourceName}　|　工作表：${sheetName}`, { x: 0.84, y: 3.57, w: 11.4, h: 0.3, fontFace: "Microsoft YaHei", fontSize: 10, color: "66717F", margin: 0 });

  slide = pptx.addSlide();
  addPptHeader(slide, pptx, "订单履约概览", `报告截至 ${reportDate} · 当前筛选范围`);
  const metrics = [["未结订单行", formatNumber(openRows.length)], ["未结数量", formatNumber(totalOpen)], ["逾期未结行", formatNumber(rows.filter((row) => row.status === "逾期").length)], ["7天内到期行", formatNumber(dueSoon)]];
  metrics.forEach(([label, value], index) => pptMetricCard(slide, pptx, 0.55 + index * 3.15, label, value));
  slide.addText(`逾期未结数量：${formatNumber(overdueQty)}`, { x: 0.72, y: 2.95, w: 5.2, h: 0.3, fontFace: "Microsoft YaHei", fontSize: 14, bold: true, color: "B9232E", margin: 0 });
  slide.addText(state.quantityNote, { x: 0.72, y: 3.4, w: 11.8, h: 0.5, fontFace: "Microsoft YaHei", fontSize: 10, color: "66717F", margin: 0.02, breakLine: false });

  slide = pptx.addSlide();
  addPptHeader(slide, pptx, "未结数量按交期周分布", "柱形图为 PowerPoint 原生图表，可在 PowerPoint 中编辑");
  pptChart(slide, pptx, weeklyEntries(rows), { seriesName: "未结数量", emptyText: "没有可用的承诺交期，无法生成交期趋势图。" });

  slide = pptx.addSlide();
  addPptHeader(slide, pptx, "优先处理订单行", "按逾期状态、交期和未结数量排列");
  pptTable(slide, orderPptTableRows());

  slide = pptx.addSlide();
  addPptHeader(slide, pptx, "数据口径与质量检查", "用于说明本次分析的范围和可用字段");
  const checks = [
    `源数据行数：${formatNumber(state.quality.rows || 0)}`,
    `缺少未结数量：${formatNumber(state.quality.missing_open_qty || 0)} 行`,
    `缺少交期：${formatNumber(state.quality.missing_due_date || 0)} 行`,
    `重复订单行键：${formatNumber(state.quality.duplicate_keys || 0)} 行`,
    `未结量与预留 / 已拣 / 欠交合计不一致：${formatNumber(state.quality.quantity_mismatch || 0)} 行`,
  ];
  slide.addText("数量口径", { x: 0.72, y: 1.5, w: 2, h: 0.3, fontFace: "Microsoft YaHei", fontSize: 14, bold: true, color: "B9232E", margin: 0 });
  slide.addText(state.quantityNote, { x: 0.72, y: 1.9, w: 11.8, h: 0.55, fontFace: "Microsoft YaHei", fontSize: 11, color: "27313C", margin: 0.03 });
  slide.addText("检查结果", { x: 0.72, y: 2.65, w: 2, h: 0.3, fontFace: "Microsoft YaHei", fontSize: 14, bold: true, color: "B9232E", margin: 0 });
  slide.addText(checks.map((value) => `• ${value}`).join("\n"), { x: 0.82, y: 3.05, w: 11.2, h: 2.75, fontFace: "Microsoft YaHei", fontSize: 11, color: "27313C", breakLine: false, paraSpaceAfterPt: 11, margin: 0.03, valign: "top" });
  slide.addText(`来源文件：${sourceName}　|　工作表：${sheetName}`, { x: 0.72, y: 6.42, w: 11.8, h: 0.25, fontFace: "Microsoft YaHei", fontSize: 9, color: "66717F", margin: 0 });
  await pptx.writeFile({ fileName: `${safeFilename(sourceName)}_订单分析_${reportDate.replace(/-/g, "")}.pptx` });
}

function safeFilename(value) {
  return String(value || "order_analysis").replace(/\.[^.]+$/, "").replace(/[<>:"/\\|?*\s]+/g, "_").replace(/^_+|_+$/g, "") || "order_analysis";
}

async function generateGenericPpt() {
  const pptx = createPptx();
  const reportDate = $("reportDateInput").value || localToday();
  const dim = $("genericDimension").value === NO_COLUMN ? null : $("genericDimension").value;
  const measure = $("genericMeasure").value === NO_COLUMN ? null : $("genericMeasure").value;
  const aggregation = $("genericAggregation").value;
  const grouped = state.genericGrouped;
  const isCount = aggregation === "计数" || !measure;
  const metricLabel = isCount ? "记录数" : `${aggregation}：${measure}`;
  const metricValue = $("genericValueMetric").textContent;
  let slide = pptx.addSlide();
  addPptBand(slide, pptx);
  slide.addText("通用数据分析报告", { x: 0.8, y: 1.55, w: 11.8, h: 0.65, fontFace: "Microsoft YaHei", fontSize: 30, bold: true, color: "202833", margin: 0 });
  slide.addText(`分析维度：${dim || "全部数据"}`, { x: 0.82, y: 2.3, w: 11.6, h: 0.42, fontFace: "Microsoft YaHei", fontSize: 18, bold: true, color: "B9232E", margin: 0 });
  slide.addText(`报告日期：${reportDate}`, { x: 0.84, y: 3.15, w: 11.4, h: 0.3, fontFace: "Microsoft YaHei", fontSize: 14, color: "27313C", margin: 0 });
  slide.addText(`来源：${state.fileName}　|　工作表：${state.sheetName}`, { x: 0.84, y: 3.57, w: 11.4, h: 0.3, fontFace: "Microsoft YaHei", fontSize: 10, color: "66717F", margin: 0 });

  slide = pptx.addSlide();
  addPptHeader(slide, pptx, "数据概览", `汇总方式：${metricLabel}　|　分组字段：${dim || "全部数据"}`);
  pptMetricCard(slide, pptx, 0.75, "数据行数", formatNumber(state.genericFiltered.length));
  pptMetricCard(slide, pptx, 4.05, "分组数量", formatNumber(grouped.length));
  pptMetricCard(slide, pptx, 7.35, metricLabel, metricValue);
  pptChart(slide, pptx, grouped.map((item) => ({ label: item.category, value: item.value })), { seriesName: metricLabel, chartOptions: { barDir: "bar" } });

  slide = pptx.addSlide();
  addPptHeader(slide, pptx, "分组结果与数据检查", "展示当前分组下的前 20 个结果");
  const missingDimension = dim ? state.genericFiltered.filter((row) => row[dim] == null || String(row[dim]).trim() === "").length : 0;
  const missingMeasure = measure ? state.genericFiltered.filter((row) => numericValue(row[measure]) == null).length : 0;
  slide.addText(`缺少分组值：${formatNumber(missingDimension)} 行`, { x: 0.7, y: 1.12, w: 5.5, h: 0.25, fontFace: "Microsoft YaHei", fontSize: 9, color: "66717F", margin: 0 });
  slide.addText(`缺少数值：${formatNumber(missingMeasure)} 行`, { x: 6.1, y: 1.12, w: 5.5, h: 0.25, fontFace: "Microsoft YaHei", fontSize: 9, color: "66717F", margin: 0 });
  pptTable(slide, [[dim || "分类", metricLabel], ...grouped.map((item) => [item.category, formatNumber(item.value)])]);
  await pptx.writeFile({ fileName: `${safeFilename(state.fileName)}_通用分析_${reportDate.replace(/-/g, "")}.pptx` });
}

function renderMode() {
  state.mode = $("modeSelect").value;
  renderMappingPanel();
  renderReportSettings();
  if (state.mode === "orders") {
    updateOrderAnalysis();
  } else {
    setupGenericControls();
    renderGenericDashboard();
  }
}

async function readMappingProfile(file) {
  try {
    const profile = JSON.parse(await file.text());
    if (!profile || typeof profile.mapping !== "object") throw new Error("JSON 中未找到 mapping 字段。");
    state.profileMapping = profile.mapping;
    if (profile.sheet_name && state.workbook?.SheetNames.includes(profile.sheet_name)) {
      state.sheetName = profile.sheet_name;
      $("sheetSelect").value = profile.sheet_name;
      if (profile.header_row_excel) state.headerRow = Number(profile.header_row_excel);
      loadCurrentSheet(false);
      return;
    }
    if (profile.header_row_excel) state.headerRow = Number(profile.header_row_excel);
    readCurrentTable();
    setMessage("analysisNotice", "字段映射配置已载入，请检查当前列名是否匹配。", "success");
  } catch (error) {
    setMessage("analysisNotice", `无法载入映射配置：${error.message || error}`, "error");
  }
}

function downloadMappingProfile() {
  const mapping = getMappingFromUi();
  const profile = {
    template: "customer_order_delivery",
    source_file: state.fileName,
    sheet_name: state.sheetName,
    header_row_excel: state.headerRow,
    mapping,
  };
  const blob = new Blob([JSON.stringify(profile, null, 2)], { type: "application/json;charset=utf-8" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `${safeFilename(state.fileName)}_字段映射.json`;
  link.click();
  URL.revokeObjectURL(link.href);
}

function refreshOrderCsv() {
  const mapped = [
    ["customer", "customer"], ["order", "order"], ["line", "line"], ["item", "item"], ["description", "description"],
    ["due_date", "dueDate"], ["actual_date", "actualDate"], ["open_qty", "openQty"], ["backorder_qty", "backorderQty"],
    ["shipped_qty", "shippedQty"], ["carrier", "carrier"], ["destination", "destination"],
  ].filter(([field]) => state.mapping[field]);
  const headers = ["Status", ...mapped.map(([field]) => state.mapping[field]), "Excel Row"];
  const data = state.orderFiltered.map((row) => ["逾期", ...mapped.map(([field, property]) => {
    if (property === "dueDate" || property === "actualDate") return row[property] || "";
    if (["openQty", "backorderQty", "shippedQty"].includes(property)) return row[property] ?? "";
    return row[property] ?? "";
  }), row.excelRow]);
  downloadCsv(headers, data, `${safeFilename(state.fileName)}_逾期待办.csv`);
}

function refreshGenericCsv() {
  const headers = ["Excel行", ...state.headers];
  const rows = state.genericFiltered.map((row) => [row.__excelRow, ...state.headers.map((header) => row[header] == null ? "" : String(row[header]))]);
  downloadCsv(headers, rows, `${safeFilename(state.fileName)}_筛选明细.csv`);
}

function refreshGenericGroupedCsv() {
  const dim = $("genericDimension").value === NO_COLUMN ? null : $("genericDimension").value;
  const measure = $("genericMeasure").value === NO_COLUMN ? null : $("genericMeasure").value;
  const isCount = $("genericAggregation").value === "计数" || !measure;
  const headers = [dim || "分类", isCount ? "记录数" : `${$("genericAggregation").value}：${measure}`];
  const rows = state.genericGrouped.map((item) => [item.category, item.value]);
  downloadCsv(headers, rows, `${safeFilename(state.fileName)}_分类汇总.csv`);
}

function setUpEvents() {
  $("fileInput").addEventListener("change", (event) => {
    const file = event.target.files?.[0];
    if (file) loadWorkbook(file);
    event.target.value = "";
  });
  $("replaceFileButton").addEventListener("click", () => $("fileInput").click());
  const dropZone = $("dropZone");
  ["dragenter", "dragover"].forEach((type) => dropZone.addEventListener(type, (event) => {
    event.preventDefault();
    dropZone.classList.add("dragging");
  }));
  ["dragleave", "drop"].forEach((type) => dropZone.addEventListener(type, (event) => {
    event.preventDefault();
    dropZone.classList.remove("dragging");
  }));
  dropZone.addEventListener("drop", (event) => {
    const file = event.dataTransfer?.files?.[0];
    if (file) loadWorkbook(file);
  });
  $("sheetSelect").addEventListener("change", () => {
    state.sheetName = $("sheetSelect").value;
    state.profileMapping = {};
    state.genericDimension = null;
    state.genericMeasure = null;
    loadCurrentSheet(true);
  });
  $("headerRowInput").addEventListener("change", () => {
    state.headerRow = Math.max(1, Number($("headerRowInput").value) || 1);
    readCurrentTable();
  });
  $("reportDateInput").addEventListener("change", () => {
    if (state.mode === "orders") updateOrderAnalysis();
    else renderGenericDashboard();
  });
  $("modeSelect").addEventListener("change", renderMode);
  $("customerFilter").addEventListener("change", renderOrderDashboard);
  $("carrierFilter").addEventListener("change", renderOrderDashboard);
  $("statusFilter").addEventListener("change", renderOrderDashboard);
  $("orderSearch").addEventListener("input", renderOrderDashboard);
  $("genericAggregation").addEventListener("change", renderGenericDashboard);
  $("genericSearch").addEventListener("input", renderGenericDashboard);
  $("downloadOrderCsv").addEventListener("click", refreshOrderCsv);
  $("downloadGenericCsv").addEventListener("click", refreshGenericCsv);
  $("downloadGenericGroupedCsv").addEventListener("click", refreshGenericGroupedCsv);
  $("downloadMappingButton").addEventListener("click", downloadMappingProfile);
  $("profileInput").addEventListener("change", (event) => {
    const file = event.target.files?.[0];
    if (file) readMappingProfile(file);
    event.target.value = "";
  });
  document.querySelectorAll(".tab-button").forEach((button) => button.addEventListener("click", () => setActiveView(button.dataset.view)));
  $("generatePptButton").addEventListener("click", async () => {
    setMessage("reportMessage", "正在生成 PowerPoint 文件…");
    $("generatePptButton").disabled = true;
    try {
      if (state.mode === "orders") {
        const built = buildOrderAnalysis();
        if (!built.rows.length || !state.mapping.open_qty && !(state.mapping.ordered_qty && state.mapping.shipped_qty)) throw new Error(built.note);
        state.orderFiltered = orderFilterRows();
        await generateOrderPpt();
      } else {
        renderGenericDashboard();
        await generateGenericPpt();
      }
      setMessage("reportMessage", "PowerPoint 已生成并下载。", "success");
    } catch (error) {
      setMessage("reportMessage", `PPT 生成失败：${error.message || error}`, "error");
    } finally {
      $("generatePptButton").disabled = false;
    }
  });
}

document.addEventListener("DOMContentLoaded", () => {
  $("reportDateInput").value = localToday();
  state.reportDate = localToday();
  setUpEvents();
});

