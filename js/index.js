// js/index.js
// 功能：把 JSON 畫成接近 Excel 的多層製程追蹤表
// 重點：
// 1. 第 1 列：製程大分類
// 2. 第 2 列：料號 / 製程說明，使用 colspan 合併
// 3. 第 3 列：製程總數
// 4. 第 4 列：欄位名稱
// 5. tbody：每筆製令資料

let allRows = [];
let tableData = null;

const tableHead = document.querySelector("#tableHead");
const tableBody = document.querySelector("#tableBody");
const searchWorkOrder = document.querySelector("#searchWorkOrder");
const resetBtn = document.querySelector("#resetBtn");

// =========================
// 載入 JSON
// =========================

fetch("./data/inventory_demo.json")
  .then(function (response) {
    return response.json();
  })
  .then(function (data) {
    tableData = data;
    allRows = data.rows || [];

    renderExcelHead(data);
    renderExcelBody(allRows, data.columns || []);
  })
  .catch(function (error) {
    console.error("讀取 JSON 失敗：", error);

    tableBody.innerHTML = `
      <tr>
        <td class="empty">讀取資料失敗，請確認 data/inventory_demo.json 是否存在</td>
      </tr>
    `;
  });

// =========================
// 產生 Excel 表頭
// =========================

function renderExcelHead(data) {
  const groups = data.header_groups || [];
  const columns = data.columns || [];
  const summary = data.summary || {};

  let html = "";

  // =========================
  // 第 1 列：製程大分類
  // =========================
  html += `<tr>`;

  groups.forEach(function (group) {
    html += `
      <th class="${escapeHtml(group.className)}" colspan="${group.colspan}">
        ${formatMultiLine(group.title)}
      </th>
    `;
  });

  html += `</tr>`;

  // =========================
  // 第 2 列：料號 / 製程說明
  // 這裡用 colspan 合併，讓它更像 Excel
  // 欄位總數必須跟 columns 數量一致：目前 21 欄
  // =========================
  html += `
    <tr>
      <th class="header-label"></th>
      <th class="header-label"></th>
      <th class="header-label">30-4118-018-XX</th>
      <th class="header-label">30-4118-018-XX</th>
      <th class="header-label">30-4118-018-XXX#1</th>

      <th class="header-label" colspan="2">30-4118-018-XXX--#1-1</th>

      <th class="header-label" colspan="2">30-4118-018-XXX--#1-2</th>

      <th class="header-label" colspan="2">30-4118-018-XXX</th>

      <th class="header-label" colspan="4">30-4118-018-000--#2</th>

      <th class="header-label" colspan="2">30-4118-018-XXX</th>

      <th class="header-label"></th>
      <th class="header-label"></th>
      <th class="header-label"></th>
      <th class="header-label"></th>
    </tr>
  `;

  // =========================
  // 第 3 列：製程總數
  // =========================
  html += `<tr>`;

  columns.forEach(function (column) {
    let value = "";

    if (column.type === "number") {
      value = formatNumber(summary[column.key]);
    }

    const ngClass = isNgColumn(column.key) ? "ng" : "";

    html += `
      <th class="summary-cell ${ngClass}">
        ${value}
      </th>
    `;
  });

  html += `</tr>`;

  // =========================
  // 第 4 列：欄位名稱
  // =========================
  html += `<tr>`;

  columns.forEach(function (column) {
    html += `
      <th class="header-label">
        ${escapeHtml(column.label)}
      </th>
    `;
  });

  html += `</tr>`;

  tableHead.innerHTML = html;
}

// =========================
// 產生表格資料
// =========================

function renderExcelBody(rows, columns) {
  if (rows.length === 0) {
    tableBody.innerHTML = `
      <tr>
        <td colspan="${columns.length}" class="empty">查無資料</td>
      </tr>
    `;
    return;
  }

  let html = "";

  rows.forEach(function (row) {
    html += `<tr>`;

    columns.forEach(function (column) {
      const key = column.key;
      const value = row[key];

      let className = "";

      if (column.type === "number") {
        className += " number";
      }

      if (isNgColumn(key)) {
        className += " ng";
      }

      if (key === "work_order") {
        className += " work-order-cell";
      }

      if (key === "yield_rate") {
        className += ` ${getYieldClass(value)}`;
      }

      html += `
        <td class="${className}">
          ${formatCellValue(value, column)}
        </td>
      `;
    });

    html += `</tr>`;
  });

  tableBody.innerHTML = html;
}

// =========================
// 搜尋功能
// =========================

function filterData() {
  const keyword = searchWorkOrder.value.trim().toLowerCase();

  const filteredRows = allRows.filter(function (row) {
    const workOrder = String(row.work_order || "").toLowerCase();

    return workOrder.includes(keyword);
  });

  renderExcelBody(filteredRows, tableData.columns || []);
}

// =========================
// 清除條件
// =========================

function resetFilter() {
  searchWorkOrder.value = "";
  renderExcelBody(allRows, tableData.columns || []);
}

// =========================
// 格式化資料
// =========================

function formatCellValue(value, column) {
  if (column.key === "yield_rate") {
    return formatYield(value);
  }

  if (column.type === "number") {
    return formatNumber(value);
  }

  return escapeHtml(value);
}

function formatNumber(value) {
  const number = Number(value) || 0;

  return number.toLocaleString("zh-TW");
}

function formatYield(value) {
  const number = Number(value) || 0;

  if (number === 0) {
    return "-";
  }

  // Excel 如果讀出 0.9966，代表 99.66%
  if (number <= 1) {
    return `${(number * 100).toFixed(2)}%`;
  }

  return `${number.toFixed(2)}%`;
}

function getYieldClass(value) {
  const number = Number(value) || 0;
  const percent = number <= 1 ? number * 100 : number;

  if (percent >= 98) {
    return "rate-good";
  }

  if (percent >= 95) {
    return "rate-warning";
  }

  return "rate-bad";
}

function isNgColumn(key) {
  return key.toLowerCase().includes("ng") || key === "bad_qty";
}

function formatMultiLine(text) {
  return escapeHtml(text).replaceAll("\n", "<br>");
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

// =========================
// 綁定事件
// =========================

searchWorkOrder.addEventListener("input", filterData);
resetBtn.addEventListener("click", resetFilter);