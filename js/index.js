// js/index.js
// 功能：把 JSON 畫成接近 Excel 的多層表格

let allRows = [];
let tableData = null;

const tableHead = document.querySelector("#tableHead");
const tableBody = document.querySelector("#tableBody");
const searchWorkOrder = document.querySelector("#searchWorkOrder");
const resetBtn = document.querySelector("#resetBtn");

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


function renderExcelHead(data) {
  const groups = data.header_groups || [];
  const columns = data.columns || [];
  const summary = data.summary || {};

  let html = "";

  html += `<tr>`;

  groups.forEach(function (group) {
    html += `
      <th class="${escapeHtml(group.className)}" colspan="${group.colspan}">
        ${formatMultiLine(group.title)}
      </th>
    `;
  });

  html += `</tr>`;


  html += `<tr>`;

  columns.forEach(function (column) {
    const itemText = getSecondHeaderText(column.key);

    html += `
      <th class="header-label">
        ${escapeHtml(itemText)}
      </th>
    `;
  });

  html += `</tr>`;


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


function filterData() {
  const keyword = searchWorkOrder.value.trim().toLowerCase();

  const filteredRows = allRows.filter(function (row) {
    const workOrder = String(row.work_order || "").toLowerCase();

    return workOrder.includes(keyword);
  });

  renderExcelBody(filteredRows, tableData.columns || []);
}


function resetFilter() {
  searchWorkOrder.value = "";
  renderExcelBody(allRows, tableData.columns || []);
}


function getSecondHeaderText(key) {
  const map = {
    date: "",
    work_order: "",
    input_qty: "30-4118-018-XXX",
    incoming_wait_qc: "30-4118-018-XX",
    wait_issue_qty: "30-4118-018-XXX#1",
    process_1_1_qty: "30-4118-018-XXX--#1-1",
    process_1_1_ng: "",
    process_1_2_qty: "30-4118-018-XXX--#1-2",
    process_1_2_ng: "",
    wait_qc_01: "30-4118-018-XXX",
    ng_01: "",
    wait_outsource_qty: "30-4118-018-000--#2",
    vendor: "",
    outsource_wait_return_qty: "",
    outsource_ng: "",
    wait_qc_02: "30-4118-018-XXX",
    ng_02: "",
    not_stock_in_qty: "",
    not_allocated_qty: "",
    bad_qty: "",
    yield_rate: "",
  };

  return map[key] || "";
}


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


searchWorkOrder.addEventListener("input", filterData);
resetBtn.addEventListener("click", resetFilter);