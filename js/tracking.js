// 功能：仿 Excel 製程追蹤表

// 1. 顯示全部欄位 / 未加工餘料欄位 / NG不良欄位

// 2. 搜尋後重新計算目前篩選結果總數

// 3. 不加入系統判斷欄位，只顯示 Excel 轉出的資料

let allRows = [];
let tableData = null;

const tableHead = document.querySelector("#tableHead");
const tableBody = document.querySelector("#tableBody");
const searchWorkOrder = document.querySelector("#searchWorkOrder");
const columnMode = document.querySelector("#columnMode");
const resetBtn = document.querySelector("#resetBtn");
const sheetSelect = document.querySelector("#sheetSelect");


/* 前端欄位設定 */
const FRONTEND_COLUMNS = [
  {
    key: "date",
    label: "來料日期",
    type: "text",
    groupTitle: "115年度",
    groupClass: "group-year",
    secondTitle: "",
    fixed: true
  },
  {
    key: "work_order",
    label: "製令單號",
    type: "text",
    groupTitle: "素材廠：淳梓（自購料）\n30-4118-018-XXX#1\n定容定量：小籃168pcs",
    groupClass: "group-material",
    secondTitle: "",
    fixed: true
  },
  {
    key: "input_qty",
    label: "入料數",
    type: "number",
    groupTitle: "總來料數",
    groupClass: "group-total",
    secondTitle: "30-4118-018-XX"
  },
  {
    key: "incoming_wait_qc",
    label: "進料待檢",
    type: "number",
    groupTitle: "品管進料檢驗",
    groupClass: "group-qc",
    secondTitle: "30-4118-018-XX"
  },
  {
    key: "wait_issue_qty",
    label: "#01待發料數量",
    type: "number",
    groupTitle: "發料\n168/籃",
    groupClass: "group-issue",
    secondTitle: "30-4118-018-XXX#1"
  },
  {
    key: "process_1_1_qty",
    label: "#1-1未加工",
    type: "number",
    groupTitle: "車床--粗車",
    groupClass: "group-process",
    secondTitle: "30-4118-018-XXX--#1-1"
  },
  {
    key: "process_1_1_ng",
    label: "#1-1 NG",
    type: "number",
    groupTitle: "車床--粗車",
    groupClass: "group-process",
    secondTitle: "30-4118-018-XXX--#1-1",
    ng: true
  },
  {
    key: "process_1_2_qty",
    label: "#1-2未加工",
    type: "number",
    groupTitle: "車床",
    groupClass: "group-process",
    secondTitle: "30-4118-018-XXX--#1-2"
  },
  {
    key: "process_1_2_ng",
    label: "#1-2 NG",
    type: "number",
    groupTitle: "車床",
    groupClass: "group-process",
    secondTitle: "30-4118-018-XXX--#1-2",
    ng: true
  },
  {
    key: "wait_qc_01",
    label: "#01待檢驗",
    type: "number",
    groupTitle: "外觀檢驗",
    groupClass: "group-qc",
    secondTitle: "30-4118-018-XXX"
  },
  {
    key: "ng_01",
    label: "#01NG",
    type: "number",
    groupTitle: "外觀檢驗",
    groupClass: "group-qc",
    secondTitle: "30-4118-018-XXX",
    ng: true
  },
  {
    key: "wait_outsource_qty",
    label: "#02待委外數量",
    type: "number",
    groupTitle: "委外－俊杰 Y003 頭部拋磨",
    groupClass: "group-outsource",
    secondTitle: "30-4118-018-000--#2"
  },
  {
    key: "vendor",
    label: "#02廠商",
    type: "text",
    groupTitle: "委外－俊杰 Y003 頭部拋磨",
    groupClass: "group-outsource",
    secondTitle: "30-4118-018-000--#2"
  },
  {
    key: "outsource_wait_return_qty",
    label: "#02待回數量",
    type: "number",
    groupTitle: "委外－俊杰 Y003 頭部拋磨",
    groupClass: "group-outsource",
    secondTitle: "30-4118-018-000--#2"
  },
  {
    key: "outsource_ng",
    label: "#02委外NG",
    type: "number",
    groupTitle: "委外－俊杰 Y003 頭部拋磨",
    groupClass: "group-outsource",
    secondTitle: "30-4118-018-000--#2",
    ng: true
  },
  {
    key: "wait_qc_02",
    label: "#02待檢驗",
    type: "number",
    groupTitle: "品管檢驗",
    groupClass: "group-qc",
    secondTitle: "30-4118-018-XXX"
  },
  {
    key: "ng_02",
    label: "#02 NG",
    type: "number",
    groupTitle: "品管檢驗",
    groupClass: "group-qc",
    secondTitle: "30-4118-018-XXX",
    ng: true
  },
  {
    key: "not_stock_in_qty",
    label: "未入庫數量",
    type: "number",
    groupTitle: "尚未入庫",
    groupClass: "group-stock",
    secondTitle: ""
  },
  {
    key: "not_allocated_qty",
    label: "尚未分料數量",
    type: "number",
    groupTitle: "尚未分料",
    groupClass: "group-stock",
    secondTitle: ""
  },
  {
    key: "bad_qty",
    label: "總不良數",
    type: "number",
    groupTitle: "總不良數量",
    groupClass: "group-ng",
    secondTitle: "",
    ng: true
  },
  {
    key: "yield_rate",
    label: "直通率",
    type: "number",
    groupTitle: "直通率",
    groupClass: "group-rate",
    secondTitle: ""
  }
];

fetch("/inventory-demo/data/inventory_demo.json")
  .then(function (response) {
    return response.json();
  })
  .then(function (data) {
    tableData = data;

    renderSheetOptions();
    renderPage();
  })
  .catch(function (error) {
    console.error("讀取 JSON 失敗：", error);

    tableBody.innerHTML = `
      <tr>
        <td class="empty">讀取資料失敗，請確認 data/inventory_demo.json 是否存在</td>
      </tr>
    `;
  });

function renderPage() {
  if (!tableData) {
    return;
  }

  const filteredRows = getFilteredRows();
  const visibleColumns = getVisibleColumns(filteredRows);

  document
    .querySelector(".excel-table")
    .classList.toggle("compact-table", columnMode.value === "remain");

  renderExcelHead(filteredRows, visibleColumns);
  renderExcelBody(filteredRows, visibleColumns);
}

function toNumber(value) {
  if (value === null || value === undefined || value === "") {
    return 0;
  }

  return Number(String(value).replaceAll(",", "")) || 0;
}

function getFilteredRows() {
  const keyword = searchWorkOrder.value.trim().toLowerCase();
  const sheetName = sheetSelect.value;
  const mode = columnMode.value;

  const currentSheet = tableData.sheets?.[sheetName] || {};
  const rows = currentSheet.rows || [];

  return rows.filter(function (row) {
    const workOrder = String(row.work_order || "").toLowerCase();

    // 製令單號搜尋
    if (!workOrder.includes(keyword)) {
      return false;
    }

    // 未加工 / 餘料模式：
    // 只要除了入料數、NG、直通率以外，有任何數量 > 0，就顯示這列
    if (mode === "remain") {
      return FRONTEND_COLUMNS.some(function (column) {
        return isRemainQtyColumn(column) && toNumber(row[column.key]) > 0;
      });
    }

    return true;
  });
}

function isRemainQtyColumn(column) {
  // 固定欄位：來料日期、製令單號，不拿來判斷
  if (column.fixed) {
    return false;
  }

  // 入料數不算未完成
  if (column.key === "input_qty") {
    return false;
  }

  // 廠商是文字欄位，不拿來判斷數量
  if (column.key === "vendor") {
    return false;
  }

  // NG 不算未完成
  if (column.ng) {
    return false;
  }

  // 直通率不算未完成
  if (column.key === "yield_rate") {
    return false;
  }

  // 剩下的數量欄位都算未完成
  return column.type === "number";
}

function getVisibleColumns(rows) {
  const mode = columnMode.value;

  if (mode === "all") {
    return FRONTEND_COLUMNS;
  }

  if (mode === "remain") {
  // 只要目前資料裡 #02待回數量 有任何一筆 > 0，才顯示廠商欄位
  const hasOutsourceWaitReturn = rows.some(function (row) {
    return toNumber(row.outsource_wait_return_qty) > 0;
  });

  return FRONTEND_COLUMNS.filter(function (column) {
    // 固定欄位：來料日期、製令單號
    if (column.fixed) {
      return true;
    }

    // 廠商欄位：#02待回數量全部是 0 就不顯示
    if (column.key === "vendor") {
      return hasOutsourceWaitReturn;
    }

    // 不是未完成數量欄位就不顯示
    if (!isRemainQtyColumn(column)) {
      return false;
    }

    // 這個欄位在目前資料裡有數量 > 0 才顯示
    return rows.some(function (row) {
      return toNumber(row[column.key]) > 0;
    });
  });
}

  if (mode === "ng") {
    return FRONTEND_COLUMNS.filter(function (column) {
      return column.fixed || column.ng;
    });
  }

  return FRONTEND_COLUMNS;
}

function renderExcelHead(rows, columns) {
  let html = "";

  html += renderGroupRow(columns);
  html += renderSecondTitleRow(columns);
  html += renderSummaryRow(rows, columns);
  html += renderColumnLabelRow(columns);

  tableHead.innerHTML = html;
}

function renderGroupRow(columns) {
  let html = `<tr>`;

  let index = 0;

  while (index < columns.length) {
    const current = columns[index];

    let colspan = 1;
    let nextIndex = index + 1;

    while (
      nextIndex < columns.length &&
      columns[nextIndex].groupTitle === current.groupTitle &&
      columns[nextIndex].groupClass === current.groupClass
    ) {
      colspan++;
      nextIndex++;
    }

    html += `
      <th class="${escapeHtml(current.groupClass)}" colspan="${colspan}">
        ${formatMultiLine(current.groupTitle)}
      </th>
    `;

    index = nextIndex;
  }

  html += `</tr>`;

  return html;
}

function renderSecondTitleRow(columns) {
  let html = `<tr>`;

  let index = 0;

  while (index < columns.length) {
    const current = columns[index];

    let colspan = 1;
    let nextIndex = index + 1;

    while (
      nextIndex < columns.length &&
      columns[nextIndex].secondTitle === current.secondTitle &&
      columns[nextIndex].groupTitle === current.groupTitle
    ) {
      colspan++;
      nextIndex++;
    }

    html += `
      <th class="header-label" colspan="${colspan}">
        ${escapeHtml(current.secondTitle)}
      </th>
    `;

    index = nextIndex;
  }

  html += `</tr>`;

  return html;
}

function renderSummaryRow(rows, columns) {
  let html = `<tr>`;

  columns.forEach(function (column) {
    let value = "";

    if (column.type === "number") {
      const total = rows.reduce(function (sum, row) {
        return sum + (Number(row[column.key]) || 0);
      }, 0);

      value = formatNumber(total);
    }

    if (column.key === "yield_rate") {
      value = calcAverageYield(rows);
    }

    const className = column.ng ? "summary-cell ng" : "summary-cell";

    html += `
      <th class="${className}">
        ${value}
      </th>
    `;
  });

  html += `</tr>`;

  return html;
}

function renderColumnLabelRow(columns) {
  let html = `<tr>`;

  columns.forEach(function (column) {
    html += `
      <th class="header-label">
        ${escapeHtml(column.label)}
      </th>
    `;
  });

  html += `</tr>`;

  return html;
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

      if (column.ng) {
        className += " ng";
      }
        
      if (
        columnMode.value === "remain" &&
        column.type === "number" &&
        toNumber(value) > 0
        ) {
        className += " has-qty";
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


function formatCellValue(value, column) {
  if (column.key === "date") {
    return formatDate(value);
  }

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

function formatDate(value) {
  if (!value) {
    return "";
  }

  const date = new Date(value);

  if (isNaN(date.getTime())) {
    return String(value).split(" ")[0];
  }

  return date.toLocaleDateString("zh-TW");
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

function calcAverageYield(rows) {
  const validRows = rows.filter(function (row) {
    return Number(row.yield_rate) > 0;
  });

  if (validRows.length === 0) {
    return "-";
  }

  const total = validRows.reduce(function (sum, row) {
    return sum + Number(row.yield_rate);
  }, 0);

  return formatYield(total / validRows.length);
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


sheetSelect.addEventListener("change", renderPage);
searchWorkOrder.addEventListener("input", renderPage);
columnMode.addEventListener("change", renderPage);
sheetSelect.addEventListener("change", renderPage);
resetBtn.addEventListener("click", function () {
  searchWorkOrder.value = "";
  columnMode.value = "all";

  if (tableData && tableData.sheets && tableData.sheets["018#1"]) {
    sheetSelect.value = "018#1";
  }

  renderPage();
});

function renderSheetOptions() {
  if (!tableData || !tableData.sheets) {
    return;
  }

  const sheetOrder = [
    "018#1",
    "4XX",
    "5XX",
    "6XX",
    "370",
    "400",
    "420",
    "450",
    "470",
    "500",
    "520",
    "550",
    "570",
    "600",
    "620",
    "650"
  ];

  const sheetNames = Object.keys(tableData.sheets).sort(function (a, b) {
    const indexA = sheetOrder.indexOf(a);
    const indexB = sheetOrder.indexOf(b);

    if (indexA === -1 && indexB === -1) {
      return a.localeCompare(b);
    }

    if (indexA === -1) {
      return 1;
    }

    if (indexB === -1) {
      return -1;
    }

    return indexA - indexB;
  });

  let html = "";

  sheetNames.forEach(function (sheetName) {
    html += `
      <option value="${escapeHtml(sheetName)}">
        ${escapeHtml(sheetName)}
      </option>
    `;
  });

  sheetSelect.innerHTML = html;

  if (sheetNames.includes("018#1")) {
    sheetSelect.value = "018#1";
  }
}