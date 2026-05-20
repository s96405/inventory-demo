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

    allRows = data.rows || [];
      
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
  const filteredRows = getFilteredRows();
  const visibleColumns = getVisibleColumns(filteredRows);

  document
    .querySelector(".excel-table")
    .classList.toggle("compact-table", columnMode.value === "remain");

  renderExcelHead(filteredRows, visibleColumns);
  renderExcelBody(filteredRows, visibleColumns);
}

function getFilteredRows() {
  const keyword = searchWorkOrder.value.trim().toLowerCase();

  return allRows.filter(function (row) {
    const workOrder = String(row.work_order || "").toLowerCase();

    return workOrder.includes(keyword);
  });
}

function getVisibleColumns(rows) {
  const mode = columnMode.value;

  if (mode === "all") {
    return FRONTEND_COLUMNS;
  }

  if (mode === "remain") {
  return FRONTEND_COLUMNS.filter(function (column) {
    // 固定欄位：來料日期、製令單號
    if (column.fixed) {
      return true;
    }

    // 保留廠商欄位
    if (column.key === "vendor") {
      return true;
    }

    // NG 欄位不要顯示
    if (column.ng) {
      return false;
    }

    // 顯示：未加工數量、廠商待回數量、未入庫數量
    const isRemainColumn =
      column.label.includes("未加工") ||
      column.label.includes("待回數量") ||
      column.label.includes("未入庫數量");

    if (!isRemainColumn) {
      return false;
    }

    // 只有目前查詢結果裡，有數量的欄位才顯示
    return rows.some(function (row) {
      return Number(row[column.key]) > 0;
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
        Number(value) > 0
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

searchWorkOrder.addEventListener("input", renderPage);
columnMode.addEventListener("change", renderPage);
resetBtn.addEventListener("click", function () {
  searchWorkOrder.value = "";
  columnMode.value = "all";
  renderPage();
});