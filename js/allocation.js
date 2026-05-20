const allocationBody = document.querySelector("#allocationBody");
const searchOrder = document.querySelector("#searchOrder");
const typeSelect = document.querySelector("#typeSelect");
const resetBtn = document.querySelector("#resetBtn");
const summaryCards = document.querySelector("#summaryCards");
const allocationContent = document.querySelector("#allocationContent");
const noAllocationMessage = document.querySelector("#noAllocationMessage");

let allocationRows = [];
let allocationSummary = [];

fetch("data/inventory_demo.json")
  .then(function (response) {
    return response.json();
  })
  .then(function (data) {
    allocationRows = data.allocations || [];
    allocationSummary = data.allocation_summary || [];

    if (allocationRows.length === 0) {
      allocationContent.classList.add("hidden");
      noAllocationMessage.classList.remove("hidden");
      return;
    }

    allocationContent.classList.remove("hidden");
    noAllocationMessage.classList.add("hidden");

    renderSheetOptions();
    renderSummaryCards();
    renderAllocationTable();
  })
  .catch(function (error) {
    console.error("讀取分配明細 JSON 失敗：", error);

    summaryCards.innerHTML = `
      <div class="summary-card">
        <p>分配摘要讀取失敗</p>
      </div>
    `;

    allocationBody.innerHTML = `
      <tr>
        <td colspan="6" class="empty">讀取資料失敗</td>
      </tr>
    `;
  });

function renderSheetOptions() {
  const sheetSet = new Set();

  allocationRows.forEach(function (row) {
    if (row.sheet_name) {
      sheetSet.add(row.sheet_name);
    }
  });

  const sheets = Array.from(sheetSet).sort();

  let html = `<option value="all">全部</option>`;

  sheets.forEach(function (sheetName) {
    html += `
      <option value="${escapeHtml(sheetName)}">
        ${escapeHtml(sheetName)}
      </option>
    `;
  });

  typeSelect.innerHTML = html;
}

function renderSummaryCards() {
  const sheetName = typeSelect.value;

  // 全部模式不要顯示一堆卡片
  if (sheetName === "all") {
    summaryCards.innerHTML = `
      <div class="summary-card">
        <p>請選擇製程分頁後查看分配摘要</p>
      </div>
    `;
    return;
  }

  const filteredSummary = allocationSummary.filter(function (item) {
    return item.sheet_name === sheetName;
  });

  if (filteredSummary.length === 0) {
    summaryCards.innerHTML = `
      <div class="summary-card">
        <p>此製程分頁沒有分配摘要</p>
      </div>
    `;
    return;
  }

  let html = "";

  filteredSummary.forEach(function (item) {
    html += `
      <div class="summary-card">
        <h2>${escapeHtml(item.allocation_type)}</h2>
        <p>分配數量：${formatNumber(item.total_qty)}</p>
        <p>製程中數量：${formatNumber(item.in_process_qty)}</p>
        <p>已做好待檢驗：${formatNumber(item.waiting_qc_qty)}</p>
      </div>
    `;
  });

  summaryCards.innerHTML = html;
}

function renderAllocationTable() {
  const keyword = searchOrder.value.trim().toLowerCase();
  const sheetName = typeSelect.value;

  const filteredRows = allocationRows.filter(function (row) {
    const fromOrder = String(row.from_work_order || "").toLowerCase();
    const toOrder = String(row.to_work_order || "").toLowerCase();

    const matchOrder =
      fromOrder.includes(keyword) ||
      toOrder.includes(keyword);

    const matchSheet =
      sheetName === "all" ||
      row.sheet_name === sheetName;

    return matchOrder && matchSheet;
  });

  if (filteredRows.length === 0) {
    allocationBody.innerHTML = `
      <tr>
        <td colspan="6" class="empty">查無資料</td>
      </tr>
    `;
    return;
  }

  let html = "";

  filteredRows.forEach(function (row) {
    html += `
      <tr>
        <td>${escapeHtml(formatDate(row.date))}</td>
        <td>${escapeHtml(row.sheet_name)}</td>
        <td>${escapeHtml(row.from_work_order)}</td>
        <td>${escapeHtml(row.allocation_type)}</td>
        <td>${escapeHtml(row.to_work_order)}</td>
        <td>${formatNumber(row.allocation_qty)}</td>
      </tr>
    `;
  });

  allocationBody.innerHTML = html;
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
    return String(value);
  }

  return date.toLocaleDateString("zh-TW");
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

searchOrder.addEventListener("input", renderAllocationTable);

typeSelect.addEventListener("change", function () {
  renderSummaryCards();
  renderAllocationTable();
});

resetBtn.addEventListener("click", function () {
  searchOrder.value = "";
  typeSelect.value = "all";
  renderSummaryCards();
  renderAllocationTable();
});