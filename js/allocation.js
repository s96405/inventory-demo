const allocationBody = document.querySelector("#allocationBody");
const searchOrder = document.querySelector("#searchOrder");
const typeSelect = document.querySelector("#typeSelect");
const resetBtn = document.querySelector("#resetBtn");
const summaryCards = document.querySelector("#summaryCards");

let allocationRows = [];
let allocationSummary = [];

fetch("data/inventory_demo.json")
  .then(function (response) {
    return response.json();
  })
  .then(function (data) {
    allocationRows = data.allocations || [];
    allocationSummary = data.allocation_summary || [];

    renderTypeOptions();
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
        <td colspan="5" class="empty">讀取資料失敗</td>
      </tr>
    `;
  });

function renderTypeOptions() {
  const typeSet = new Set();

  allocationRows.forEach(function (row) {
    if (row.allocation_type) {
      typeSet.add(row.allocation_type);
    }
  });

  const types = Array.from(typeSet).sort();

  let html = `<option value="all">全部</option>`;

  types.forEach(function (type) {
    html += `
      <option value="${escapeHtml(type)}">
        ${escapeHtml(type)}
      </option>
    `;
  });

  typeSelect.innerHTML = html;
}

function renderSummaryCards() {
  const type = typeSelect.value;

  const filteredSummary = allocationSummary.filter(function (item) {
    return type === "all" || item.allocation_type === type;
  });

  if (filteredSummary.length === 0) {
    summaryCards.innerHTML = `
      <div class="summary-card">
        <p>查無分配摘要</p>
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
  const type = typeSelect.value;

  const filteredRows = allocationRows.filter(function (row) {
    const fromOrder = String(row.from_work_order || "").toLowerCase();
    const toOrder = String(row.to_work_order || "").toLowerCase();

    const matchOrder =
      fromOrder.includes(keyword) ||
      toOrder.includes(keyword);

    const matchType =
      type === "all" ||
      row.allocation_type === type;

    return matchOrder && matchType;
  });

  if (filteredRows.length === 0) {
    allocationBody.innerHTML = `
      <tr>
        <td colspan="5" class="empty">查無資料</td>
      </tr>
    `;
    return;
  }

  let html = "";

  filteredRows.forEach(function (row) {
    html += `
      <tr>
        <td>${escapeHtml(formatDate(row.date))}</td>
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