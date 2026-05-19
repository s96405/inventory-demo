// js/index.js
// 功能：讀取 Excel 轉出的 JSON，顯示庫存查詢 Demo
// 注意：這版只顯示 Excel 裡面抓出的欄位，不額外加入自訂狀態

let inventoryData = [];

// =========================
// 抓取 HTML 元素
// =========================

const tableBody = document.querySelector("#tableBody");

const searchWorkOrder = document.querySelector("#searchWorkOrder");
const searchItemNo = document.querySelector("#searchItemNo");
const resetBtn = document.querySelector("#resetBtn");

const totalCount = document.querySelector("#totalCount");
const totalInputQty = document.querySelector("#totalInputQty");
const totalStockInQty = document.querySelector("#totalStockInQty");
const totalBadQty = document.querySelector("#totalBadQty");


// =========================
// 載入 JSON 資料
// =========================

fetch("./data/inventory_demo.json")
  .then(function (response) {
    return response.json();
  })
  .then(function (data) {
    inventoryData = data;
    renderPage(inventoryData);
  })
  .catch(function (error) {
    console.error("讀取 JSON 失敗：", error);

    tableBody.innerHTML = `
      <tr>
        <td colspan="11" class="empty">
          讀取資料失敗，請確認 data/inventory_demo.json 是否存在
        </td>
      </tr>
    `;
  });


// =========================
// 畫面渲染
// =========================

function renderPage(data) {
  renderSummaryCards(data);
  renderTable(data);
}


// =========================
// 統計卡片
// =========================

function renderSummaryCards(data) {
  let inputTotal = 0;
  let stockInTotal = 0;
  let badTotal = 0;

  data.forEach(function (item) {
    inputTotal += Number(item.input_qty) || 0;
    stockInTotal += Number(item.stock_in_qty) || 0;
    badTotal += Number(item.bad_qty) || 0;
  });

  totalCount.textContent = formatNumber(data.length);
  totalInputQty.textContent = formatNumber(inputTotal);
  totalStockInQty.textContent = formatNumber(stockInTotal);
  totalBadQty.textContent = formatNumber(badTotal);
}


// =========================
// 表格
// =========================

function renderTable(data) {
  if (data.length === 0) {
    tableBody.innerHTML = `
      <tr>
        <td colspan="11" class="empty">查無資料</td>
      </tr>
    `;
    return;
  }

  let html = "";

  data.forEach(function (item) {
    html += `
      <tr>
        <td>${escapeHtml(item.source_sheet)}</td>
        <td>${formatDate(item.date)}</td>
        <td>${escapeHtml(item.work_order)}</td>
        <td>${escapeHtml(item.item_no)}</td>
        <td>${formatNumber(item.input_qty)}</td>
        <td>${formatNumber(item.stock_in_qty)}</td>
        <td>${formatNumber(item.not_stock_in_qty)}</td>
        <td>${formatNumber(item.available_qty)}</td>
        <td>${formatNumber(item.shipped_qty)}</td>
        <td>${formatNumber(item.bad_qty)}</td>
        <td>${formatYield(item.yield_rate)}</td>
      </tr>
    `;
  });

  tableBody.innerHTML = html;
}


// =========================
// 搜尋功能
// =========================

function filterData() {
  const workOrderKeyword = searchWorkOrder.value.trim().toLowerCase();
  const itemNoKeyword = searchItemNo.value.trim().toLowerCase();

  const filteredData = inventoryData.filter(function (item) {
    const workOrder = String(item.work_order || "").toLowerCase();
    const itemNo = String(item.item_no || "").toLowerCase();

    const matchWorkOrder = workOrder.includes(workOrderKeyword);
    const matchItemNo = itemNo.includes(itemNoKeyword);

    return matchWorkOrder && matchItemNo;
  });

  renderPage(filteredData);
}


// =========================
// 清除條件
// =========================

function resetFilter() {
  searchWorkOrder.value = "";
  searchItemNo.value = "";

  renderPage(inventoryData);
}


// =========================
// 格式化工具
// =========================

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
searchItemNo.addEventListener("input", filterData);
resetBtn.addEventListener("click", resetFilter);