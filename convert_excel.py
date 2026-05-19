# convert_excel.py
# 功能：讀取客戶 Excel，抽出庫存/製令相關資料，轉成前端可用的 JSON
# 注意：這版不會自己新增「狀態」，只輸出 Excel 內抓到的欄位資料

import json
from pathlib import Path
from openpyxl import load_workbook

# =========================
# 路徑設定
# =========================

BASE_DIR = Path(__file__).parent

EXCEL_FILE = BASE_DIR / "excel" / "30-4118-018-XXX.xlsm"

OUTPUT_FILE = BASE_DIR / "data" / "inventory_demo.json"


# =========================
# 要讀取的工作表
# =========================

TARGET_SHEETS = [
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
    "650",
]


# =========================
# 欄位關鍵字設定
# =========================

FIELD_KEYWORDS = {
    "date": ["來料日期", "日期"],
    "work_order": ["製令單號", "製令"],
    "item_no": ["料號"],
    "input_qty": ["入料數", "結束料數", "數量"],
    "stock_in_qty": ["入庫"],
    "not_stock_in_qty": ["未入庫"],
    "available_qty": ["可出貨數量", "可出貨"],
    "shipped_qty": ["已出貨"],
    "bad_qty": ["總不良數量", "總不良數", "不良"],
    "yield_rate": ["直通率"],
}


def safe_text(value):
    """把 Excel 儲存格內容轉成乾淨文字"""

    if value is None:
        return ""

    return str(value).strip()


def safe_number(value):
    """把 Excel 數字轉成 float，轉不了就回傳 0"""

    if value is None or value == "":
        return 0

    if isinstance(value, (int, float)):
        return value

    text = str(value).replace(",", "").replace("%", "").strip()

    try:
        return float(text)
    except ValueError:
        return 0


def find_header_row(sheet):
    """找出表頭列"""

    for row_index in range(1, min(sheet.max_row, 30) + 1):
        row_values = [safe_text(cell.value) for cell in sheet[row_index]]
        row_text = " ".join(row_values)

        if "製令" in row_text and "料號" in row_text:
            return row_index

    return None


def build_column_map(sheet, header_row):
    """建立欄位對照表"""

    column_map = {}

    for cell in sheet[header_row]:
        header_text = safe_text(cell.value)

        if not header_text:
            continue

        for field_name, keywords in FIELD_KEYWORDS.items():
            for keyword in keywords:
                if keyword in header_text and field_name not in column_map:
                    column_map[field_name] = cell.column
                    break

    return column_map


def get_cell_value(sheet, row_index, column_map, field_name):
    """根據欄位名稱取得指定列的資料"""

    col_index = column_map.get(field_name)

    if not col_index:
        return ""

    return sheet.cell(row=row_index, column=col_index).value


def convert_excel_to_json():
    """主程式：讀 Excel → 整理資料 → 輸出 JSON"""

    if not EXCEL_FILE.exists():
        print(f"找不到 Excel 檔案：{EXCEL_FILE}")
        return

    workbook = load_workbook(EXCEL_FILE, data_only=True, read_only=True)

    result = []

    for sheet_name in TARGET_SHEETS:
        if sheet_name not in workbook.sheetnames:
            print(f"略過不存在的工作表：{sheet_name}")
            continue

        sheet = workbook[sheet_name]

        header_row = find_header_row(sheet)

        if not header_row:
            print(f"找不到表頭，略過工作表：{sheet_name}")
            continue

        column_map = build_column_map(sheet, header_row)

        if "work_order" not in column_map and "item_no" not in column_map:
            print(f"缺少製令或料號欄位，略過工作表：{sheet_name}")
            continue

        for row_index in range(header_row + 1, sheet.max_row + 1):
            work_order = safe_text(
                get_cell_value(sheet, row_index, column_map, "work_order")
            )

            item_no = safe_text(get_cell_value(sheet, row_index, column_map, "item_no"))

            if not work_order and not item_no:
                continue

            row_data = {
                "source_sheet": sheet_name,
                "date": safe_text(get_cell_value(sheet, row_index, column_map, "date")),
                "work_order": work_order,
                "item_no": item_no,
                "input_qty": safe_number(
                    get_cell_value(sheet, row_index, column_map, "input_qty")
                ),
                "stock_in_qty": safe_number(
                    get_cell_value(sheet, row_index, column_map, "stock_in_qty")
                ),
                "not_stock_in_qty": safe_number(
                    get_cell_value(sheet, row_index, column_map, "not_stock_in_qty")
                ),
                "available_qty": safe_number(
                    get_cell_value(sheet, row_index, column_map, "available_qty")
                ),
                "shipped_qty": safe_number(
                    get_cell_value(sheet, row_index, column_map, "shipped_qty")
                ),
                "bad_qty": safe_number(
                    get_cell_value(sheet, row_index, column_map, "bad_qty")
                ),
                "yield_rate": safe_number(
                    get_cell_value(sheet, row_index, column_map, "yield_rate")
                ),
            }

            result.append(row_data)

    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)

    with open(OUTPUT_FILE, "w", encoding="utf-8") as file:
        json.dump(result, file, ensure_ascii=False, indent=2)

    print(f"轉換完成，共 {len(result)} 筆資料")
    print(f"輸出檔案：{OUTPUT_FILE}")


if __name__ == "__main__":
    convert_excel_to_json()
