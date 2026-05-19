# convert_excel.py
# 功能：
# 1. 讀取 Excel 的 018#1 工作表
# 2. 抓出 Excel 上方製程總數
# 3. 抓出每筆製令的製程數量
# 4. 輸出 data/inventory_demo.json 給前端使用

import json
from pathlib import Path
from openpyxl import load_workbook

BASE_DIR = Path(__file__).parent

EXCEL_FILE = BASE_DIR / "excel" / "30-4118-018-XXX.xlsm"
OUTPUT_FILE = BASE_DIR / "data" / "inventory_demo.json"

TARGET_SHEET = "018#1"


COLUMNS = [
    {"key": "date", "label": "來料日期", "keywords": ["來料日期"], "type": "text"},
    {
        "key": "work_order",
        "label": "製令單號",
        "keywords": ["製令單號"],
        "type": "text",
    },
    {"key": "input_qty", "label": "入料數", "keywords": ["入料數"], "type": "number"},
    {
        "key": "incoming_wait_qc",
        "label": "進料待檢",
        "keywords": ["進料待檢"],
        "type": "number",
    },
    {
        "key": "wait_issue_qty",
        "label": "#01待發料數量",
        "keywords": ["#01待發料數量"],
        "type": "number",
    },
    {
        "key": "process_1_1_qty",
        "label": "#1-1未加工",
        "keywords": ["#1-1未加工"],
        "type": "number",
    },
    {
        "key": "process_1_1_ng",
        "label": "#1-1 NG",
        "keywords": ["#1-1NG", "#1-1 NG"],
        "type": "number",
    },
    {
        "key": "process_1_2_qty",
        "label": "#1-2未加工",
        "keywords": ["#1-2未加工"],
        "type": "number",
    },
    {
        "key": "process_1_2_ng",
        "label": "#1-2 NG",
        "keywords": ["#1-2NG", "#1-2 NG"],
        "type": "number",
    },
    {
        "key": "wait_qc_01",
        "label": "#01待檢驗",
        "keywords": ["#01待檢驗"],
        "type": "number",
    },
    {
        "key": "ng_01",
        "label": "#01NG",
        "keywords": ["#01NG", "#01 NG"],
        "type": "number",
    },
    {
        "key": "wait_outsource_qty",
        "label": "#02待委外數量",
        "keywords": ["#02待委外數量"],
        "type": "number",
    },
    {"key": "vendor", "label": "#02廠商", "keywords": ["#02廠商"], "type": "text"},
    {
        "key": "outsource_wait_return_qty",
        "label": "#02待回數量",
        "keywords": ["#02待回數量"],
        "type": "number",
    },
    {
        "key": "outsource_ng",
        "label": "#02委外NG",
        "keywords": ["#02委外NG"],
        "type": "number",
    },
    {
        "key": "wait_qc_02",
        "label": "#02待檢驗",
        "keywords": ["#02待檢驗"],
        "type": "number",
    },
    {
        "key": "ng_02",
        "label": "#02 NG",
        "keywords": ["#02NG", "#02 NG"],
        "type": "number",
    },
    {
        "key": "not_stock_in_qty",
        "label": "未入庫數量",
        "keywords": ["未入庫數量"],
        "type": "number",
    },
    {
        "key": "not_allocated_qty",
        "label": "尚未分料數量",
        "keywords": ["尚未分料數量"],
        "type": "number",
    },
    {"key": "bad_qty", "label": "總不良數", "keywords": ["總不良數"], "type": "number"},
    {"key": "yield_rate", "label": "直通率", "keywords": ["直通率"], "type": "number"},
]


HEADER_GROUPS = [
    {"title": "115年度", "colspan": 1, "className": "group-year"},
    {
        "title": "素材廠：淳梓（自購料）\n30-4118-018-XXX#1\n定容定量：小籃168pcs",
        "colspan": 1,
        "className": "group-material",
    },
    {"title": "總來料數", "colspan": 1, "className": "group-total"},
    {"title": "品管進料檢驗", "colspan": 1, "className": "group-qc"},
    {"title": "發料\n168/籃", "colspan": 1, "className": "group-issue"},
    {"title": "車床--粗車", "colspan": 2, "className": "group-process"},
    {"title": "車床", "colspan": 2, "className": "group-process"},
    {"title": "外觀檢驗", "colspan": 2, "className": "group-qc"},
    {"title": "委外－俊杰 Y003 頭部拋磨", "colspan": 4, "className": "group-outsource"},
    {"title": "品管檢驗", "colspan": 2, "className": "group-qc"},
    {"title": "尚未入庫", "colspan": 1, "className": "group-stock"},
    {"title": "尚未分料", "colspan": 1, "className": "group-stock"},
    {"title": "總不良數量", "colspan": 1, "className": "group-ng"},
    {"title": "直通率", "colspan": 1, "className": "group-rate"},
]


def safe_text(value):
    if value is None:
        return ""

    return str(value).strip()


def safe_number(value):
    if value is None or value == "":
        return 0

    if isinstance(value, (int, float)):
        return value

    text = str(value).replace(",", "").replace("%", "").strip()

    try:
        return float(text)
    except ValueError:
        return 0


def normalize_text(value):
    return safe_text(value).replace(" ", "").replace("\n", "").replace("\r", "")


def find_header_row(sheet):
    for row_index in range(1, min(sheet.max_row, 30) + 1):
        row_values = [normalize_text(cell.value) for cell in sheet[row_index]]
        row_text = " ".join(row_values)

        if "來料日期" in row_text and "製令單號" in row_text and "入料數" in row_text:
            return row_index

    return None


def build_column_map(sheet, header_row):
    column_map = {}

    for cell in sheet[header_row]:
        header_text = normalize_text(cell.value)

        if not header_text:
            continue

        for col_def in COLUMNS:
            key = col_def["key"]

            if key in column_map:
                continue

            for keyword in col_def["keywords"]:
                if normalize_text(keyword) in header_text:
                    column_map[key] = cell.column
                    break

    return column_map


def get_cell_value(sheet, row_index, column_map, col_def):
    col_index = column_map.get(col_def["key"])

    if not col_index:
        return "" if col_def["type"] == "text" else 0

    value = sheet.cell(row=row_index, column=col_index).value

    if col_def["type"] == "text":
        return safe_text(value)

    return safe_number(value)


def find_summary_row(sheet, header_row, column_map):
    """
    找 Excel 上方總數列。
    邏輯：表頭上方幾列，哪一列數字最多，就當總數列。
    """

    best_row = None
    best_score = -1

    start_row = max(1, header_row - 5)
    end_row = header_row - 1

    for row_index in range(start_row, end_row + 1):
        score = 0

        for col_def in COLUMNS:
            if col_def["type"] != "number":
                continue

            col_index = column_map.get(col_def["key"])

            if not col_index:
                continue

            value = sheet.cell(row=row_index, column=col_index).value

            if safe_number(value) != 0:
                score += 1

        if score > best_score:
            best_score = score
            best_row = row_index

    return best_row


def build_summary(sheet, summary_row, column_map):
    summary = {}

    for col_def in COLUMNS:
        if col_def["type"] != "number":
            continue

        summary[col_def["key"]] = get_cell_value(
            sheet, summary_row, column_map, col_def
        )

    return summary


def build_rows(sheet, header_row, column_map):
    rows = []

    for row_index in range(header_row + 1, sheet.max_row + 1):
        work_order_col = column_map.get("work_order")
        input_qty_col = column_map.get("input_qty")

        work_order = ""
        input_qty = 0

        if work_order_col:
            work_order = safe_text(
                sheet.cell(row=row_index, column=work_order_col).value
            )

        if input_qty_col:
            input_qty = safe_number(
                sheet.cell(row=row_index, column=input_qty_col).value
            )

        if not work_order and input_qty == 0:
            continue

        row_data = {}

        for col_def in COLUMNS:
            row_data[col_def["key"]] = get_cell_value(
                sheet, row_index, column_map, col_def
            )

        rows.append(row_data)

    return rows


def convert_excel_to_json():
    if not EXCEL_FILE.exists():
        print(f"找不到 Excel 檔案：{EXCEL_FILE}")
        return

    workbook = load_workbook(EXCEL_FILE, data_only=True, read_only=True)

    if TARGET_SHEET not in workbook.sheetnames:
        print(f"找不到工作表：{TARGET_SHEET}")
        return

    sheet = workbook[TARGET_SHEET]

    header_row = find_header_row(sheet)

    if not header_row:
        print("找不到表頭列")
        return

    column_map = build_column_map(sheet, header_row)

    summary_row = find_summary_row(sheet, header_row, column_map)

    summary = build_summary(sheet, summary_row, column_map)

    rows = build_rows(sheet, header_row, column_map)

    output_data = {
        "sheet_name": TARGET_SHEET,
        "header_groups": HEADER_GROUPS,
        "columns": COLUMNS,
        "summary": summary,
        "rows": rows,
    }

    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)

    with open(OUTPUT_FILE, "w", encoding="utf-8") as file:
        json.dump(output_data, file, ensure_ascii=False, indent=2)

    print(f"工作表：{TARGET_SHEET}")
    print(f"表頭列：{header_row}")
    print(f"總數列：{summary_row}")
    print(f"資料筆數：{len(rows)}")
    print(f"輸出檔案：{OUTPUT_FILE}")


if __name__ == "__main__":
    convert_excel_to_json()
