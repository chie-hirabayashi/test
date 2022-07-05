import os
import openpyxl as excel
import sqlite3

# 各シートのDB対応
excel_db_map = {
    # "day": {  # エクセルのテーブル名
    #     "min_row": 2,  # テーブルの数？？
    #     "table_name": "day",  # DBのテーブル名
    #     "column": {  # カラムのタイプ指定
    #         "no.": {"type": "text", "index": 1},
    #         "add_day": {"type": "text", "index": 2},
    #         "born_day1": {"type": "text", "index": 3},
    #         "born_day2": {"type": "text", "index": 4},
    #         "born_day3": {"type": "text", "index": 5},
    #         "born_day4": {"type": "text", "index": 6},
    #         "born_day5": {"type": "text", "index": 7},
    #         "born_day6": {"type": "text", "index": 8},
    #         "born_day7": {"type": "text", "index": 9},
    #         "born_day8": {"type": "text", "index": 10},
    #         "born_day9": {"type": "text", "index": 11},
    #         "born_day10": {"type": "text", "index": 12},
    #         "born_day11": {"type": "text", "index": 13},
    #         "born_day12": {"type": "text", "index": 14},
    #         "delete_day": {"type": "text", "index": 15},
    #     },
    # },
    "number": {
        "min_row": 2,
        "table_name": "number",
        "column": {
            "no.": {"type": "text", "index": 1},
            "born_num1": {"type": "int", "index": 2},
            "born_num2": {"type": "int", "index": 3},
            "born_num3": {"type": "int", "index": 4},
            "born_num4": {"type": "int", "index": 5},
            "born_num5": {"type": "int", "index": 6},
            "born_num6": {"type": "int", "index": 7},
            "born_num7": {"type": "int", "index": 8},
            "born_num8": {"type": "int", "index": 9},
            "born_num9": {"type": "int", "index": 10},
            "born_num10": {"type": "int", "index": 11},
            "born_num11": {"type": "int", "index": 12},
            "born_num12": {"type": "int", "index": 13},
            # "delete_day": {"type": "text", "index": 14},
        },
    },
}


def db_init(db_map, db, sheet_name):
    param = []
    table_name = db_map[sheet_name]["table_name"]

    for k, v in db_map[sheet_name]["column"].items():
        param.append(f"{k} {v['type']}")

    params = ",".join(param)

    db.execute(f"CREATE TABLE IF NOT EXISTS {table_name}({params})")
    db.execute(f"DELETE FROM {table_name}")


def db_insert(book, db_map):
    current_dir = os.path.dirname(__file__)
    dbname = os.path.join(current_dir, "pig.db")  # DBの名前
    conn = sqlite3.connect(dbname)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    for sheet_name in db_map:
        db_init(db_map, cur, sheet_name)
        sheet = book[sheet_name]
        col_name = []
        val = []
        min_row = db_map[sheet_name]["min_row"]
        table_name = db_map[sheet_name]["table_name"]

        for k, v in db_map[sheet_name]["column"].items():
            col_name.append(k)
            val.append(v["index"])

        col_names = ",".join(col_name)

        for r in sheet.iter_rows(min_row=min_row):
            values = []

            if r[0].value is None:
                break

            for v in val:
                cell_val = r[v - 1].value

                if (
                    type(cell_val) is not str
                    and type(cell_val) is not int
                    and cell_val is not None
                ):
                    values.append(str(cell_val))
                else:
                    values.append(cell_val)

            place_holder = ",".join("?" * len(values))
            sql = f"INSERT INTO {table_name} ({col_names}) VALUES({place_holder})"
            cur.execute(sql, tuple(values))

    conn.commit()
    cur.close()
    conn.close()


wb = excel.load_workbook(
    "/Users/admin/camp/create/db.xlsx", data_only=True
)  # エクセルデータの保存場所
db_insert(wb, excel_db_map)
