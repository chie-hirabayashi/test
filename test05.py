import os
import openpyxl as excel
import sqlite3

# 各シートのDB対応
excel_db_map = {
    "WEBサイト": {
        "min_row": 2,
        "table_name": "web",
        "column": {
            "url": {"type": "text", "index": 1},
            "div_name": {"type": "text", "index": 2},
        },
    },
    "人口": {
        "min_row": 2,
        "table_name": "population",
        "column": {
            "div_code": {"type": "text", "index": 1},
            "div_name": {"type": "text", "index": 2},
            "town_name": {"type": "text", "index": 3},
            "town_kana": {"type": "text", "index": 4},
            "households": {"type": "int", "index": 5},
            "man": {"type": "int", "index": 6},
            "woman": {"type": "int", "index": 7},
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
    dbname = os.path.join(current_dir, "saitama3.db")
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
    "/Users/admin/camp/create/saitama_city3.xlsx", data_only=True
)
db_insert(wb, excel_db_map)
