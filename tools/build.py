#!/usr/bin/env python3
"""Собирает index.html из src/template.html и таблицы Excel.

Запуск из корня проекта:
    python3 tools/build.py data/plan.xlsm "Мероприятие 24–28 ноября"
Нужен openpyxl:  pip3 install openpyxl

Веб-версия для выкладки (без вшитой таблицы, данные только из Google) — в папку site/:
    python3 tools/build.py --web
"""
import datetime, json, sys, pathlib

def conv(v):
    if isinstance(v, datetime.datetime):
        return (v - datetime.datetime(1899, 12, 30)).total_seconds() / 86400
    if isinstance(v, datetime.time):
        return (v.hour * 3600 + v.minute * 60 + v.second) / 86400
    return v

def main():
    if len(sys.argv) > 1 and sys.argv[1] == "--web":
        tpl = pathlib.Path("src/template.html").read_text(encoding="utf-8")
        out = pathlib.Path("site"); out.mkdir(exist_ok=True)
        (out / "index.html").write_text(tpl.replace("__RAW__", "null").replace("__NAME__", ""), encoding="utf-8")
        print("site/index.html собран (веб-версия без вшитой таблицы)")
        return
    xlsx = sys.argv[1] if len(sys.argv) > 1 else "data/plan.xlsm"
    name = sys.argv[2] if len(sys.argv) > 2 else pathlib.Path(xlsx).stem
    import openpyxl
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    sheets = {}
    for ws in wb.worksheets:
        rows = [[conv(v) for v in r] for r in ws.iter_rows(max_col=min(ws.max_column, 12), values_only=True)]
        while rows and not any(x is not None for x in rows[-1]):
            rows.pop()
        sheets[ws.title] = rows
    tpl = pathlib.Path("src/template.html").read_text(encoding="utf-8")
    html = tpl.replace("__RAW__", json.dumps(sheets, ensure_ascii=False)).replace("__NAME__", name)
    pathlib.Path("index.html").write_text(html, encoding="utf-8")
    print("index.html собран:", xlsx, "→", name)

if __name__ == "__main__":
    main()
