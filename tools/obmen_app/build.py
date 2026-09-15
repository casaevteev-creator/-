# -*- coding: utf-8 -*-
"""Сборка приложения «Обмен УМЦ → Бухгалтерия» в один файл.

Кладёт внутрь SheetJS, чтобы страница работала без интернета: бухгалтер
открывает её двойным щелчком рядом с 1С, и никакие файлы никуда не уходят.

    python3 tools/obmen_app/build.py
"""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SRC = Path(__file__).resolve().parent / "index.html"
OUT = ROOT / "out" / "obmen.html"
LIB = ROOT / "node_modules" / "xlsx" / "dist" / "xlsx.full.min.js"

HEAD = """<!doctype html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<style>html{color-scheme:light}body{margin:0}img{max-width:100%}[hidden]{display:none!important}</style>
"""


def main():
    html = SRC.read_text(encoding="utf-8")
    if not LIB.exists():
        raise SystemExit(f"нет SheetJS: {LIB}\nпоставьте его: npm install xlsx@0.18.5 --no-save")
    lib = LIB.read_text(encoding="utf-8")
    html = html.replace(
        "__SHEETJS__",
        "/* SheetJS 0.18.5, Apache-2.0 — вшит в файл, чтобы работало без интернета */\n" + lib)

    # заголовок и стили уезжают в <head>, остальное — в <body>
    head_part, rest = html.split("</style>", 1)
    doc = HEAD + head_part + "</style>\n</head>\n<body>\n" + rest + "\n</body>\n</html>\n"
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text(doc, encoding="utf-8")
    print(f"собрано: {OUT}  ({len(doc):,} символов)".replace(",", " "))


if __name__ == "__main__":
    main()
