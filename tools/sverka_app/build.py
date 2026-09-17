# -*- coding: utf-8 -*-
"""Сборка приложения «Сверка Такском ↔ УМЦ».

    python3 tools/sverka_app/build.py    -> out/sverka.html

Один файл: открывается двойным щелчком, работает без интернета, ничего
никуда не отправляет. SheetJS, стили и разборщики выгрузок вшиваются внутрь.
Разборщики и стили общие с приложением обмена — лежат в tools/lib.
"""
import argparse
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SRC = Path(__file__).resolve().parent / "index.html"
LIBDIR = ROOT / "tools" / "lib"
LIB = ROOT / "node_modules" / "xlsx" / "dist" / "xlsx.full.min.js"

ЧАСТИ = {
    "__СТИЛЬ__": LIBDIR / "app.css",
    "__ОСНОВА__": LIBDIR / "utils.js",
    "__РАЗБОРЩИКИ__": LIBDIR / "umc_parsers.js",
    "__СВЕРКА__": LIBDIR / "sverka.js",
}

HEAD = """<!doctype html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<style>html{color-scheme:light}body{margin:0}img{max-width:100%}[hidden]{display:none!important}</style>
"""


def собрать():
    html = SRC.read_text(encoding="utf-8")
    for метка, путь in ЧАСТИ.items():
        if метка not in html:
            raise SystemExit(f"в index.html нет метки {метка}")
        html = html.replace(метка, путь.read_text(encoding="utf-8"))
    if not LIB.exists():
        raise SystemExit(f"нет SheetJS: {LIB}\nпоставьте его: npm install xlsx@0.18.5 --no-save")
    html = html.replace(
        "__SHEETJS__",
        "/* SheetJS 0.18.5, Apache-2.0 — вшит в файл, чтобы работало без интернета */\n"
        + LIB.read_text(encoding="utf-8"))
    head, rest = html.split("</style>", 1)
    return HEAD + head + "</style>\n</head>\n<body>\n" + rest + "\n</body>\n</html>\n"


def main():
    ap = argparse.ArgumentParser(description="Сборка приложения сверки")
    ap.add_argument("--out", default="sverka.html", help="имя файла в out/")
    a = ap.parse_args()
    out = ROOT / "out" / a.out
    doc = собрать()
    out.parent.mkdir(exist_ok=True)
    out.write_text(doc, encoding="utf-8")
    print(f"собрано: {out}  ({len(doc):,} символов)".replace(",", " "))


if __name__ == "__main__":
    main()
