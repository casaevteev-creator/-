# -*- coding: utf-8 -*-
"""Сборка отчёта: единый файл ввода (или канонический JSON) -> HTML-отчёт + Excel-расшифровка.

    python3 tools/build.py --input "templates/Ввод-данных-Август-2026.xlsx" --year 2026 --month 8
    python3 tools/build.py --canon data/canonical/2026-06.json
"""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import input_xlsx  # noqa: E402
import model  # noqa: E402
import report_html  # noqa: E402
import workbook  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent


def main():
    ap = argparse.ArgumentParser(description="Сборка отчёта учредителям")
    ap.add_argument("--input", help="единый файл ввода (xlsx)")
    ap.add_argument("--canon", help="готовый канонический JSON")
    ap.add_argument("--year", type=int)
    ap.add_argument("--month", type=int)
    ap.add_argument("--narrative", help="JSON с текстами резюме и фокуса (необязательно)")
    ap.add_argument("--out", default=str(ROOT / "out"))
    a = ap.parse_args()

    if a.canon:
        canon = json.loads(Path(a.canon).read_text(encoding="utf-8"))
    elif a.input:
        if not (a.year and a.month):
            ap.error("--input требует --year и --month")
        canon = input_xlsx.read(a.input, a.year, a.month)
        cpath = ROOT / "data" / "canonical" / f"{a.year}-{a.month:02d}.json"
        cpath.parent.mkdir(parents=True, exist_ok=True)
        cpath.write_text(json.dumps(canon, ensure_ascii=False, indent=1), encoding="utf-8")
        print("канонические данные:", cpath)
    else:
        ap.error("нужен --input или --canon")

    r = model.compute(canon)
    if not r["revenue"]["total"]:
        print("В файле ввода нет ни одного платежа — отчёт собирать не из чего.\n"
              "Заполните лист «1 Оплаты» (и остальные) и запустите сборку снова.")
        return 2
    narrative = json.loads(Path(a.narrative).read_text(encoding="utf-8")) if a.narrative else None

    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    tag = f"{canon['meta']['month_ru']}-{canon['meta']['year']}"
    html_path = out / f"otchet-{canon['meta']['year']}-{canon['meta']['month']:02d}.html"
    xlsx_path = out / f"Расшифровка-отчёта-{tag}.xlsx"
    html_path.write_text(report_html.render(r, narrative), encoding="utf-8")
    workbook.build(canon, r, xlsx_path)

    print("отчёт:      ", html_path)
    print("расшифровка:", xlsx_path)
    print()
    bad = [k for k in r["checks"] if not k["ok"] and "разные базы" not in k.get("note", "")]
    for k in r["checks"]:
        info = "разные базы" in k.get("note", "")
        mark = "справочно" if info else ("OK" if k["ok"] else "РАСХОЖДЕНИЕ")
        print(f"  [{mark}] {k['name']}: {k['left']:,.2f} vs {k['right']:,.2f} (Δ {k['delta']:,.2f})")
    if r["trend"].get("missing_prev") or r["trend"].get("missing_cur"):
        print("\n  ВНИМАНИЕ: нет данных за месяцы —",
              ", ".join(r["trend"]["missing_prev"] + r["trend"]["missing_cur"]))
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
