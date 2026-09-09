# -*- coding: utf-8 -*-
"""Классификация контрагентов счёта 60 без ручной разметки.

Три источника, в порядке приоритета:
  1. фамилия врача из управленки  -> гонорары ИП хирургов;
  2. классификатор прошлого месяца (файл «Расходы» с назначениями платежа);
  3. корреспондирующие счета кредитового оборота 60 — что от контрагента
     приходило: 10 материалы, 41 товары, 26 общехоз, 08/20 капвложения и работы.

Остаётся горстка действительно новых контрагентов — их печатаем списком.

    python3 tools/vendors.py data/source/2026-08/1С-Обороты-счета-60-август-2026.xlsx
"""
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from june_source import parse_expenses  # noqa: E402
from new_program import _num, money, open_fixed, parse_finresult  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
PAY_COL = 5                      # «Оборот Дт / 51» — оплачено с расчётного счёта
CORR = {"08": 9, "10": 10, "19": 11, "20": 12, "26": 13, "41": 14, "76": 17}


def read_turnover(path):
    """Контрагенты из «Оборотов счета 60».

    Уровни отчёта различаются отступом первой колонки: 0 — счёт и «Итого»,
    2 — подразделение, 4 — контрагент. Подразделений может быть несколько
    (в августе платёж Квалитету с кредитной линии попал в отдельное «<...>»),
    поэтому один контрагент встречается дважды и суммируется.
    """
    ws = open_fixed(path).worksheets[0]
    out = {}
    for r in range(9, ws.max_row + 1):
        cell = ws.cell(row=r, column=1)
        name = cell.value
        if not name:
            continue
        indent = int((cell.alignment.indent if cell.alignment else 0) or 0)
        if indent < 4:                      # счёт, подразделение, «Итого»
            continue
        n = " ".join(str(name).split())
        pay = _num(ws.cell(row=r, column=PAY_COL).value) or 0.0
        if not pay:
            continue
        prev_pay, prev_corr = out.get(n, (0.0, {k: 0.0 for k in CORR}))
        corr = {k: prev_corr[k] + (_num(ws.cell(row=r, column=c).value) or 0.0)
                for k, c in CORR.items()}
        out[n] = (round(prev_pay + pay, 2), corr)
    return out


def load_notes():
    """Ручные пометки по новым контрагентам: data/vendor_notes.json."""
    import json
    f = ROOT / "data" / "vendor_notes.json"
    if not f.exists():
        return {}
    raw = json.loads(f.read_text(encoding="utf-8"))
    return {" ".join(k.split()).lower(): v for k, v in raw.items() if not k.startswith("_")}


def classify(turnover, doctors=None, prev_month=None, notes=None):
    lastnames = {d.split()[0].lower() for d in (doctors or set()) if d.split()}
    prev = {" ".join(v["name"].split()).lower(): v.get("note", "") for v in (prev_month or [])}
    notes = notes or {}
    buckets, unknown = defaultdict(float), {}
    detail = defaultdict(list)
    for name, (pay, corr) in turnover.items():
        low = " ".join(name.split()).lower()
        first = low.split()[0].strip('"')
        if first in lastnames:
            key = "Гонорары врачей (ИП)"
        elif low in notes:
            key = "Ручные пометки: " + notes[low].get("bucket", "управленческие")
        elif prev.get(low):
            key = "Классификатор прошлого месяца"
        elif corr["10"] > 0:
            key = "Медицинские материалы (сч. 10)"
        elif corr["41"] > 0:
            key = "Товары (сч. 41)"
        elif corr["26"] > 0:
            key = "Общехозяйственные (сч. 26)"
        elif corr["08"] > 0 or corr["20"] > 0:
            key = "Капвложения и работы (сч. 08, 20)"
        elif low in prev:
            key = "Классификатор прошлого месяца"
        else:
            unknown[name] = pay
            continue
        buckets[key] += pay
        detail[key].append((name, pay))
    return buckets, unknown, detail


def main(path):
    turnover = read_turnover(path)
    doctors = {c["doctor"] for c in parse_finresult(
        ROOT / "data/source/2026-08/Новая-Финрезультат-10-31.08.2026.xlsx")}
    prev = parse_expenses(ROOT / "data/source/Расходы-июнь-2026.xls")["vendors"]
    buckets, unknown, _ = classify(turnover, doctors, prev, load_notes())

    total = sum(p for p, _ in turnover.values())
    print(f"оплачено с расчётного счёта: {money(total)} ₽, контрагентов {len(turnover)}")
    for k, v in sorted(buckets.items(), key=lambda x: -x[1]):
        print(f"  {k:34s} {money(v):>13}   {v/total*100:5.1f}%")
    rest = sum(unknown.values())
    print(f"  {'НЕ ОПОЗНАНО':34s} {money(rest):>13}   {rest/total*100:5.1f}%   ({len(unknown)} контрагентов)")
    if unknown:
        print("\nтребуют уточнения:")
        for k, v in sorted(unknown.items(), key=lambda x: -x[1]):
            print(f"  {k[:48]:50s} {money(v):>12}")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1
         else ROOT / "data/source/2026-08/1С-Обороты-счета-60-август-2026.xlsx")
