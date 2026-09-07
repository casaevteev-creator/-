# -*- coding: utf-8 -*-
"""Разбор выгрузок из управленки (старая и новая программы) и проверка их пригодности.

Понимает две формы:
  * «Отчёт по выручке»  — Направление / Специалист / Пациент / Процедура (старая программа);
  * «Валовая прибыль»   — Специализация / Сотрудник, колонка «Сумма продажи» (новая программа).

Считает итоги по направлениям и по группам врачей и печатает, чего в выгрузке
не хватает для сборки отчёта (дат, способа оплаты, возвратов).

    python3 tools/upravlenka.py data/source/2026-08/*.xlsx
"""
import re
import shutil
import sys
import tempfile
import zipfile
from collections import defaultdict
from pathlib import Path

import openpyxl

FOUNDERS = ("Маланичев", "Погосян")


def money(v):
    return f"{v:,.0f}".replace(",", " ") if isinstance(v, (int, float)) else "—"


def _num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def _open(path):
    """1С иногда пишет xl/SharedStrings.xml с заглавной S — openpyxl такое не открывает."""
    try:
        return openpyxl.load_workbook(path, data_only=True)
    except KeyError:
        src = zipfile.ZipFile(path)
        tmp = Path(tempfile.mkdtemp()) / "fixed.xlsx"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
            for item in src.infolist():
                out.writestr(item.filename.replace("SharedStrings", "sharedStrings"),
                             src.read(item.filename))
        return openpyxl.load_workbook(tmp, data_only=True)


def _indent(cell):
    return int((cell.alignment.indent if cell.alignment else 0) or 0)


def _period(ws):
    for row in ws.iter_rows(min_row=1, max_row=6, max_col=8):
        for c in row:
            if isinstance(c.value, str) and "Период" in c.value:
                m = re.search(r"(\d{2}\.\d{2}\.\d{4})\s*-\s*(\d{2}\.\d{2}\.\d{4})", c.value)
                if m:
                    return m.group(1), m.group(2)
    return None, None


def _group(name):
    for f in FOUNDERS:
        if f in name:
            return f
    return "Другие хирурги / прочие"


def parse(path):
    wb = _open(path)
    ws = wb.worksheets[0]
    period = _period(ws)
    head = " ".join(str(ws.cell(row=r, column=c).value or "")
                    for r in range(1, 7) for c in range(1, 12))

    if "Сумма продажи" in head:              # новая программа
        kind, name_col, sum_col, start = "валовая прибыль", 1, 7, 6
    else:                                    # старая программа
        kind, name_col, sum_col, start = "отчёт по выручке", 2, 4, 7

    # уровни вложенности задаются отступом; у разных выгрузок он начинается по-разному
    indents = sorted({_indent(ws.cell(row=r, column=name_col))
                      for r in range(start, ws.max_row + 1)
                      if ws.cell(row=r, column=name_col).value not in (None, "")})
    top = indents[0] if indents else 0
    sub = indents[1] if len(indents) > 1 else top + 1

    directions, specialists, cur = defaultdict(float), defaultdict(float), None
    rows_total = 0
    for r in range(start, ws.max_row + 1):
        cell = ws.cell(row=r, column=name_col)
        name = cell.value
        s = _num(ws.cell(row=r, column=sum_col).value)
        if name in (None, "") and s is None:
            continue
        label = str(name).strip() if name else "(без направления)"
        if label.lower().startswith("итог"):
            break
        level = _indent(cell)
        if level == top:
            cur = label
            directions[cur] += s or 0.0
        elif level == sub and cur:
            specialists[label] += s or 0.0
        rows_total += 1

    groups = defaultdict(float)
    for n, v in specialists.items():
        groups[_group(n)] += v

    return {
        "file": Path(path).name, "kind": kind, "period": period,
        "directions": dict(directions), "specialists": dict(specialists),
        "groups": dict(groups), "total": sum(directions.values()),
        "has_dates": bool(re.search(r"дат", head, re.I)),
        "has_payment": bool(re.search(r"оплат|касс|эквайр|нал\b", head, re.I)),
    }


def report(paths):
    parsed = [parse(p) for p in paths]
    for p in parsed:
        print("=" * 78)
        print(f"{p['file']}")
        print(f"  тип выгрузки: {p['kind']}   период: {p['period'][0]} — {p['period'][1]}")
        print(f"  итого: {money(p['total'])} ₽")
        print("  по направлениям:")
        for k, v in sorted(p["directions"].items(), key=lambda x: -x[1]):
            if v:
                print(f"    {k[:38]:40s} {money(v)}")
        print("  по группам:")
        for k, v in sorted(p["groups"].items(), key=lambda x: -x[1]):
            print(f"    {k[:38]:40s} {money(v)}")
        missing = []
        if not p["has_dates"]:
            missing.append("даты платежей (нет графика по дням)")
        if not p["has_payment"]:
            missing.append("способ оплаты (нет разбивки касса / эквайринг / р/с)")
        if missing:
            print("  НЕ ХВАТАЕТ: " + "; ".join(missing))

    print("=" * 78)
    print(f"НАИВНАЯ СУММА ВЫГРУЗОК: {money(sum(p['total'] for p in parsed))} ₽")
    ends = [p["period"][1] for p in parsed if p["period"][1]]
    starts = [p["period"][0] for p in parsed if p["period"][0]]
    overlap = sorted(set(ends) & set(starts))
    if overlap:
        print(f"ВНИМАНИЕ: периоды пересекаются по дате {', '.join(overlap)} — "
              "этот день входит в обе выгрузки и посчитан дважды.")
        print("Схлопнуть его нечем: в выгрузках нет ни дат, ни номеров платежей.")
    return parsed


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    report(sys.argv[1:])
