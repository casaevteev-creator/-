# -*- coding: utf-8 -*-
"""Разбор выгрузок старой управленки.

«Отчет по врачам (УК)» — плоская таблица с колонкой «Фактическая оплата»:
направление, пациент, врач, процедура, дата, выручка, оплата. Это и есть
источник денег за первую декаду месяца.
"""
import datetime
import re

from new_program import _num, open_fixed

DATE_RE = re.compile(r"(\d{2})\.(\d{2})\.(\d{4})")


def _date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    m = DATE_RE.match(str(v).strip()) if v is not None else None
    return datetime.date(int(m[3]), int(m[2]), int(m[1])) if m else None


def _header_row(ws, needle="Фактическая оплата", limit=15):
    for r in range(1, min(limit, ws.max_row) + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and v.strip() == needle:
                return r
    return None


def parse_doctor_report(path):
    """Возвращает список строк: направление, пациент, врач, дата, выручка, оплата."""
    ws = open_fixed(path).worksheets[0]
    hrow = _header_row(ws)
    if not hrow:
        raise ValueError("не найдена колонка «Фактическая оплата» — это другой отчёт")
    cols = {str(ws.cell(row=hrow, column=c).value).strip(): c
            for c in range(1, ws.max_column + 1) if ws.cell(row=hrow, column=c).value}

    def g(r, name):
        c = cols.get(name)
        return ws.cell(row=r, column=c).value if c else None

    rows = []
    for r in range(hrow + 1, ws.max_row + 1):
        direction = g(r, "Направление")
        if isinstance(direction, str) and direction.strip().lower().startswith("итог"):
            continue
        rev = _num(g(r, "Сумма выручки")) or 0.0
        paid = _num(g(r, "Фактическая оплата")) or 0.0
        dt = _date(g(r, "Дата"))
        if dt is None and not (rev or paid):
            continue
        rows.append({
            "direction": (str(direction).strip() if direction else ""),
            "patient": g(r, "Пациент"), "doctor": g(r, "Врач"),
            "service": g(r, "Процедура / Товар"), "date": dt,
            "revenue": rev, "paid": paid,
        })
    return rows
