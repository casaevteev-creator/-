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


def parse_patient_payments(path):
    """«Оплаты от пациента» старой программы — реестр платежей с видом оплаты.

    Колонки: Дата, Номер, Пациент, Специалист, Бух учет, Касса, Сумма руб,
    Оплата плат. картой, Оплата без. нал. Наличные — остаток от суммы.
    """
    import re as _re
    ws = open_fixed(path).worksheets[0]
    hdr = {str(ws.cell(row=1, column=c).value or "").strip(): c
           for c in range(1, ws.max_column + 1)}
    need = ("Дата", "Сумма руб", "Оплата плат. картой", "Оплата без. нал.")
    if not all(k in hdr for k in need):
        raise ValueError("это не «Оплаты от пациента»: нет колонок " +
                         ", ".join(k for k in need if k not in hdr))

    def dt(v):
        if isinstance(v, datetime.datetime):
            return v
        m = _re.match(r"(\d{2})\.(\d{2})\.(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})",
                      str(v).strip()) if v is not None else None
        return (datetime.datetime(int(m[3]), int(m[2]), int(m[1]),
                                  int(m[4]), int(m[5]), int(m[6])) if m else None)

    rows = []
    for r in range(2, ws.max_row + 1):
        d = dt(ws.cell(row=r, column=hdr["Дата"]).value)
        if not d:
            continue
        total = _num(ws.cell(row=r, column=hdr["Сумма руб"]).value) or 0.0
        card = _num(ws.cell(row=r, column=hdr["Оплата плат. картой"]).value) or 0.0
        bank = _num(ws.cell(row=r, column=hdr["Оплата без. нал."]).value) or 0.0
        rows.append({
            "dt": d, "number": ws.cell(row=r, column=hdr.get("Номер", 1)).value,
            "patient": ws.cell(row=r, column=hdr.get("Пациент", 1)).value,
            "doctor": str(ws.cell(row=r, column=hdr.get("Специалист", 1)).value or "").strip(),
            "kassa": ws.cell(row=r, column=hdr.get("Касса", 1)).value,
            "total": total, "card": card, "bank": bank, "cash": round(total - card - bank, 2),
        })
    return rows


def parse_daily_cash(path):
    """«Ежедневный отчет по кассам» — единственный источник возвратов первой декады.

    Шапка двухэтажная: строка «Касса | Период | Пациент | … | Специалист | Вид услуги»
    и над ней «Денежные средства» с колонками ПОЛУЧЕНО / касса / Б/нал / кред.карта /
    Возврат денег. Строки без даты в колонке «Период» — итоги по кассе.
    """
    ws = open_fixed(path).worksheets[0]
    hrow = _header_row(ws, "ПОЛУЧЕНО", limit=20)
    if not hrow:
        raise ValueError("это не «Ежедневный отчет по кассам»: нет колонки «ПОЛУЧЕНО»")
    cols = {str(ws.cell(row=hrow, column=c).value).strip(): c
            for c in range(1, ws.max_column + 1) if ws.cell(row=hrow, column=c).value}

    def g(r, name):
        c = cols.get(name)
        return ws.cell(row=r, column=c).value if c else None

    rows, totals = [], {}
    for r in range(hrow + 1, ws.max_row + 1):
        period = g(r, "Период")
        label = str(period).strip() if period is not None else ""
        got = _num(g(r, "ПОЛУЧЕНО")) or 0.0
        back = _num(g(r, "Возврат денег")) or 0.0
        if label and not label[0].isdigit():          # строка-итог по кассе или «Итог»
            totals[label] = {"got": got, "refund": back,
                             "cash": _num(g(r, "В ТОМ ЧИСЛЕ: касса")) or 0.0,
                             "bank": _num(g(r, "В ТОМ ЧИСЛЕ: Б/нал")) or 0.0,
                             "card": _num(g(r, "В ТОМ ЧИСЛЕ: кред.карта")) or 0.0,
                             "services": _num(g(r, "Оказано услуг с учетом скидки")) or 0.0}
            continue
        d = _date(label)
        if d is None or not (got or back):
            continue
        rows.append({"date": d, "patient": g(r, "Пациент"),
                     "doctor": str(g(r, "Специалист") or "").strip(),
                     "service": g(r, "Вид услуги"), "got": got, "refund": back})
    return rows, totals
