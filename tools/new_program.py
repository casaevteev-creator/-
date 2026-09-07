# -*- coding: utf-8 -*-
"""Разбор выгрузок новой управленки и сборка реестра платежей за месяц.

Читает четыре отчёта и сводит их в один реестр платежей с врачом:

  «Оплаты»                    — вид оплаты, документ (дата и номер внутри названия), клиент, сумма
  «Услуги»                    — плоский реестр: дата, номер, клиент, врач, касса, продажа/возврат
  «Взаиморасчёты по хирургам» — реализация и оплата в разрезе сотрудника
  «Кассы по документам»       — обороты по каждой ККТ

Платёж получает врача через номер документа «Оказание услуг»; платежи-авансы
(оплата картой, кассовые ордера) врача не имеют — скрипт показывает, сколько
денег осталось неразнесённым.

    python3 tools/new_program.py data/source/2026-08
"""
import datetime
import re
import sys
import tempfile
import zipfile
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl

DOC_RE = re.compile(r"^(.+?)\s+([A-ZА-Я]?\d{6,})\s+от\s+(\d{2}\.\d{2}\.\d{4})\s+(\d{1,2}:\d{2}:\d{2})$")
DT_RE = re.compile(r"(\d{2})\.(\d{2})\.(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})")
FOUNDERS = ("Маланичев", "Погосян")


def money(v):
    return f"{v:,.0f}".replace(",", " ")


def _num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def open_fixed(path):
    """1С пишет xl/SharedStrings.xml с заглавной S — openpyxl такое не открывает."""
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


def _dt(v):
    if isinstance(v, datetime.datetime):
        return v
    m = DT_RE.match(str(v).strip()) if v is not None else None
    return datetime.datetime(int(m[3]), int(m[2]), int(m[1]), int(m[4]), int(m[5]), int(m[6])) if m else None


def _indent(cell):
    return int((cell.alignment.indent if cell.alignment else 0) or 0)


# --------------------------------------------------------------------------- отчёты
def parse_payments(path):
    ws = open_fixed(path).worksheets[0]
    rows, way = [], None
    for r in range(6, ws.max_row + 1):
        label = ws.cell(row=r, column=1).value
        if label in (None, ""):
            continue
        label = str(label).strip()
        if label.lower().startswith("итог"):
            continue
        if _indent(ws.cell(row=r, column=1)) == 0:
            way = label
            continue
        m = DOC_RE.match(label)
        if not m:
            continue
        rows.append({
            "doc_type": m.group(1), "doc_num": m.group(2),
            "dt": datetime.datetime.strptime(f"{m.group(3)} {m.group(4)}", "%d.%m.%Y %H:%M:%S"),
            "client": ws.cell(row=r, column=5).value, "way": way,
            "amount": _num(ws.cell(row=r, column=6).value) or 0.0,
        })
    return rows


def parse_services(path):
    ws = open_fixed(path).worksheets[0]
    cols = {str(ws.cell(row=1, column=c).value).strip(): c
            for c in range(1, ws.max_column + 1) if ws.cell(row=1, column=c).value}

    def get(r, name):
        c = cols.get(name)
        return ws.cell(row=r, column=c).value if c else None

    rows = []
    for r in range(2, ws.max_row + 1):
        num = get(r, "Номер")
        if not num:
            continue
        rows.append({
            "dt": _dt(get(r, "Дата работ")), "num": str(num).strip(),
            "client": get(r, "Клиент"), "clinic": get(r, "Клиника"),
            "doctor": get(r, "Врач"), "op": get(r, "Вид операции"),
            "amount": _num(get(r, "Сумма")) or 0.0,
            "kassa": get(r, "Касса"), "check": get(r, "Чек пробит"),
        })
    return rows


def parse_doctors(path):
    """«Взаиморасчёты по хирургам»: на каждого врача — реализация и оплата внутри документа.

    Разница между ними — та часть услуг, которую закрыли отдельным платежом
    (карта, кассовый ордер). Врач по ней известен именно отсюда.
    """
    ws = open_fixed(path).worksheets[0]
    out, total = {}, None
    for r in range(6, ws.max_row + 1):
        label = ws.cell(row=r, column=1).value
        if label in (None, ""):
            continue
        label = str(label).strip()
        vals = [_num(ws.cell(row=r, column=c).value) or 0.0 for c in range(5, 11)]
        rec = {"debt_open": vals[0], "deposit_open": vals[1], "realization": vals[2],
               "paid_in_doc": vals[3], "debt_close": vals[4], "deposit_close": vals[5]}
        rec["paid_by_advance"] = rec["realization"] - rec["paid_in_doc"]
        if label.lower().startswith("итог"):
            total = rec
        elif _indent(ws.cell(row=r, column=1)) == 0:
            out[label] = rec
    # в строке «Итого» остатки часто не выводятся — собираем их из строк сотрудников
    for key in ("debt_open", "deposit_open", "debt_close", "deposit_close"):
        if total is not None and not total[key]:
            total[key] = sum(rec[key] for rec in out.values())
    return out, total


def parse_finresult(path):
    """«Финансовый результат»: специализация -> сотрудник -> сумма продажи.

    Даёт справочник «кто чем занимается» — из него берётся деление на хирургов,
    косметологов и прочее, которого нет ни в одном другом отчёте.
    """
    ws = open_fixed(path).worksheets[0]
    cells, spec = [], None
    for r in range(6, ws.max_row + 1):
        label = ws.cell(row=r, column=1).value
        if label in (None, ""):
            continue
        label = str(label).strip()
        if label.lower().startswith("итого"):
            continue
        sale = _num(ws.cell(row=r, column=7).value) or 0.0
        if _indent(ws.cell(row=r, column=1)) == 0:
            spec = label
        else:
            cells.append({"spec": spec, "doctor": label, "sale": sale})
    return cells


def report_group(spec, doctor):
    """Пять групп отчёта: два учредителя, другие хирурги, косметология, прочее."""
    for f in FOUNDERS:
        if f in str(doctor):
            return f
    if spec == "Косметология":
        return "Косметология"
    if spec in ("Материал", "Не указано", "Товары", "КДЛ"):
        return "Прочие доходы"
    return "Другие хирурги"


def parse_grouped(path, value_cols):
    """Отчёты вида «группа -> документы» (взаиморасчёты, кассы): итоги по группам."""
    ws = open_fixed(path).worksheets[0]
    out, total = {}, None
    for r in range(6, ws.max_row + 1):
        label = ws.cell(row=r, column=1).value
        if label in (None, ""):
            continue
        label = str(label).strip()
        vals = [_num(ws.cell(row=r, column=c).value) or 0.0 for c in value_cols]
        if label.lower().startswith("итог"):
            total = vals
        elif _indent(ws.cell(row=r, column=1)) == 0:
            out[label] = vals
    return out, total


def period_of(path):
    ws = open_fixed(path).worksheets[0]
    for row in ws.iter_rows(min_row=1, max_row=6, max_col=8):
        for c in row:
            if isinstance(c.value, str) and "Период" in c.value:
                m = re.search(r"(\d{2}\.\d{2}\.\d{4})\s*-\s*(\d{2}\.\d{2}\.\d{4})", c.value)
                if m:
                    return m.group(1), m.group(2)
    return None, None


# --------------------------------------------------------------------------- сведение
def group_of(doctor):
    for f in FOUNDERS:
        if doctor and f in str(doctor):
            return f
    return "Другие"


def enrich(payments, services):
    by_num = {s["num"]: s for s in services}
    for p in payments:
        s = by_num.get(p["doc_num"]) if p["doc_type"].startswith("Оказание услуг") else None
        p["doctor"] = s["doctor"] if s else None
        p["kassa"] = s["kassa"] if s else None
        p["clinic"] = s["clinic"] if s else None
    return payments


def report(folder):
    folder = Path(folder)

    def find(*keys):
        for f in sorted(folder.glob("*.xlsx")):
            if all(k.lower() in f.name.lower() for k in keys):
                return f
        return None

    p_pay, p_srv = find("оплаты"), find("услуги")
    p_doc, p_cash = find("хирург"), find("кассы")
    if not (p_pay and p_srv):
        print("не найдены файлы «Оплаты» и «Услуги» в", folder)
        return

    payments = enrich(parse_payments(p_pay), parse_services(p_srv))
    services = parse_services(p_srv)
    print(f"«Оплаты»: период {' — '.join(x or '?' for x in period_of(p_pay))}, платежей: {len(payments)}")
    print(f"«Услуги»: строк: {len(services)}, "
          f"даты {min(s['dt'] for s in services if s['dt']).date()} — "
          f"{max(s['dt'] for s in services if s['dt']).date()}")
    print()

    # --- выручка по дням и каналам
    byday = defaultdict(lambda: defaultdict(float))
    for p in payments:
        byday[p["dt"].date()][p["way"]] += p["amount"]
    ways = sorted({p["way"] for p in payments})
    print("ПЛАТЕЖИ ПО ДНЯМ")
    print("  дата         " + "".join(f"{w[:14]:>16}" for w in ways) + f"{'всего':>16}")
    totals = defaultdict(float)
    for d in sorted(byday):
        line = "".join(f"{money(byday[d].get(w, 0)):>16}" for w in ways)
        s = sum(byday[d].values())
        for w in ways:
            totals[w] += byday[d].get(w, 0)
        print(f"  {d}" + line + f"{money(s):>16}")
    print(f"  {'ИТОГО':10s}" + "".join(f"{money(totals[w]):>16}" for w in ways)
          + f"{money(sum(totals.values())):>16}")
    print()

    # --- разнесение по врачам
    with_doc = [p for p in payments if p["doctor"]]
    without = [p for p in payments if not p["doctor"]]
    s_all = sum(p["amount"] for p in payments)
    s_doc = sum(p["amount"] for p in with_doc)
    print("РАЗНЕСЕНИЕ ПЛАТЕЖЕЙ ПО ВРАЧАМ")
    print(f"  всего платежей:            {money(s_all):>16}")
    print(f"  привязано к врачу:         {money(s_doc):>16}  ({s_doc/s_all*100:.1f}%)")
    print(f"  без врача (авансы и т.п.): {money(s_all-s_doc):>16}  ({(s_all-s_doc)/s_all*100:.1f}%)")
    for t, c in Counter(p["doc_type"] for p in without).most_common():
        amt = sum(p["amount"] for p in without if p["doc_type"] == t)
        print(f"      {t[:38]:40s} {money(amt):>14}  строк {c}")
    grp = defaultdict(float)
    for p in with_doc:
        grp[group_of(p["doctor"])] += p["amount"]
    print("  по группам (только по документам услуг; полный разрез — ниже, по реализации):")
    for k, v in sorted(grp.items(), key=lambda x: -x[1]):
        print(f"      {k:40s} {money(v):>14}")
    print()

    # --- сверки с другими отчётами
    residual = None
    if p_doc:
        emp, total = parse_doctors(p_doc)
        print("ВЫРУЧКА ПО ВРАЧАМ («Взаиморасчёты по хирургам»)")
        print(f"  {'врач':34s}{'реализация':>15}{'оплата в док-те':>17}{'закрыто авансом':>17}")
        grp2 = defaultdict(lambda: defaultdict(float))
        for name, rec in sorted(emp.items(), key=lambda x: -x[1]["realization"]):
            g = group_of(name)
            for k in ("realization", "paid_in_doc", "paid_by_advance"):
                grp2[g][k] += rec[k]
            if rec["realization"] >= 500_000:
                print(f"  {name[:34]:34s}{money(rec['realization']):>15}"
                      f"{money(rec['paid_in_doc']):>17}{money(rec['paid_by_advance']):>17}")
        print(f"  {'ИТОГО':34s}{money(total['realization']):>15}"
              f"{money(total['paid_in_doc']):>17}{money(total['paid_by_advance']):>17}")
        print("  по группам:")
        for g, v in sorted(grp2.items(), key=lambda x: -x[1]["realization"]):
            print(f"      {g:12s} реализация {money(v['realization']):>14}"
                  f"   в т.ч. закрыто авансом {money(v['paid_by_advance']):>14}")
        print()
        print("СВЕРКА ДЕНЕГ И УСЛУГ")
        print(f"  оплата внутри документов услуг:      {money(total['paid_in_doc']):>16}"
              f"   (из «Оплат»: {money(s_doc)}, Δ {money(s_doc - total['paid_in_doc'])})")
        adv = s_all - s_doc
        residual = adv - total["paid_by_advance"]
        print(f"  отдельные платёжные документы:       {money(adv):>16}")
        print(f"  из них закрыли реализацию по врачам: {money(total['paid_by_advance']):>16}")
        print(f"  остаток без объяснения:              {money(residual):>16}")
        print(f"  депозиты клиентов: начало {money(total['deposit_open'])} -> "
              f"конец {money(total['deposit_close'])}; "
              f"долг: {money(total['debt_open'])} -> {money(total['debt_close'])}")
        print()
    if p_cash:
        desks, total = parse_grouped(p_cash, [5, 6, 7, 8])
        print("КАССЫ ПО ДОКУМЕНТАМ")
        for k, v in desks.items():
            print(f"  {k[:34]:36s} приход {money(v[1]):>14}  расход {money(v[2]):>14}  остаток {money(v[3]):>12}")
        print()

    # --- чего не хватает
    print("ЧЕГО НЕ ХВАТАЕТ ДЛЯ ОТЧЁТА")
    first = min(p["dt"].date() for p in payments)
    if first.day > 1:
        print(f"  · платежей за 01–{first.day:02d} число нет — нужна такая же выгрузка «Оплаты» из старой программы")
    if residual is not None and abs(residual) > 1000:
        print(f"  · {money(residual)} ₽ денег не объясняются реализацией по врачам — "
              "нужна методология: авансы под услуги следующего месяца, депозиты или что-то ещё")
    elif without:
        print(f"  · {money(s_all-s_doc)} ₽ платежей без врача — нужен отчёт «Оплаты» с полем «Сотрудник»")
    if len(ways) < 3:
        print("  · «Безналичными» одной строкой — нужно понять, входят ли туда оплаты на р/с "
              "или это только эквайринг")
    return payments


if __name__ == "__main__":
    report(sys.argv[1] if len(sys.argv) > 1 else "data/source/2026-08")
