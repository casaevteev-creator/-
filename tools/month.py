# -*- coding: utf-8 -*-
"""Сводит месяц из двух управленок в одну картину по деньгам.

Первая часть месяца — старая программа («Отчет по врачам (УК)», колонка
«Фактическая оплата»), вторая — новая («Оплаты» + «Взаиморасчёты по хирургам»).

    python3 tools/month.py data/source/2026-08 --switch 2026-08-11
"""
import argparse
import datetime
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from new_program import (FOUNDERS, group_of, money, parse_doctors,  # noqa: E402
                         parse_payments, parse_services, enrich)
from old_program import parse_doctor_report  # noqa: E402


def find(folder, *keys):
    for f in sorted(Path(folder).glob("*.xlsx")):
        if all(k.lower() in f.name.lower() for k in keys):
            return f
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("folder")
    ap.add_argument("--switch", required=True, help="первый день, который считаем по новой программе (ГГГГ-ММ-ДД)")
    a = ap.parse_args()
    switch = datetime.date.fromisoformat(a.switch)
    year, month = switch.year, switch.month

    f_old = find(a.folder, "старая", "врачам")
    f_pay = find(a.folder, "оплаты")
    f_srv = find(a.folder, "услуги")
    f_doc = find(a.folder, "хирург")
    missing = [n for n, f in (("«Отчет по врачам (УК)» старой программы", f_old),
                              ("«Оплаты» новой программы", f_pay),
                              ("«Услуги» новой программы", f_srv)) if not f]
    if missing:
        print("не найдено:", "; ".join(missing))
        return 1

    # ---------------------------------------------------------- старая часть
    old = [r for r in parse_doctor_report(f_old)
           if r["date"] and r["date"].year == year and r["date"].month == month]
    old_in = [r for r in old if r["date"] < switch]
    old_after = [r for r in old if r["date"] >= switch]

    # ---------------------------------------------------------- новая часть
    pays = enrich(parse_payments(f_pay), parse_services(f_srv))
    new_in = [p for p in pays if p["dt"].date() >= switch]
    new_before = [p for p in pays if p["dt"].date() < switch]

    old_sum = sum(r["paid"] for r in old_in)
    new_sum = sum(p["amount"] for p in new_in)
    print(f"ДЕНЬГИ ЗА {month:02d}.{year}")
    print(f"  01–{(switch - datetime.timedelta(days=1)).day:02d} · старая программа "
          f"(фактическая оплата): {money(old_sum):>16}")
    print(f"  {switch.day:02d}–конец · новая программа (оплаты):      {money(new_sum):>16}")
    print(f"  ИТОГО ЗА МЕСЯЦ:                                {money(old_sum + new_sum):>16}")
    print()
    if new_before or old_after:
        print("ЗОНА ПЕРЕСЕЧЕНИЯ (проверить на задвоение)")
        if new_before:
            print(f"  новая программа до {switch}: {money(sum(p['amount'] for p in new_before))} "
                  f"({len(new_before)} платежей)")
            for p in sorted(new_before, key=lambda x: x["dt"]):
                print(f"      {p['dt']:%d.%m %H:%M}  {p['doc_type'][:28]:30s} {money(p['amount']):>12}  "
                      f"{str(p['client'])[:26]}")
        if old_after:
            print(f"  старая программа с {switch}: выручка "
                  f"{money(sum(r['revenue'] for r in old_after))}, "
                  f"оплата {money(sum(r['paid'] for r in old_after))} ({len(old_after)} строк) "
                  f"— как правило это возвраты по услугам первой декады")
            for r in sorted(old_after, key=lambda x: x["date"]):
                print(f"      {r['date']}  {r['direction'][:14]:16s} {str(r['doctor'])[:24]:26s} "
                      f"выручка {money(r['revenue']):>11}  оплата {money(r['paid']):>11}")
        print()

    # ---------------------------------------------------------- по дням
    byday = defaultdict(float)
    for r in old_in:
        byday[r["date"]] += r["paid"]
    for p in new_in:
        byday[p["dt"].date()] += p["amount"]
    print("ПО ДНЯМ")
    for d in sorted(byday):
        src = "старая" if d < switch else "новая"
        print(f"  {d}  {money(byday[d]):>14}   {src}")
    print(f"  {'ИТОГО':10s}  {money(sum(byday.values())):>14}")
    print()

    # ---------------------------------------------------------- по группам
    print("ПЕРВАЯ ЧАСТЬ МЕСЯЦА · по направлениям (фактическая оплата)")
    bn = defaultdict(float)
    for r in old_in:
        bn[r["direction"] or "(без направления)"] += r["paid"]
    for k, v in sorted(bn.items(), key=lambda x: -x[1]):
        print(f"  {k[:32]:34s} {money(v):>14}")
    fnd = defaultdict(float)
    for r in old_in:
        fnd[group_of(r["doctor"])] += r["paid"]
    print("  из них учредители: " + ", ".join(f"{f} {money(fnd.get(f, 0))}" for f in FOUNDERS))
    print()

    if f_doc:
        emp, total = parse_doctors(f_doc)
        grp = defaultdict(float)
        for name, rec in emp.items():
            grp[group_of(name)] += rec["realization"]
        print("ВТОРАЯ ЧАСТЬ МЕСЯЦА · по врачам (реализация из «Взаиморасчётов»)")
        for k, v in sorted(grp.items(), key=lambda x: -x[1]):
            print(f"  {k[:32]:34s} {money(v):>14}")
        print(f"  {'ИТОГО':34s} {money(total['realization']):>14}")
        print()
        adv = sum(p["amount"] for p in pays) - sum(p["amount"] for p in pays if p["doctor"])
        print("НЕ СХОДИТСЯ")
        print(f"  отдельные платёжные документы:        {money(adv):>14}")
        print(f"  закрыли реализацию по врачам:         {money(total['paid_by_advance']):>14}")
        print(f"  разница (методология неизвестна):     {money(adv - total['paid_by_advance']):>14}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
