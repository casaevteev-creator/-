# -*- coding: utf-8 -*-
"""Сборка канонических данных за август 2026 из выгрузок двух управленок и 1С."""
import datetime
import json
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from june_source import parse_expenses  # noqa: E402
from new_program import (allocate_advances, enrich, open_fixed, parse_finresult,  # noqa: E402
                         parse_payments, parse_services, report_group, _num)
from old_program import parse_doctor_report, parse_patient_payments  # noqa: E402
from vendors import load_notes, read_turnover  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "data" / "source" / "2026-08"
SWITCH = datetime.date(2026, 8, 11)
GRP = {"Маланичев": "mal", "Погосян": "pog", "Другие хирурги": "other",
       "Косметология": "cosm", "Прочие доходы": "misc"}
GROUPS = ["mal", "pog", "other", "misc", "cosm"]
MED_WORDS = ["расходк", "шовн", "имплант", "лекарств", "белье", "наркот", "кров", "анализ",
             "лаборат", "медоборуд", "мед.", "мед ", "инструмент", "перевяз", "стерил",
             "утилизац", "отход", "inmode", "склиф", "центр крови", "питание"]


def norm(s):
    return " ".join(str(s or "").split()).lower().replace("ё", "е")


# --------------------------------------------------------------------------- выручка
def revenue(first_decade_card=None):
    old = [r for r in parse_doctor_report(SRC / "Старая-Отчет-по-врачам-УК-01.08-10.09.2026.xlsx")
           if r["date"] and r["date"].year == 2026 and r["date"].month == 8 and r["date"] < SWITCH]
    dec1 = parse_patient_payments(SRC / "Старая-Оплаты-от-пациента-01-10.08.2026.xlsx")
    cells = parse_finresult(SRC / "Новая-Финрезультат-10-31.08.2026.xlsx")
    dom = defaultdict(lambda: defaultdict(float))
    for c in cells:
        dom[c["doctor"]][c["spec"]] += c["sale"]
    cosmetologists = {d.split()[0] for d, v in dom.items()
                      if sum(v.values()) > 0
                      and max(v.items(), key=lambda x: x[1])[0] == "Косметология" and d.split()}
    main_spec = {d: (max(v.items(), key=lambda x: x[1])[0] if sum(v.values()) > 0 else "Хирургия")
                 for d, v in dom.items()}

    def group_old(r):
        doctor = str(r["doctor"] or "")
        direction = (r["direction"] or "").strip()
        if "Маланичев" in doctor:
            return "mal"
        if "Погосян" in doctor:
            return "pog"
        if direction == "Косметология" or (doctor.split() and doctor.split()[0] in cosmetologists):
            return "cosm"
        if direction in ("Товары", "КДЛ"):
            return "misc"
        return "other"

    # первая декада считается по реестру платежей: это деньги, а не оказанные услуги.
    # платежи без специалиста разносим по врачам того же пациента, остаток — пропорционально
    by_patient = defaultdict(lambda: defaultdict(float))
    for r in old:
        if r["doctor"] and r["revenue"] > 0:
            by_patient[norm(r["patient"])][str(r["doctor"]).strip()] += r["revenue"]

    def group_doctor(doctor):
        if "Маланичев" in doctor:
            return "mal"
        if "Погосян" in doctor:
            return "pog"
        return "cosm" if (doctor.split() and doctor.split()[0] in cosmetologists) else "other"

    first, unresolved = defaultdict(float), 0.0
    for p in dec1:
        if p["doctor"]:
            first[group_doctor(p["doctor"])] += p["total"]
            continue
        docs = by_patient.get(norm(p["patient"]))
        if not docs:
            unresolved += p["total"]
            continue
        tot = sum(docs.values())
        for d, v in docs.items():
            first[group_doctor(d)] += p["total"] * v / tot
    base = sum(first.values())
    for k in list(first):
        first[k] += unresolved * first[k] / base
    dec1_channels = {"total": round(sum(p["total"] for p in dec1), 2),
                     "card": round(sum(p["card"] for p in dec1), 2),
                     "bank": round(sum(p["bank"] for p in dec1), 2),
                     "cash": round(sum(p["cash"] for p in dec1), 2)}

    pays = [p for p in enrich(parse_payments(SRC / "Новая-Оплаты-с-врачом-01-31.08.2026.xlsx"),
                              parse_services(SRC / "Новая-Услуги-реестр-01-31.08.2026.xlsx"))
            if p["dt"].date() >= SWITCH]
    srv = [s for s in parse_services(SRC / "Новая-Услуги-реестр-01-31.08.2026.xlsx")
           if s["dt"] and s["dt"].year == 2026 and s["dt"].month == 8]
    allocated, unmatched = allocate_advances(pays, srv)

    second = defaultdict(float)
    for p in pays:
        if p["doctor"]:
            second[GRP[report_group(main_spec.get(p["doctor"], "Хирургия"), p["doctor"])]] += p["amount"]
    for d, v in allocated.items():
        second[GRP[report_group(main_spec.get(d, "Хирургия"), d)]] += v
    rest = sum(p["amount"] for p in unmatched)
    base = sum(second.values())
    for k in list(second):                      # авансы без совпадения по клиенту — пропорционально
        second[k] += rest * second[k] / base

    by_group = {g: {"cash": 0.0, "card": 0.0, "bank_ind": 0.0, "refund": 0.0,
                    "total": round(first.get(g, 0.0) + second.get(g, 0.0), 2)} for g in GROUPS}
    cash2 = round(sum(p["amount"] for p in pays if p["way"] == "Наличными"), 2)
    card2 = round(sum(p["amount"] for p in pays if p["way"] == "Безналичными"), 2)
    totals = {"refund": 0.0,
              "cash": round(dec1_channels["cash"] + cash2, 2),
              "card": round(dec1_channels["card"] + card2, 2),
              "bank_ind": dec1_channels["bank"],
              "total": round(sum(v["total"] for v in by_group.values()), 2),
              "split": {"first_cash": dec1_channels["cash"], "first_card": dec1_channels["card"],
                        "first_bank": dec1_channels["bank"],
                        "second_cash": cash2, "second_card": card2}}

    daily = defaultdict(float)
    for p in dec1:
        daily[p["dt"].date()] += p["total"]
    for p in pays:
        daily[p["dt"].date()] += p["amount"]
    days = [{"day": d.day, "date": d.isoformat(), "cash": {}, "card": {},
             "total": round(v, 2), "kkt": None, "kkt_diff": None}
            for d, v in sorted(daily.items())]
    return by_group, totals, days


# --------------------------------------------------------------------------- расходы
def expenses(adjustments=None):
    turn = read_turnover(SRC / "1С-Обороты-счета-60-август-2026.xlsx")
    for name, delta in (adjustments or {}).items():      # платежи, не попавшие в выгрузку
        pay, corr = turn.get(name, (0.0, {k: 0.0 for k in ("08", "10", "19", "20", "26", "41", "76")}))
        turn[name] = (pay + delta, corr)
    prev = {norm(v["name"]): (v.get("note") or "")
            for v in parse_expenses(ROOT / "data/source/Расходы-июнь-2026.xls")["vendors"]}
    notes = {norm(k): v for k, v in load_notes().items()}
    doctors = {c["doctor"] for c in parse_finresult(SRC / "Новая-Финрезультат-10-31.08.2026.xlsx")}
    doctors |= {str(r["doctor"]) for r in
                parse_doctor_report(SRC / "Старая-Отчет-по-врачам-УК-01.08-10.09.2026.xlsx")
                if r["doctor"]}
    last = {norm(d).split()[0] for d in doctors if norm(d).split()}
    last.discard("прочее")

    med, mgmt, vendors, unclear = defaultdict(float), defaultdict(float), [], {}
    for name, (pay, corr) in turn.items():
        low = norm(name)
        note = notes[low]["article"] if low in notes else prev.get(low, "")
        words = set(low.replace(".", " ").replace(",", " ").split())
        n = norm(note)
        # «Самозанятый Маланичев Ю.В.» — инженер, а не хирург: однофамильцев отсекаем
        # по тому, что у них есть назначение платежа из классификатора
        if (words & last) and not n and "самозанят" not in low:
            bucket, article = "med", "Гонорары ИП хирургов"
        elif low in notes and notes[low].get("bucket") == "медицинские":
            bucket, article = "med", note
        elif any(k in n for k in MED_WORDS):
            bucket, article = "med", note
        elif "ремонт" in n and ("стр" in n or "самокатная" in n):
            bucket, article = "mgmt", "Ремонт нового корпуса"
        elif n:
            bucket, article = "mgmt", note
        elif corr["10"] > 0:
            bucket, article = "med", "Медицинские материалы"
        elif corr["41"] > 0:
            bucket, article = "mgmt", "Товары"
        elif corr["26"] > 0:
            bucket, article = "mgmt", "Общехозяйственные"
        elif corr["08"] > 0 or corr["20"] > 0:
            bucket, article = "mgmt", "Капвложения и работы"
        else:
            bucket, article = "mgmt", "Прочие управленческие"
            unclear[name] = pay
        (med if bucket == "med" else mgmt)[article] += pay
        vendors.append({"name": name, "amount": round(pay, 2), "note": note or article})

    to_list = lambda d: [{"name": k, "amount": round(v, 2)}
                         for k, v in sorted(d.items(), key=lambda x: -x[1])]
    return {
        "vendors": vendors,
        "medical": to_list(med), "management": to_list(mgmt),
        "totals": {"medical": round(sum(med.values()), 2),
                   "management": round(sum(mgmt.values()), 2),
                   "grand": round(sum(med.values()) + sum(mgmt.values()), 2)},
        "unclear": {k: round(v, 2) for k, v in unclear.items()},
    }


# --------------------------------------------------------------------------- ДДС
def cashflow():
    ws = open_fixed(SRC / "1С-Анализ-счета-51-август-2026.xlsx").worksheets[0]
    hdr = [str(ws.cell(row=6, column=c).value or "") for c in range(1, ws.max_column + 1)]
    row = next(r for r in range(7, ws.max_row + 1)
               if str(ws.cell(row=r, column=1).value or "").strip() == "51")
    val = {}
    for i, h in enumerate(hdr, start=1):
        v = _num(ws.cell(row=row, column=i).value)
        if v:
            val.setdefault(h, []).append(v)
    dt_names = ["55", "57", "60", "67", "76", "91"]
    kt_names = ["55", "57", "60", "62", "68", "69", "70", "75", "76", "91"]
    labels = {
        "55": ("Возврат средств с депозитов", "Размещение на депозиты"),
        "57": ("Инкассация и эквайринг", "Переводы между счетами"),
        "60": ("Возвраты от поставщиков", "Оплаты поставщикам и подрядчикам"),
        "62": ("Оплаты пациентов на р/с", "Возвраты пациентам"),
        "67": ("Кредит банка", "Погашение кредита"),
        "68": ("", "Налоги"), "69": ("", "Страховые взносы"),
        "70": ("", "Зарплата"), "75": ("", "Выплата дивидендов"),
        "76": ("Прочее", "Алименты"), "91": ("Проценты по депозитам", "Услуги банка"),
    }
    flows = []
    idx_dt = hdr.index("Оборот Дт")
    for i, h in enumerate(hdr):
        v = _num(ws.cell(row=row, column=i + 1).value)
        if not v or h in ("Счет", "Оборот Дт", "Оборот Кт",
                          "Начальное сальдо Дт", "Начальное сальдо Кт",
                          "Конечное сальдо Дт", "Конечное сальдо Кт"):
            continue
        debit = i < hdr.index("Оборот Кт")
        name = labels.get(h, (h, h))[0 if debit else 1] or f"Счёт {h}"
        flows.append({"label": name, "debit": v if debit else None,
                      "credit": None if debit else v, "note": f"кор. счёт {h}"})
    return {
        "opening": _num(ws.cell(row=row, column=2).value),
        "closing": _num(ws.cell(row=row, column=hdr.index("Конечное сальдо Дт") + 1).value),
        "turnover_debit": _num(ws.cell(row=row, column=idx_dt + 1).value),
        "turnover_credit": _num(ws.cell(row=row, column=hdr.index("Оборот Кт") + 1).value),
        "flows": flows,
    }


def main():
    manual = json.loads((ROOT / "data" / "manual_2026-08.json").read_text(encoding="utf-8"))
    a = manual["august_2026"]
    by_group, totals, daily = revenue(a.get("first_decade_card"))
    exp = expenses(a.get("vendor_adjustments"))
    cf = cashflow()
    adj = sum((a.get("vendor_adjustments") or {}).values())
    if adj:                                   # тот же платёж в ДДС: выборка кредита -> поставщику
        for f in cf["flows"]:
            if f["label"] == "Оплаты поставщикам и подрядчикам":
                f["credit"] = round(f["credit"] + adj, 2)
            if f["label"] == "Кредит банка":
                f["debit"] = round(f["debit"] + adj, 2)
        cf["turnover_debit"] = round(cf["turnover_debit"] + adj, 2)
        cf["turnover_credit"] = round(cf["turnover_credit"] + adj, 2)

    half = lambda x: {"total": round(x, 2), "mal": round(x / 2, 2), "pog": round(x / 2, 2)}
    # то, что оплачено за счёт кредитной линии, в расходы учредителей не входит:
    # стройка финансируется заёмными, а не выручкой (та же логика в июне и июле)
    credit_funded = a.get("credit_funded", 0.0)
    suppliers = exp["totals"]["grand"] - credit_funded
    for item in exp["management"]:
        if item["name"] == "Ремонт нового корпуса":
            item["amount"] = round(item["amount"] - credit_funded, 2)
    exp["management"] = [i for i in exp["management"] if i["amount"]]
    exp["totals"]["management"] = round(exp["totals"]["management"] - credit_funded, 2)
    exp["totals"]["grand"] = round(exp["totals"]["grand"] - credit_funded, 2)
    exp["credit_funded"] = credit_funded
    costs = {
        "payroll_taxes": half(a["payroll_taxes"] + a["social_69"]),
        "salary": half(a["salary_total"]),
        "cash_expenses": half(0.0),
        "bank_common": half(a["bank_services"]),
        "bank_acquiring": half(a["acquiring_fee"]),
        "alimony": half(a["alimony"]),
        "credit_interest": half(a["credit_interest"]),
        "adjust": half(0.0),
        "suppliers_bank": half(suppliers),
        "ip_surgeons": half(next((i["amount"] for i in exp["medical"]
                                  if i["name"] == "Гонорары ИП хирургов"), 0.0)),
    }
    costs["total"] = {k: round(sum(costs[c][k] for c in
                                   ("payroll_taxes", "salary", "cash_expenses", "bank_common",
                                    "bank_acquiring", "alimony", "credit_interest", "adjust",
                                    "suppliers_bank")), 2) for k in ("total", "mal", "pog")}

    hist = json.loads((ROOT / "data" / "history.json").read_text(encoding="utf-8"))
    canon = {
        "meta": {"company": "ООО «ОМЕГА»", "brand": "Форма", "year": 2026, "month": 8,
                 "month_ru": "Август", "period": "Август 2026",
                 "source": "две управленки + 1С (счета 51, 60, 68, 91)"},
        "revenue": {"daily": daily, "by_group": by_group, "totals": totals,
                    "month_cash_basis": totals["total"], "acquiring_fee": {},
                    "acquiring_fee_total": a["acquiring_fee"], "kkt_diff_total": None},
        "history": hist,
        "expenses": exp,
        "cashflow": cf,
        "costs": costs,
        "founders_reference": {"prev_balance": {"row": None, **a["opening_balance"],
                                                "total": round(sum(a["opening_balance"].values()), 2)}},
        "manual": {
            "reserve_total": 0.0,
            "reserve_month": {"mal": 0.0, "pog": 0.0},
            "reserve_return": {"mal": 0.0, "pog": 0.0},
            "dividends": {"mal": a["dividends_gross"] / 2, "pog": a["dividends_gross"] / 2},
            "deposit_share": {"mal": a["deposit_interest"] / 2, "pog": a["deposit_interest"] / 2},
            "credit_debt": a["credit_line_debt"],
            "credit_funded": credit_funded,
            "vendor_adjustments": a.get("vendor_adjustments"),
            "credit_line_limit": a.get("credit_line_limit"),
            "credit_line_available": a.get("credit_line_available"),
            "cash_on_hand": a["cash_on_hand_31_08"],
            "deposit_interest": a["deposit_interest"],
        },
    }
    out = ROOT / "data" / "canonical" / "2026-08.json"
    out.write_text(json.dumps(canon, ensure_ascii=False, indent=1), encoding="utf-8")
    print("записано:", out)
    return canon


if __name__ == "__main__":
    main()
