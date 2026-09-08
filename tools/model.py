# -*- coding: utf-8 -*-
"""Расчётная модель отчёта: из канонических данных -> все цифры отчёта + контрольные сверки.

Каждая цифра здесь имеет формулу и ссылку на источник — они же попадают
в лист «01 Цифры отчёта» файла-расшифровки.
"""
GROUPS = ["mal", "pog", "other", "misc", "cosm"]
GROUP_RU = {"mal": "Маланичев", "pog": "Погосян", "other": "Другие хирурги",
            "misc": "Прочие доходы", "cosm": "Косметология"}
FOUNDERS = ["mal", "pog"]
MONTHS_RU = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
             "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]


def _n(x):
    return 0.0 if x is None else float(x)


def _line(costs, key, field="total"):
    return _n((costs.get(key) or {}).get(field))


def compute(c):
    """c — канонический словарь. Возвращает словарь всех расчётных величин."""
    rev, costs, man = c["revenue"], c["costs"], c["manual"]
    exp = c["expenses"]
    out = {"meta": c["meta"]}

    # ---------------------------------------------------------------- выручка
    by_group = rev["by_group"]
    tot = rev["totals"]
    daily = rev.get("daily") or []
    out["revenue"] = {
        "daily": daily,
        "total": tot["total"],
        "cash": tot["cash"],
        "card": tot["card"],
        "bank_ind": tot["bank_ind"],
        "refund": tot["refund"],
        "by_group": {g: by_group[g]["total"] for g in GROUPS},
        "share": {g: (by_group[g]["total"] / tot["total"] * 100 if tot["total"] else 0) for g in GROUPS},
        "acquiring_fee": rev.get("acquiring_fee_total"),
        "acquiring_fee_pct": (rev["acquiring_fee_total"] / tot["card"] * 100
                              if rev.get("acquiring_fee_total") and tot["card"] else None),
        "month_cash_basis": rev.get("month_cash_basis"),
        "kkt_diff_total": rev.get("kkt_diff_total"),
    }
    if daily:
        best = max(daily, key=lambda d: d["total"])
        worst = min(daily, key=lambda d: d["total"])
        out["revenue"].update({
            "days": len(daily),
            "avg_day": sum(d["total"] for d in daily) / len(daily),
            "daily_sum": sum(d["total"] for d in daily),
            "best_day": best, "worst_day": worst, "daily": daily,
        })

    # ---------------------------------------------------------------- динамика
    y, m = c["meta"]["year"], c["meta"]["month"]
    hist = c.get("history", {})
    cur, prev = hist.get(str(y), {}), hist.get(str(y - 1), {})
    out["trend"] = {
        "cur": [cur.get(str(i)) for i in range(1, 13)],
        "prev": [prev.get(str(i)) for i in range(1, 13)],
        "months": MONTHS_RU,
    }
    same_month_prev = prev.get(str(m))
    this_month = cur.get(str(m))
    out["trend"]["yoy_month"] = ((this_month / same_month_prev - 1) * 100
                                 if same_month_prev and this_month else None)
    ytd_cur = [cur.get(str(i)) for i in range(1, m + 1)]
    ytd_prev = [prev.get(str(i)) for i in range(1, m + 1)]
    out["trend"]["ytd_cur"] = sum(v for v in ytd_cur if v) if all(ytd_cur) else None
    out["trend"]["ytd_prev"] = sum(v for v in ytd_prev if v) if all(ytd_prev) else None
    out["trend"]["ytd_growth"] = ((out["trend"]["ytd_cur"] / out["trend"]["ytd_prev"] - 1) * 100
                                  if out["trend"]["ytd_cur"] and out["trend"]["ytd_prev"] else None)
    out["trend"]["missing_prev"] = [MONTHS_RU[i - 1] for i in range(1, m + 1) if not prev.get(str(i))]
    out["trend"]["missing_cur"] = [MONTHS_RU[i - 1] for i in range(1, m + 1) if not cur.get(str(i))]

    # ---------------------------------------------------------------- расходы
    med = {r["name"]: r["amount"] for r in exp["medical"]}
    mgmt = {r["name"]: r["amount"] for r in exp["management"]}
    med_total = _n(exp["totals"].get("medical"))
    mgmt_total = _n(exp["totals"].get("management"))
    ip_surgeons = _line(costs, "ip_surgeons")
    repair = next((v for k, v in mgmt.items() if "ремонт" in k.lower()), 0.0)
    out["expenses"] = {
        "acc60_total": _n(exp["totals"].get("grand")),
        "medical_total": med_total, "management_total": mgmt_total,
        "medical": exp["medical"], "management": exp["management"],
        "vendors": exp["vendors"],
        "ip_surgeons": ip_surgeons,
        "medical_wo_ip": med_total - ip_surgeons,
        "repair": repair,
        "management_wo_repair": mgmt_total - repair,
        "share_of_revenue": (_n(exp["totals"].get("grand")) / tot["total"] * 100) if tot["total"] else None,
    }

    # ---------------------------------------------------------------- ФОТ и банк
    payroll = _line(costs, "salary") + _line(costs, "payroll_taxes") + _line(costs, "alimony")
    bank_fees = (_line(costs, "bank_common") + _line(costs, "bank_acquiring")
                 + _line(costs, "credit_interest"))
    # «Халва» в исходном файле не имеет итога в колонке «всего» — собираем из долей
    halva = _line(costs, "halva") or (_line(costs, "halva", "mal") + _line(costs, "halva", "pog"))
    other_small = _line(costs, "cash_expenses") + halva
    out["payroll"] = {
        "salary": _line(costs, "salary"),
        "taxes": _line(costs, "payroll_taxes"),
        "alimony": _line(costs, "alimony"),
        "total": payroll,
    }
    out["bank_fees"] = {"common": _line(costs, "bank_common"),
                        "acquiring": _line(costs, "bank_acquiring"),
                        "credit_interest": _line(costs, "credit_interest"), "total": bank_fees}

    # ---------------------------------------------------------------- водопад
    steps = [
        ("Выручка", "Выручка", tot["total"], "pos"),
        ("Гонорары ИП хирургов", "Гонорары хирургов", -ip_surgeons, "exp"),
        ("Медицинские расходы", "Медицина", -(med_total - ip_surgeons), "exp"),
        ("ФОТ, налоги, алименты", "ФОТ и налоги", -payroll, "exp"),
        ("Управленческие расходы", "Управление", -(mgmt_total - repair), "exp"),
        ("Ремонт нового корпуса", "Ремонт", -repair, "inv"),
        ("Услуги банка", "Банк", -bank_fees, "exp"),
        ("Прочее (касса, Халва)", "Прочее", -other_small, "exp"),
        ("Проценты по депозитам", "Депозиты",
         _n((man.get("deposit_share") or {}).get("mal")) + _n((man.get("deposit_share") or {}).get("pog")),
         "pos"),
    ]
    steps = [s for s in steps if s[2] or s[0] == "Выручка"]
    running, wf = 0.0, []
    for i, (name, short, val, cls) in enumerate(steps):
        if i == 0:
            wf.append({"name": name, "short": short, "value": val, "kind": cls, "from": 0.0, "to": val})
            running = val
        else:
            wf.append({"name": name, "short": short, "value": val, "kind": cls,
                       "from": running + val, "to": running})
            running += val
    wf.append({"name": "Доход месяца", "short": "Доход", "value": running, "kind": "pos",
               "from": 0.0, "to": running})
    out["waterfall"] = wf
    out["waterfall_result"] = running

    # ---------------------------------------------------------------- учредители
    # «% по кредиту» с июля выделены отдельной строкой; в июне сидели в услугах банка
    cost_keys = ["payroll_taxes", "salary", "cash_expenses", "bank_common",
                 "bank_acquiring", "alimony", "credit_interest", "adjust", "suppliers_bank"]
    founders = {}
    for f in FOUNDERS:
        own = _n(by_group[f]["total"])
        share_costs = sum(_line(costs, k, f) for k in cost_keys)
        half_other = _n(by_group["other"]["total"]) / 2
        half_cosm = _n(by_group["cosm"]["total"]) / 2
        half_misc = _n(by_group["misc"]["total"]) / 2
        # проценты по депозитам делятся пополам и входят в доход месяца
        deposit_share = _n((man.get("deposit_share") or {}).get(f))
        income = own - share_costs + half_other + half_cosm + half_misc + deposit_share
        prev_bal = _n((c["founders_reference"].get("prev_balance") or {}).get(f))
        reserve = _n((man.get("reserve_month") or {}).get(f))
        reserve_return = _n((man.get("reserve_return") or {}).get(f))
        dividends = _n((man.get("dividends") or {}).get(f))
        founders[f] = {
            "name": GROUP_RU[f],
            "own_revenue": own,
            "half_other": half_other,
            "half_cosm": half_cosm,
            "half_misc": half_misc,
            "costs": share_costs,
            "costs_detail": {k: _line(costs, k, f) for k in cost_keys},
            "income": income,
            "deposit_share": deposit_share,
            "prev_balance": prev_bal,
            "reserve": reserve,
            "reserve_return": reserve_return,
            "dividends": dividends,
            "payout": income + prev_bal - reserve + reserve_return - dividends,
        }
    out["founders"] = founders
    out["founders_total_income"] = sum(founders[f]["income"] for f in FOUNDERS)

    # ---------------------------------------------------------------- ДДС
    cf = c["cashflow"]
    out["cashflow"] = {
        "opening": _n(cf.get("opening")), "closing": _n(cf.get("closing")),
        "in_total": _n(cf.get("turnover_debit")), "out_total": _n(cf.get("turnover_credit")),
        "delta": _n(cf.get("closing")) - _n(cf.get("opening")),
        "flows": cf.get("flows", []),
    }
    out["manual"] = man

    out["checks"] = _checks(c, out)
    return out


def _checks(c, out):
    """Контрольные равенства. Каждое — из брифа «точность и проверка цифр»."""
    tot = c["revenue"]["totals"]
    by_group = c["revenue"]["by_group"]
    eps = 1.0
    ch = []

    def add(name, left, right, note=""):
        left, right = _n(left), _n(right)
        ch.append({"name": name, "left": left, "right": right,
                   "delta": left - right, "ok": abs(left - right) <= eps, "note": note})

    add("Сумма выручки по группам = итог выручки",
        sum(by_group[g]["total"] for g in GROUPS), tot["total"])
    note = ("часть месяца без разбивки по каналам" if tot.get("undivided_first_decade") else "")
    add("Касса + эквайринг + р/с физлиц − возвраты = итог выручки",
        tot["cash"] + tot["card"] + tot["bank_ind"] - tot["refund"]
        + _n(tot.get("undivided_first_decade")), tot["total"], note)
    add("Сумма шагов водопада = доход месяца",
        out["waterfall_result"], out["founders_total_income"],
        "водопад строится от выручки, доход — по методике учредителей")
    add("Доход Маланичева + Погосяна = доход компании",
        out["founders"]["mal"]["income"] + out["founders"]["pog"]["income"],
        out["founders_total_income"])
    ref = c.get("founders_reference", {})
    if ref.get("income"):
        add("Доход Маланичева = расчёт бухгалтерии",
            out["founders"]["mal"]["income"], ref["income"].get("mal"), "сверка с эталоном")
        add("Доход Погосяна = расчёт бухгалтерии",
            out["founders"]["pog"]["income"], ref["income"].get("pog"), "сверка с эталоном")
    if ref.get("payout"):
        add("К выплате Маланичеву = расчёт бухгалтерии",
            out["founders"]["mal"]["payout"], ref["payout"].get("mal"), "сверка с эталоном")
        add("К выплате Погосяну = расчёт бухгалтерии",
            out["founders"]["pog"]["payout"], ref["payout"].get("pog"), "сверка с эталоном")
    add("Расходы 60: медицинские + управленческие = итог по счёту",
        _n(c["expenses"]["totals"].get("medical")) + _n(c["expenses"]["totals"].get("management")),
        _n(c["expenses"]["totals"].get("grand")))
    add("Расходы 60: сумма по контрагентам = итог по счёту",
        sum(v["amount"] for v in c["expenses"]["vendors"]),
        _n(c["expenses"]["totals"].get("grand")) + _n(c["manual"].get("credit_funded")),
        "из итога исключено оплаченное за счёт кредитной линии"
        if c["manual"].get("credit_funded") else "")
    add("ДДС: сальдо нач. + приход − расход = сальдо кон.",
        _n(c["cashflow"]["opening"]) + _n(c["cashflow"]["turnover_debit"])
        - _n(c["cashflow"]["turnover_credit"]), _n(c["cashflow"]["closing"]))
    if c["revenue"].get("month_cash_basis"):
        add("Выручка (отчёт) vs выручка кассовым методом (график динамики)",
            tot["total"], c["revenue"]["month_cash_basis"],
            "две разные базы — расхождение ожидаемо, см. README")
    return ch
