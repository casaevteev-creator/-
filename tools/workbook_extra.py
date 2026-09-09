# -*- coding: utf-8 -*-
"""Дополнительные листы расшифровки за август: проверка классификации и сравнение с июлем."""
import json
import sys
from collections import defaultdict
from pathlib import Path

import openpyxl
from openpyxl.styles import Font

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_august import norm, MED_WORDS  # noqa: E402
from june_source import parse_expenses  # noqa: E402
from new_program import parse_finresult  # noqa: E402
from old_program import parse_doctor_report  # noqa: E402
from vendors import load_notes, read_turnover  # noqa: E402
from workbook import BOX, MONEY, SUB_FILL, TOT_FILL, _header, _row, _title  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "data" / "source" / "2026-08"

JULY = {"Выручка": 108126415.72, "Налоги с ФОТ": 6370709.49, "Зарплата": 12148685.72,
        "Расходы по кассе": 1170.0, "Услуги банка (общие)": 399308.17,
        "Эквайринг": 41563.80, "Алименты": 23201.61, "% по кредиту": 151764.47,
        "Поставщики (счёт 60)": 53416727.72}


def _n(v):
    return v if isinstance(v, (int, float)) else 0.0


def add_sheets(path, canon, model_out):
    wb = openpyxl.load_workbook(path)
    for name in list(wb.sheetnames):          # повторный запуск не должен плодить копии
        if name[:2] in ("11", "12", "13"):
            del wb[name]

    # ------------------------------------------------ 11 Классификация расходов
    turn = read_turnover(SRC / "1С-Обороты-счета-60-август-2026.xlsx")
    for name, delta in (canon["manual"].get("vendor_adjustments") or {}).items():
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

    ws = wb.create_sheet("11 Классификация расходов")
    top = _title(ws, "Счёт 60: как классифицирован каждый контрагент",
                 "Проверьте столбец «основание» — по нему видно, откуда взялась категория")
    _header(ws, ["Контрагент", "Оплачено с р/с", "Блок", "Статья", "Основание"],
            [44, 18, 20, 34, 46], row=top)
    r = top + 1
    for name, (pay, corr) in sorted(turn.items(), key=lambda x: -x[1][0]):
        low = norm(name)
        note = notes[low]["article"] if low in notes else prev.get(low, "")
        words = set(low.replace(".", " ").replace(",", " ").split())
        n = norm(note)
        if (words & last) and not n and "самозанят" not in low:
            block, art, why = "Медицинские", "Гонорары ИП хирургов", "фамилия совпала с врачом из управленки"
        elif low in notes and notes[low].get("bucket") == "медицинские":
            block, art, why = "Медицинские", note, "ручная пометка"
        elif any(k in n for k in MED_WORDS):
            block, art, why = "Медицинские", note, "назначение платежа за июнь"
        elif "ремонт" in n and ("стр" in n or "самокатная" in n):
            block, art, why = "Управленческие", "Ремонт нового корпуса", "назначение платежа"
        elif n:
            block, art, why = "Управленческие", note, ("ручная пометка" if low in notes
                                                       else "назначение платежа за июнь")
        elif corr["10"] > 0:
            block, art, why = "Медицинские", "Медицинские материалы", "кор. счёт 10 — материалы"
        elif corr["41"] > 0:
            block, art, why = "Управленческие", "Товары", "кор. счёт 41 — товары"
        elif corr["26"] > 0:
            block, art, why = "Управленческие", "Общехозяйственные", "кор. счёт 26"
        elif corr["08"] > 0 or corr["20"] > 0:
            block, art, why = "Управленческие", "Капвложения и работы", "кор. счета 08, 20"
        else:
            block, art, why = "Управленческие", "Прочие управленческие", "ТРЕБУЕТ УТОЧНЕНИЯ"
        r = _row(ws, r, [name, pay, block, art, why], money_cols=(2,))
    _row(ws, r, ["ИТОГО", sum(p for p, _ in turn.values()), "", "", ""],
         money_cols=(2,), bold=True, fill=TOT_FILL)

    # ------------------------------------------------ 12 Июль → август
    ws = wb.create_sheet("12 Июль-август")
    top = _title(ws, "Что изменилось против июля",
                 "Из расходов исключено оплаченное за счёт кредитной линии: 6 545 366,22 в июле и "
                 "11 000 000 в августе")
    _header(ws, ["Показатель", "Июль", "Август", "Разница"], [40, 20, 20, 20], row=top)
    r = top + 1
    a = canon["manual"]
    aug = {"Выручка": model_out["revenue"]["total"],
           "Налоги с ФОТ": canon["costs"]["payroll_taxes"]["total"],
           "Зарплата": canon["costs"]["salary"]["total"],
           "Расходы по кассе": canon["costs"]["cash_expenses"]["total"],
           "Услуги банка (общие)": canon["costs"]["bank_common"]["total"],
           "Эквайринг": canon["costs"]["bank_acquiring"]["total"],
           "Алименты": canon["costs"]["alimony"]["total"],
           "% по кредиту": canon["costs"]["credit_interest"]["total"],
           "Поставщики (счёт 60)": canon["costs"]["suppliers_bank"]["total"]}
    for k in JULY:
        r = _row(ws, r, [k, JULY[k], aug[k], aug[k] - JULY[k]], money_cols=(2, 3, 4))
    je = sum(v for k, v in JULY.items() if k != "Выручка")
    ae = sum(v for k, v in aug.items() if k != "Выручка")
    r = _row(ws, r, ["ИТОГО РАСХОДЫ", je, ae, ae - je], money_cols=(2, 3, 4), bold=True, fill=SUB_FILL)
    dep_j, dep_a = 442883.07, _n(canon["manual"].get("deposit_interest"))
    r = _row(ws, r, ["+ Проценты по депозитам", dep_j, dep_a, dep_a - dep_j], money_cols=(2, 3, 4))
    inc_j, inc_a = JULY["Выручка"] - je + dep_j, aug["Выручка"] - ae + dep_a
    r = _row(ws, r, ["ДОХОД УЧРЕДИТЕЛЕЙ", inc_j, inc_a, inc_a - inc_j],
             money_cols=(2, 3, 4), bold=True, fill=TOT_FILL)
    r += 1
    ws.cell(row=r, column=1, value="Внутри поставщиков").font = Font(bold=True)
    r += 1
    _header(ws, ["Статья", "Июль", "Август", "Разница"], [40, 20, 20, 20], row=r)
    r += 1
    sup_j = {"Гонорары ИП хирургов": 25223566.0, "Аренда Самокатная 1 стр.1": 4462500.0,
             "Коммунальные услуги": 224103.63, "Досудебный спор": 0.0,
             "Прочие поставщики": 23506558.09}
    med = {i["name"]: i["amount"] for i in model_out["expenses"]["medical"]}
    mgmt = {i["name"]: i["amount"] for i in model_out["expenses"]["management"]}
    sup_a = {"Гонорары ИП хирургов": med.get("Гонорары ИП хирургов", 0.0),
             "Аренда Самокатная 1 стр.1": mgmt.get("Аренда Самокатная 1 стр1", 0.0),
             "Коммунальные услуги": mgmt.get("Коммуналка", 0.0),   # в августе 0: см. примечание ниже
             "Досудебный спор": mgmt.get("Урегулирование досудебного спора", 0.0)}
    sup_a["Прочие поставщики"] = canon["costs"]["suppliers_bank"]["total"] - sum(sup_a.values())
    for k in sup_j:
        r = _row(ws, r, [k, sup_j[k], sup_a[k], sup_a[k] - sup_j[k]], money_cols=(2, 3, 4))
    _row(ws, r, ["ИТОГО", sum(sup_j.values()), sum(sup_a.values()),
                 sum(sup_a.values()) - sum(sup_j.values())],
         money_cols=(2, 3, 4), bold=True, fill=TOT_FILL)

    # ------------------------------------------------ 13 Разнесение выручки
    ws = wb.create_sheet("13 Разнесение выручки")
    top = _title(ws, "Как выручка разложена по группам",
                 "01–10.08 — старая программа по фактической оплате; 11–31.08 — новая, "
                 "платежи с врачом напрямую, авансы — через пациента")
    _header(ws, ["Источник", "Сумма", "Комментарий"], [42, 20, 60], row=top)
    r = top + 1
    al = canon["revenue"]["totals"]["allocation"]
    rf = canon["revenue"]["totals"]["refund"]
    for label, value, note in [
        ("01–10.08, старая управленка", al["first"],
         "реестр «Оплаты от пациента» без двух документов, помеченных на удаление; "
         "сходится с «Ежедневным отчётом по кассам»"),
        ("11–31.08, новая управленка", al["second"],
         "отчёт «Оплаты», виды оплаты «Наличными» и «Безналичными», уже за вычетом возвратов"),
        ("Возвраты пациентам, 01–10.08", -rf,
         "колонка «Возврат денег» «Ежедневного отчёта по кассам»"),
        ("ИТОГО ВЫРУЧКА", al["first"] + al["second"] - rf, "деньги, поступившие за месяц"),
        ("", None, ""),
        ("01–10.08: платежи с указанным специалистом", al["first"] - al["first_unresolved"],
         "колонка «Специалист» реестра платежей"),
        ("01–10.08: без специалиста", al["first_unresolved"],
         "разнесены по врачам того же пациента, остаток — пропорционально"),
        ("11–31.08: привязано к врачу напрямую", al["second_direct"],
         "документы «Оказание услуг» — врач указан в документе"),
        ("11–31.08: авансы, разнесённые через пациента", al["second_by_patient"],
         "оплата картой и кассовые ордера"),
        ("11–31.08: авансы без совпадения по пациенту", al["second_proportional"],
         "разнесены пропорционально; пациентов нет в реестре услуг августа"),
    ]:
        r = _row(ws, r, [label, value, note], money_cols=(2,),
                 bold=label.startswith("ИТОГО"), fill=TOT_FILL if label.startswith("ИТОГО") else None)
    r += 1
    _header(ws, ["Группа", "01–10.08", "11–31.08", "Август", "Доля"], [26, 18, 18, 18, 10], row=r)
    r += 1
    from model import GROUP_RU, GROUPS
    total = model_out["revenue"]["total"]
    split = canon["revenue"]["totals"]["by_group_split"]
    for g in GROUPS:
        v = model_out["revenue"]["by_group"][g]
        s1 = split[g]["first"] - split[g]["refund"]   # возвраты — первой декады
        s2 = split[g]["second"]
        r = _row(ws, r, [GROUP_RU[g], s1, s2, v, v / total], money_cols=(2, 3, 4))
        ws.cell(row=r - 1, column=5).number_format = "0.0%"
    _row(ws, r, ["ИТОГО", sum(split[g]["first"] - split[g]["refund"] for g in GROUPS),
                 sum(split[g]["second"] for g in GROUPS), total, 1.0],
         money_cols=(2, 3, 4), bold=True, fill=TOT_FILL)
    ws.cell(row=r, column=5).number_format = "0.0%"

    wb.save(path)
    return path


if __name__ == "__main__":
    import model
    canon = json.loads((ROOT / "data/canonical/2026-08.json").read_text(encoding="utf-8"))
    out = model.compute(canon)
    p = ROOT / "out" / "Расшифровка-отчёта-Август-2026.xlsx"
    print("готово:", add_sheets(p, canon, out))
