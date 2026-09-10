# -*- coding: utf-8 -*-
"""Рабочий файл расчёта в логике бухгалтерии, но причёсанный.

Та же последовательность блоков, что в исходном файле БАНК (движение по счёту 51 →
выручка → расходы → доход учредителей → к выплате → справочно), только с живыми
формулами, нормальными заголовками и без объединённых ячеек.

    python3 tools/workbook_accountant.py data/canonical/2026-08.json
"""
import json
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

sys.path.insert(0, str(Path(__file__).resolve().parent))
import model  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
INK, GREEN, BRONZE, TINT, LINE, RED = "1C1B18", "2E5E4E", "8A7350", "F1EFEA", "D9D4C9", "A3423B"
MONEY = "#,##0.00;-#,##0.00;—"
THIN = Side(style="thin", color=LINE)
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
COLS = {"label": 2, "total": 3, "mal": 4, "pog": 5, "other": 6, "misc": 7, "cosm": 8, "note": 9}


class Sheet:
    """Тонкая обёртка: пишет строку и помнит, в какой строке что лежит."""

    def __init__(self, ws):
        self.ws, self.row, self.at = ws, 1, {}

    def _cell(self, col, value, **st):
        c = self.ws.cell(row=self.row, column=col, value=value)
        if st.get("money"):
            c.number_format = MONEY
            c.alignment = Alignment(horizontal="right")
        if st.get("bold"):
            c.font = Font(bold=True, size=st.get("size", 11), color=st.get("color", "000000"))
        elif st.get("color") or st.get("size"):
            c.font = Font(size=st.get("size", 11), color=st.get("color", "000000"))
        if st.get("fill"):
            c.fill = PatternFill("solid", fgColor=st["fill"])
        if st.get("border", True):
            c.border = BOX
        if st.get("wrap"):
            c.alignment = Alignment(wrap_text=True, vertical="top")
        return c

    def title(self, text, sub=""):
        c = self.ws.cell(row=self.row, column=2, value=text)
        c.font = Font(bold=True, size=16)
        self.row += 1
        if sub:
            c = self.ws.cell(row=self.row, column=2, value=sub)
            c.font = Font(italic=True, size=10, color="807A70")
            self.row += 1
        self.row += 1

    def section(self, text, note=""):
        self.row += 1
        for col in range(2, 10):
            self._cell(col, text if col == 2 else (note if col == 9 else None),
                       bold=(col == 2), color="FFFFFF", fill=INK, size=11)
        self.ws.row_dimensions[self.row].height = 22
        self.row += 1

    def head(self, labels):
        for key, text in labels.items():
            self._cell(COLS[key], text, bold=True, fill=TINT, size=10)
        self.ws.row_dimensions[self.row].height = 26
        self.row += 1

    def line(self, label, values=None, note="", key=None, **st):
        self._cell(COLS["label"], label, bold=st.get("bold"),
                   color=st.get("color", "000000"))
        for col, v in (values or {}).items():
            self._cell(COLS[col], v, money=True, bold=st.get("bold"),
                       fill=st.get("fill"), color=st.get("color", "000000"))
        if note:
            self._cell(COLS["note"], note, size=9, color="807A70", wrap=True)
        for col in range(2, 10):
            if st.get("fill"):
                self.ws.cell(row=self.row, column=col).fill = PatternFill("solid", fgColor=st["fill"])
            self.ws.cell(row=self.row, column=col).border = BOX
        if key:
            self.at[key] = self.row
        self.row += 1
        return self.row - 1

    def total(self, label, values, note="", key=None):
        return self.line(label, values, note, key=key, bold=True, fill=TINT)

    def result(self, label, values, note="", key=None):
        return self.line(label, values, note, bold=True, fill="EDF1EE", color=GREEN, key=key)

    def blank(self):
        self.row += 1

    def ref(self, key, col):
        return f"{L(COLS[col])}{self.at[key]}"


def build(canon, out_path):
    r = model.compute(canon)
    meta, man = canon["meta"], canon["manual"]
    cf, costs = canon["cashflow"], canon["costs"]
    f, rev = r["founders"], r["revenue"]

    wb = Workbook()
    ws = wb.active
    ws.title = "Расчёт"
    s = Sheet(ws)
    for col, w in ((1, 3), (2, 46), (3, 18), (4, 17), (5, 17), (6, 17), (7, 15), (8, 15), (9, 44)):
        ws.column_dimensions[L(col)].width = w

    s.title(f"{meta['company']} · расчёт за {meta['period'].lower()}",
            "Выручка — из управленки, расходы и движение денег — из 1С. "
            "Итоговые строки посчитаны формулами, их видно в ячейках.")

    # ---------------------------------------------------------------- 1. счёт 51
    s.section("1. ДВИЖЕНИЕ ПО РАСЧЁТНОМУ СЧЁТУ (счёт 51)")
    s.head({"label": "Статья", "total": "Сумма", "note": "Кор. счёт"})
    s.line("Остаток на начало месяца", {"total": cf["opening"]}, key="open")
    s.blank()
    s.line("ПОСТУПЛЕНИЯ", bold=True)
    first_in = s.row
    for fl in cf["flows"]:
        if fl.get("debit"):
            s.line("    " + fl["label"], {"total": fl["debit"]}, fl.get("note", ""))
    last_in = s.row - 1
    s.total("Итого поступления", {"total": f"=SUM(C{first_in}:C{last_in})"}, key="in")
    s.blank()
    s.line("СПИСАНИЯ", bold=True)
    first_out = s.row
    for fl in cf["flows"]:
        if fl.get("credit"):
            s.line("    " + fl["label"], {"total": fl["credit"]}, fl.get("note", ""))
    last_out = s.row - 1
    s.total("Итого списания", {"total": f"=SUM(C{first_out}:C{last_out})"}, key="out")
    s.result("Остаток на конец месяца",
             {"total": f"={s.ref('open','total')}+C{s.at['in']}-C{s.at['out']}"},
             "начальный остаток + поступления − списания")

    # ---------------------------------------------------------------- 2. выручка
    s.section("2. ВЫРУЧКА", "деньги, поступившие от пациентов за месяц")
    s.head({"label": "Показатель", "total": "Всего", "mal": "Маланичев", "pog": "Погосян",
            "other": "Другие хирурги", "misc": "Прочие доходы", "cosm": "Косметология",
            "note": "Комментарий"})
    g = rev["by_group"]
    s.result("ИТОГО ВЫРУЧКА ЗА МЕСЯЦ",
             {"total": "=SUM(D{0}:H{0})".format(s.row), "mal": g["mal"], "pog": g["pog"],
              "other": g["other"], "misc": g["misc"], "cosm": g["cosm"]},
             "по группам врачей", key="rev")
    s.blank()
    s.line("В том числе по каналам поступления", bold=True)
    tot = canon["revenue"]["totals"]
    ch_first = s.row
    sp = tot.get("split")
    if tot.get("undivided_first_decade"):
        s.line("    Первая декада, старая управленка", {"total": tot["undivided_first_decade"]},
               "способа оплаты в выгрузке старой программы нет")
        s.line("    Наличными", {"total": tot["cash"]}, "новая управленка, 11–31.08")
        s.line("    Безналичными (эквайринг)", {"total": tot["card"]}, "новая управленка, 11–31.08")
    else:
        s.line("    Наличными", {"total": tot["cash"]},
               (f"01–10.08: {sp['first_cash']:,.0f}".replace(",", " ")
                + f"; 11–31.08: {sp['second_cash']:,.0f}".replace(",", " ")) if sp else "")
        num = lambda v: f"{v:,.0f}".replace(",", " ")
        online = (f", из них онлайн-оплаты через сайт {num(sp['first_online'])}"
                  if sp and sp.get("first_online") else "")
        back = (f"; за вычетом возвратов пациентам {num(tot['refund'])}"
                if tot.get("refund") else "")
        s.line("    Безналичными (карты и онлайн-оплаты)",
               {"total": tot["card"] - tot.get("refund", 0.0)},
               (f"01–10.08: {num(sp['first_card'])}" + online
                + f"; 11–31.08: {num(sp['second_card'])}" + back) if sp else "")
    if tot.get("bank_ind"):
        s.line("    Оплаты физлиц на расчётный счёт", {"total": tot["bank_ind"]})
    s.total("Итого по каналам", {"total": f"=SUM(C{ch_first}:C{s.row-1})"},
            "должно совпасть с итогом выручки выше")
    if tot.get("advances"):
        sv = tot.get("services_split") or {}
        s.blank()
        s.line("Справочно: оказано услуг за месяц", {"total": tot["services"]},
               (f"01–10.08: {sv.get('first',0):,.0f}".replace(",", " ")
                + f"; 11–31.08: {sv.get('second',0):,.0f}".replace(",", " ")) if sv else "",
               color=BRONZE)
        pct = f"{tot['advances'] / tot['total'] * 100:.1f}".replace(".", ",")
        s.line("Справочно: авансы под будущие операции",
               {"total": tot["advances"]},
               "выручка минус оказанные услуги: деньги получены, операция ещё впереди — "
               f"{pct}% выручки месяца", color=BRONZE)

    # ---------------------------------------------------------------- 3. расходы
    s.section("3. РАСХОДЫ (оплаченные)", "делятся между учредителями 50/50")
    s.head({"label": "Статья", "total": "Всего", "mal": "Маланичев", "pog": "Погосян",
            "note": "Комментарий"})
    exp_first = s.row
    rows = [
        ("Налоги с ФОТ (счета 68.90 и 69)", "payroll_taxes", ""),
        ("Зарплата по банку", "salary", ""),
        ("Расходы по кассе", "cash_expenses", ""),
        ("Услуги банка (общие)", "bank_common", ""),
        ("Услуги банка (эквайринг)", "bank_acquiring", ""),
        ("Алименты (удержание из зарплаты)", "alimony", ""),
        ("Проценты по кредиту", "credit_interest", ""),
        ("Дополнительные корректировки", "adjust", ""),
    ]
    for label, key, note in rows:
        v = costs.get(key, {})
        if not (v.get("total") or v.get("mal") or v.get("pog")):
            continue
        row = s.row
        s.line(label, {"total": v.get("total", 0.0),
                       "mal": f"=C{row}/2", "pog": f"=C{row}/2"}, note)
    fnd = r["founders"]
    items = sorted({k for x in ("mal", "pog") for k in fnd[x].get("personal_items", {})})
    sup_row = s.row
    s.line("Поставщики и подрядчики (счёт 60)", {"total": costs["suppliers_bank"]["total"]},
           "без оплаченного за счёт кредитной линии", bold=True)
    for block, blk in (("медицинские и гонорары", r["expenses"]["medical"]),
                       ("управленческие", r["expenses"]["management"])):
        if not blk:
            continue
        s.line(f"        в том числе {block}:", color="807A70")
        for item in blk:                          # все статьи, без «прочих»
            s.line("            " + item["name"], {"total": item["amount"]})
    if items:
        for who, key in (("Маланичева", "mal"), ("Погосяна", "pog")):
            s.line(f"        − личные расходы {who}, они внутри счёта 60",
                   {"total": -fnd[key]["personal_costs"]},
                   "заказано лично под врача — пополам не делится", color=RED)
        net_row = s.row
        s.line("Поставщики к делению пополам",
               {"total": f"=C{sup_row}+C{net_row - 2}+C{net_row - 1}",
                "mal": f"=C{net_row}/2", "pog": f"=C{net_row}/2"}, bold=True)
    else:
        net_row = sup_row
        ws.cell(row=sup_row, column=4, value=f"=C{sup_row}/2")
        ws.cell(row=sup_row, column=5, value=f"=C{sup_row}/2")
    s.at["exp_end"] = s.row - 1
    com_row = s.row
    s.result("ИТОГО РАСХОДЫ К ДЕЛЕНИЮ ПОПОЛАМ",
             {"total": f"=SUM(C{exp_first}:C{sup_row - 1})+C{net_row}",
              "mal": f"=C{com_row}/2", "pog": f"=C{com_row}/2"},
             "строки «в том числе» в сумму не входят")
    if items:
        s.line("+ Личные расходы, каждому свои",
               {"total": f"=D{s.row}+E{s.row}",
                "mal": fnd["mal"]["personal_costs"], "pog": fnd["pog"]["personal_costs"]},
               "те самые суммы, что вычтены из счёта 60 выше")
        per_row = s.row - 1
        for name in items:
            s.line(f"        {name}",
                   {"total": f"=D{s.row}+E{s.row}",
                    "mal": fnd["mal"]["personal_items"].get(name, 0.0),
                    "pog": fnd["pog"]["personal_items"].get(name, 0.0)}, color="807A70")
        s.result("ДОЛЯ РАСХОДОВ КАЖДОГО",
                 {"total": f"=D{s.row}+E{s.row}",
                  "mal": f"=D{com_row}+D{per_row}", "pog": f"=E{com_row}+E{per_row}"},
                 "к делению пополам плюс свои личные", key="exp")
    else:
        s.at["exp"] = com_row
    if man.get("credit_funded"):
        s.line("Справочно: ремонт за счёт кредитной линии", {"total": man["credit_funded"]},
               "в расходы учредителей не входит — стройка идёт на заёмные", color=BRONZE)

    # ---------------------------------------------------------------- 4. доход
    s.section("4. РАСЧЁТ ДОХОДА УЧРЕДИТЕЛЕЙ", "методика 50/50")
    s.head({"label": "Шаг", "total": "Всего", "mal": "Маланичев", "pog": "Погосян",
            "note": "Как считается"})
    rev_row = s.at["rev"]
    inc_first = s.row
    s.line("1. Личная выручка (свои операции)",
           {"total": f"=D{s.row}+E{s.row}", "mal": f"=D{rev_row}", "pog": f"=E{rev_row}"},
           "своя строка из блока «Выручка»")
    s.line("2. Половина выручки других хирургов",
           {"total": f"=F{rev_row}", "mal": f"=F{rev_row}/2", "pog": f"=F{rev_row}/2"})
    s.line("3. Половина выручки косметологии",
           {"total": f"=H{rev_row}", "mal": f"=H{rev_row}/2", "pog": f"=H{rev_row}/2"})
    s.line("4. Половина прочих доходов",
           {"total": f"=G{rev_row}", "mal": f"=G{rev_row}/2", "pog": f"=G{rev_row}/2"})
    dep = man.get("deposit_interest") or 0.0
    if dep:
        s.line("5. Половина процентов по депозитам",
               {"total": dep, "mal": f"=C{s.row}/2", "pog": f"=C{s.row}/2"},
               "начислены банком на остатки")
    s.line("6. Минус доля расходов",
           {"total": f"=-C{s.at['exp']}", "mal": f"=-D{s.at['exp']}", "pog": f"=-E{s.at['exp']}"},
           "строка «Доля расходов каждого» блока 3", color=RED)
    inc_last = s.row - 1
    s.result("ДОХОД ЗА МЕСЯЦ",
             {"total": f"=SUM(C{inc_first}:C{inc_last})",
              "mal": f"=SUM(D{inc_first}:D{inc_last})", "pog": f"=SUM(E{inc_first}:E{inc_last})"},
             key="income")

    # ---------------------------------------------------------------- 5. к выплате
    s.section("5. К ВЫПЛАТЕ УЧРЕДИТЕЛЯМ")
    s.head({"label": "Показатель", "total": "Всего", "mal": "Маланичев", "pog": "Погосян",
            "note": "Комментарий"})
    pay_first = s.row
    s.line("Доход за месяц",
           {"total": f"=C{s.at['income']}", "mal": f"=D{s.at['income']}", "pog": f"=E{s.at['income']}"})
    prev = canon["founders_reference"]["prev_balance"]
    s.line("Остаток с прошлого месяца",
           {"total": f"=D{s.row}+E{s.row}", "mal": prev["mal"], "pog": prev["pog"]},
           "ручной ввод бухгалтерии")
    for label, key, sign, note in (
            ("Выплачены дивиденды", "dividends", -1, "начислено с НДФЛ"),
            ("Удержано в резерв на аренду", "reserve_month", -1, "ручной ввод бухгалтерии"),
            ("Возврат остатка резерва", "reserve_return", 1, "ручной ввод бухгалтерии")):
        v = man.get(key) or {}
        if not (v.get("mal") or v.get("pog")):
            continue
        s.line(label, {"total": f"=D{s.row}+E{s.row}",
                       "mal": sign * v["mal"], "pog": sign * v["pog"]}, note,
               color=RED if sign < 0 else "000000")
    pay_last = s.row - 1
    import calendar
    last_day = calendar.monthrange(meta["year"], meta["month"])[1]
    s.result(f"ИТОГО К ВЫПЛАТЕ НА {last_day:02d}.{meta['month']:02d}.{meta['year']}",
             {"total": f"=SUM(C{pay_first}:C{pay_last})",
              "mal": f"=SUM(D{pay_first}:D{pay_last})", "pog": f"=SUM(E{pay_first}:E{pay_last})"})

    # ---------------------------------------------------------------- 6. справочно
    s.section("6. СПРАВОЧНО")
    s.head({"label": "Показатель", "total": "Сумма", "note": "Комментарий"})
    for label, value, note in (
            ("Кредитная линия — долг на начало месяца", man.get("credit_debt_open"),
             "счёт 67, контрагент «ОМЕГА ООО»"),
            ("Кредитная линия — выбрано за месяц", man.get("credit_drawn"),
             "9 000 000 + 2 000 000, оба ушли Квалитету"),
            ("Кредитная линия — долг на конец месяца", man.get("credit_debt"), ""),
            ("Кредитная линия — лимит", man.get("credit_line_limit"), ""),
            ("Кредитная линия — свободный остаток", man.get("credit_line_available"), ""),
            ("Проценты по кредиту, уплачено", man.get("credit_interest_paid"),
             "счёт 67, контрагент «Филиал „Центральный“ Банка ВТБ»; в июле было 151 764,47"),
            ("Проценты по депозитам за месяц", man.get("deposit_interest"), ""),
            ("Остаток наличных в кассе", man.get("cash_on_hand"), "не проставлен"),
            ("Комиссия эквайринга", canon["revenue"].get("acquiring_fee_total"), ""),
            ("Онлайн-оплаты через сайт (интернет-эквайринг)", man.get("internet_acquiring"),
             "счёт 62.01; отдельной строкой не добавляется — уже внутри безналичных поступлений"),
            ("Авансы арендодателям на конец месяца",
             sum((man.get("landlord_advances") or {}).values()) or None,
             "дебетовое сальдо счёта 60: Паритет (стр.12) работает июльскую предоплату, "
             "поэтому коммунальных платежей в августе нет"),
    ):
        if value is not None or "касс" in label.lower():
            s.line(label, {"total": value}, note)

    ws.freeze_panes = "B5"
    ws.sheet_view.showGridLines = False

    # ---------------------------------------------------------------- контрагенты
    ws2 = wb.create_sheet("Расходы по контрагентам")
    s2 = Sheet(ws2)
    for col, w in ((1, 3), (2, 50), (3, 18), (4, 40)):
        ws2.column_dimensions[L(col)].width = w
    s2.title("Счёт 60: оплаты с расчётного счёта по контрагентам",
             "Отсортировано по убыванию суммы. Итог сходится с блоком «Расходы» основного листа.")
    s2.head({"label": "Контрагент", "total": "Оплачено", "mal": "Назначение"})
    first = s2.row
    for v in sorted(canon["expenses"]["vendors"], key=lambda x: -x["amount"]):
        s2._cell(2, v["name"])
        s2._cell(3, v["amount"], money=True)
        s2._cell(4, v.get("note", ""), size=9, color="807A70")
        s2.row += 1
    s2.total("ИТОГО ПО СЧЁТУ 60", {"total": f"=SUM(C{first}:C{s2.row-1})"},
             "в т.ч. за счёт кредитной линии — см. справочную строку основного листа")
    ws2.freeze_panes = "B5"
    ws2.sheet_view.showGridLines = False

    # ---------------------------------------------------------------- сверка с банком
    b = canon.get("bank_bridge")
    if b:
        ws5 = wb.create_sheet("Сверка с банком")
        s5 = Sheet(ws5)
        for col, w in ((1, 3), (2, 52), (3, 20), (4, 74)):
            ws5.column_dimensions[L(col)].width = w
        s5.title("Почему выручка не равна деньгам на счёте",
                 "Бухгалтерия считает по счёту 57 — это деньги, дошедшие до расчётного счёта. "
                 "Выручка — то, что пробито в управленке. Базы разные, вот переход.")
        s5.head({"label": "Шаг", "total": "Сумма", "mal": "Откуда"})
        t5 = canon["revenue"]["totals"]
        num5 = lambda v: f"{v:,.2f}".replace(",", " ")
        rows5 = [
            ("Поступило на р/с через счёт 57 (нетто)", b["net_57"],
             f"{num5(b['from57'])} с 57 на р/с минус {num5(b['transfers'])} переводов между своими счетами"),
            ("        инкассация наличных", b["inkas"], "счёт 50 → 57"),
            ("        эквайринг зачислен", b["acquiring_net"], "банк перечислил за вычетом комиссии"),
            ("+ Комиссия эквайринга, удержана банком", b["fee"],
             "счёт 57 → 91: пробито больше, чем зачислено"),
            ("Пробито по банку", b["net_57"] + b["fee"], "инкассация + эквайринг до комиссии"),
            ("− Инкассация больше пробитой налички", -(b["inkas"] - t5["cash"]),
             f"сдали {num5(b['inkas'])} при кассе {num5(t5['cash'])} — вместе с остатком кассы с июля"),
            ("− Эквайринг конца июля, зачисленный в августе", -(b["acq_gross"] - t5["card"]),
             f"по банку {num5(b['acq_gross'])} против {num5(t5['card'])} по картам"),
            ("Поступления от пациентов за август",
             b["net_57"] + b["fee"] - (b["inkas"] - t5["cash"]) - (b["acq_gross"] - t5["card"]),
             "до возвратов"),
            ("− Возвраты пациентам", -t5["refund"],
             "ушли с р/с напрямую через счёт 62, счёта 57 не касаются"),
        ]
        for label, v, note in rows5:
            s5.line(label, {"total": v}, note,
                    color=RED if v < 0 else "000000",
                    bold=label.startswith(("Пробито", "Поступ")))
        s5.total("ВЫРУЧКА В ОТЧЁТЕ", {"total": t5["total"]}, "сходится до копейки")
        ws5.freeze_panes = "B5"
        ws5.sheet_view.showGridLines = False

    # ---------------------------------------------------------------- статьи
    ws4 = wb.create_sheet("Статьи расходов")
    s4 = Sheet(ws4)
    for col, w in ((1, 3), (2, 46), (3, 18), (4, 12), (5, 40)):
        ws4.column_dimensions[L(col)].width = w
    s4.title("Все статьи расходов по счёту 60",
             "Полная расшифровка блока «Расходы» основного листа, без сворачивания")
    s4.head({"label": "Статья", "total": "Сумма", "mal": "Доля", "pog": "Блок"})
    grand = canon["expenses"]["totals"]["grand"] or 1
    first4 = None
    for block, items in (("медицинские", canon["expenses"]["medical"]),
                         ("управленческие", canon["expenses"]["management"])):
        s4.line(f"{block.capitalize()} — итого", {"total": sum(i["amount"] for i in items)}, bold=True)
        for item in items:
            first4 = first4 or s4.row
            s4._cell(2, "    " + item["name"])
            s4._cell(3, item["amount"], money=True)
            s4._cell(4, item["amount"] / grand)
            ws4.cell(row=s4.row, column=4).number_format = "0.0%"
            s4._cell(5, block, size=9, color="807A70")
            s4.row += 1
    s4.total("ИТОГО ПО СЧЁТУ 60", {"total": grand},
             "совпадает со строкой «Поставщики и подрядчики» основного листа")
    ws4.freeze_panes = "B5"
    ws4.sheet_view.showGridLines = False

    # ---------------------------------------------------------------- по дням
    daily = canon["revenue"].get("daily") or []
    if daily:
        ws3 = wb.create_sheet("Выручка по дням")
        s3 = Sheet(ws3)
        for col, w in ((1, 3), (2, 16), (3, 18), (4, 30)):
            ws3.column_dimensions[L(col)].width = w
        s3.title("Выручка по дням", "01–10 — старая управленка, с 11 — новая")
        s3.head({"label": "Дата", "total": "Сумма", "mal": "Источник"})
        first = s3.row
        for d in daily:
            s3._cell(2, d["date"])
            s3._cell(3, d["total"], money=True)
            s3._cell(4, "старая управленка" if d["day"] < 11 else "новая управленка",
                     size=9, color="807A70")
            s3.row += 1
        s3.total("ИТОГО ЗА МЕСЯЦ", {"total": f"=SUM(C{first}:C{s3.row-1})"})
        ws3.freeze_panes = "B5"
        ws3.sheet_view.showGridLines = False

    wb.save(out_path)
    return out_path


if __name__ == "__main__":
    src = Path(sys.argv[1] if len(sys.argv) > 1 else ROOT / "data/canonical/2026-08.json")
    canon = json.loads(src.read_text(encoding="utf-8"))
    name = f"Расчёт-{canon['meta']['month_ru']}-{canon['meta']['year']}.xlsx"
    print("готово:", build(canon, ROOT / "out" / name))
