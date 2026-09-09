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
        s.line("    Безналичными (эквайринг)", {"total": tot["card"]},
               (f"01–10.08: {sp['first_card']:,.0f}".replace(",", " ")
                + f"; 11–31.08: {sp['second_card']:,.0f}".replace(",", " ")) if sp else "")
    if tot.get("bank_ind"):
        s.line("    Оплаты физлиц на расчётный счёт", {"total": tot["bank_ind"]})
    if tot.get("refund"):
        s.line("    Возвраты пациентам", {"total": -tot["refund"]}, color=RED)
    s.total("Итого по каналам", {"total": f"=SUM(C{ch_first}:C{s.row-1})"},
            "должно совпасть с итогом выручки выше")

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
    sup_row = s.row
    s.line("Поставщики и подрядчики (счёт 60)",
           {"total": costs["suppliers_bank"]["total"],
            "mal": f"=C{sup_row}/2", "pog": f"=C{sup_row}/2"},
           "без оплаченного за счёт кредитной линии", bold=True)
    for block, items in (("медицинские и гонорары", r["expenses"]["medical"]),
                         ("управленческие", r["expenses"]["management"])):
        shown = [i for i in items if i["amount"] >= 300000]
        if not shown:
            continue
        s.line(f"        в том числе {block}:", color="807A70")
        for item in shown:
            s.line("            " + item["name"], {"total": item["amount"]})
        rest = sum(i["amount"] for i in items if i["amount"] < 300000)
        if rest:
            s.line("            прочие статьи блока", {"total": rest})
    exp_last = s.row - 1
    s.at["exp_end"] = exp_last
    row = s.row
    s.result("ИТОГО РАСХОДЫ",
             {"total": f"=SUM(C{exp_first}:C{sup_row})",
              "mal": f"=SUM(D{exp_first}:D{sup_row})", "pog": f"=SUM(E{exp_first}:E{sup_row})"},
             "строки «в том числе» в сумму не входят", key="exp")
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
           "итог блока «Расходы», пополам", color=RED)
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
            ("Кредитная линия — лимит", man.get("credit_line_limit"), ""),
            ("Кредитная линия — выбрано", man.get("credit_debt"), "долг на дату отчёта"),
            ("Кредитная линия — свободный остаток", man.get("credit_line_available"), ""),
            ("Проценты по депозитам за месяц", man.get("deposit_interest"), ""),
            ("Остаток наличных в кассе", man.get("cash_on_hand"), "не проставлен"),
            ("Комиссия эквайринга", canon["revenue"].get("acquiring_fee_total"), ""),
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
