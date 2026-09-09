# -*- coding: utf-8 -*-
"""Файл-расшифровка: откуда взялась каждая цифра отчёта.

Excel здесь — не источник данных, а аудиторский след: на каждую цифру
отчёта есть строка с формулой и ссылкой на источник.
"""
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from model import FOUNDERS, GROUP_RU, GROUPS, MONTHS_RU

INK = "1C1B18"
GREEN = "2E5E4E"
BRONZE = "8A7350"
TINT = "F1EFEA"
LINE = "E7E3D9"
RED = "A3423B"

H_FILL = PatternFill("solid", fgColor=INK)
H_FONT = Font(color="FFFFFF", bold=True, size=10)
SUB_FILL = PatternFill("solid", fgColor=TINT)
TOT_FILL = PatternFill("solid", fgColor="EDF1EE")
THIN = Side(style="thin", color=LINE)
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY = "#,##0.00;-#,##0.00;—"
MONEY0 = "#,##0"
PCT = "+0.0%;-0.0%"


def _header(ws, titles, widths, row=1):
    for i, (t, w) in enumerate(zip(titles, widths), start=1):
        c = ws.cell(row=row, column=i, value=t)
        c.fill, c.font, c.border = H_FILL, H_FONT, BOX
        c.alignment = Alignment(vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[row].height = 30
    ws.freeze_panes = ws.cell(row=row + 1, column=1)


def _row(ws, r, values, money_cols=(), bold=False, fill=None, fmt=MONEY):
    for i, v in enumerate(values, start=1):
        if isinstance(v, float):
            v = round(v, 2)
        if isinstance(v, str) and v.startswith("="):
            v = "'" + v
        c = ws.cell(row=r, column=i, value=v)
        c.border = BOX
        c.alignment = Alignment(vertical="top", wrap_text=(i > 3))
        if bold:
            c.font = Font(bold=True)
        if fill:
            c.fill = fill
        if i in money_cols:
            c.number_format = fmt
            c.alignment = Alignment(horizontal="right", vertical="top")
    return r + 1


def _title(ws, text, sub=""):
    ws["A1"] = text
    ws["A1"].font = Font(bold=True, size=14, color=INK)
    if sub:
        ws["A2"] = sub
        ws["A2"].font = Font(italic=True, size=10, color="807A70")
    return 4 if sub else 3


# --------------------------------------------------------------------------- 01
def _figures(c, r):
    """Список «цифра отчёта -> формула -> источник»."""
    src_bank = "БАНК (Анализ сч. 51), лист TDSheet"
    src_rev = f"ВЫРУЧКА, лист «{c['meta']['month_ru']} {c['meta']['year']}»"
    src_exp = "РАСХОДЫ (Обороты сч. 60), лист 1"
    costs = c["costs"]
    ref = c["founders_reference"]
    rv, tr, ex = r["revenue"], r["trend"], r["expenses"]
    f = r["founders"]

    def brow(key):
        row = (costs.get(key) or {}).get("row")
        return f"{src_bank}, стр. {row}" if row else src_bank

    rows = [
        ("Титул", "Выручка за месяц", rv["total"], "₽",
         "касса + эквайринг − возвраты", f"{src_bank}, блок «ИТОГО ВЫРУЧКА»"),
        ("Титул", "Доход учредителей за месяц", r["founders_total_income"], "₽",
         "доход Маланичева + доход Погосяна (методика 50/50)", "лист «07 Учредители»"),
        ("Титул", "Остаток на р/с на конец месяца", r["cashflow"]["closing"], "₽",
         "конечное сальдо счёта 51", f"{src_bank}, «Конечное сальдо»"),
        ("Титул", "Рост выручки с начала года", tr["ytd_growth"], "%",
         "(выручка янв–отч. мес. тек. года / тот же период пр. года − 1) × 100",
         "ВЫРУЧКА, строка «Выручка за месяц» на каждом листе"),
        ("KPI", "Выручка месяца, год к году", tr["yoy_month"], "%",
         "(выручка тек. месяца / тот же месяц пр. года − 1) × 100", "ВЫРУЧКА, помесячные листы"),
        ("KPI", "Выручка с начала года", tr["ytd_cur"], "₽",
         "сумма строк «Выручка за месяц» за янв–отчётный месяц", "ВЫРУЧКА, помесячные листы"),
        ("KPI", "Средняя выручка в день", rv.get("avg_day"), "₽",
         "сумма «Всего за день» / число дней в месяце", f"{src_rev}, колонка «Всего за день»"),
        ("KPI", "Оплаты поставщикам (сч. 60)", ex["acc60_total"], "₽",
         "итог оборотов по счёту 60 за месяц", f"{src_exp}, строка «Итого»"),
        ("KPI", "ФОТ с налогами по банку", r["payroll"]["salary"] + r["payroll"]["taxes"], "₽",
         "зарплата по банку + налоги с ФОТ (68.90 + 69)", brow("salary")),
        ("KPI", "Остаток на р/с на конец месяца", r["cashflow"]["closing"], "₽",
         "конечное сальдо счёта 51", src_bank),
        ("Структура", "Касса (наличные)", rv["cash"], "₽", "строка «касса» сводного блока выручки", src_bank),
        ("Структура", "Эквайринг (карты и онлайн-оплаты)", rv["card"], "₽",
         "оборот сч. 57.03: розница и интернет-эквайринг", src_bank),
        ("Структура", "Возвраты пациентам", -rv["refund"], "₽", "строка «возврат пациенту ч/з банк и кассу»", src_bank),
        ("Структура", "Комиссия эквайринга", rv["acquiring_fee"], "₽",
         "сумма комиссий по группам врачей", f"{src_rev}, справочный блок «услуги Банка в Отчет»"),
        ("Расходы", "Медицинские расходы + ИП", ex["medical_total"], "₽",
         "итог блока «МЕДИЦИНСКИЕ РАСХОДЫ+ИП»", f"{src_exp}, колонки E–F"),
        ("Расходы", "Управленческие расходы", ex["management_total"], "₽",
         "итог блока «УПРАВЛЕНЧЕСКИЕ РАСХОДЫ»", f"{src_exp}, колонки E–F"),
        ("Расходы", "в т.ч. гонорары ИП хирургов", ex["ip_surgeons"], "₽",
         "статья «ИП хирурги» медицинского блока", brow("ip_surgeons")),
        ("Расходы", "в т.ч. ремонт нового корпуса", ex["repair"], "₽",
         "статья «Ремонт Самокатная 1 стр.12»", f"{src_exp}, управленческий блок"),
        ("Экономика", "Доход месяца (итог водопада)", r["waterfall_result"], "₽",
         "выручка − все статьи расходов по шагам водопада", "лист «08 Водопад»"),
        ("ДДС", "Остаток на начало месяца", r["cashflow"]["opening"], "₽", "начальное сальдо сч. 51", src_bank),
        ("ДДС", "Поступления за месяц", r["cashflow"]["in_total"], "₽", "оборот по дебету сч. 51", src_bank),
        ("ДДС", "Списания за месяц", r["cashflow"]["out_total"], "₽", "оборот по кредиту сч. 51", src_bank),
        ("ДДС", "Остаток на конец месяца", r["cashflow"]["closing"], "₽", "конечное сальдо сч. 51", src_bank),
    ]
    for key in FOUNDERS:
        d = f[key]
        rows += [
            ("Учредители", f"{d['name']}: личная выручка", d["own_revenue"], "₽",
             "выручка своей группы (касса + эквайринг + р/с − возвраты)", src_bank),
            ("Учредители", f"{d['name']}: ½ выручки других хирургов", d["half_other"], "₽",
             "выручка «Другие хирурги» / 2", src_bank),
            ("Учредители", f"{d['name']}: ½ косметологии и прочего",
             d["half_cosm"] + d["half_misc"], "₽", "(косметология + прочие доходы) / 2", src_bank),
            ("Учредители", f"{d['name']}: доля расходов", -d["costs"], "₽",
             "½ всех расходов месяца + личные анализы учредителя",
             brow("total")),
            ("Учредители", f"{d['name']}: доход за месяц", d["income"], "₽",
             "личная выручка − доля расходов + ½ др.хирургов + ½ косметологии + ½ прочего",
             "лист «07 Учредители»"),
            ("Учредители", f"{d['name']}: остаток с прошлого месяца", d["prev_balance"], "₽",
             "ручной ввод бухгалтерии (переносящийся долг компании)",
             f"{src_bank}, стр. {ref['prev_balance']['row']}" if ref.get("prev_balance") else src_bank),
            ("Учредители", f"{d['name']}: удержано в резерв (аренда)", -d["reserve"], "₽",
             "ручной ввод бухгалтерии (депозит под аренду)", src_bank),
            ("Учредители", f"{d['name']}: к выплате на конец месяца", d["payout"], "₽",
             "доход за месяц + остаток с прошлого месяца − резерв",
             f"{src_bank}, стр. {ref['payout']['row']}" if ref.get("payout") else src_bank),
        ]
    man = r["manual"]
    rows += [
        ("Справочно", "Долг по кредитной линии", man.get("credit_debt"), "₽",
         "ручной ввод бухгалтерии", src_bank),
        ("Справочно", "Остаток наличных в кассе", man.get("cash_on_hand"), "₽",
         "ручной ввод бухгалтерии", src_bank),
        ("Справочно", "Проценты по депозитам", man.get("deposit_interest"), "₽",
         "оборот сч. 91.01 «% по депозитам»", src_bank),
        ("Справочно", "Зарезервировано на аренду (всего)", man.get("reserve_total"), "₽",
         "ручной ввод бухгалтерии", src_bank),
    ]
    return rows


def build(c, r, path):
    wb = Workbook()

    # ---------------------------------------------------------------- 00
    ws = wb.active
    ws.title = "00 Навигация"
    top = _title(ws, f"Расшифровка отчёта учредителям · {c['meta']['period']}",
                 "Excel-приложение к отчёту: на каждую цифру — формула и ссылка на источник")
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 90
    nav = [
        ("01 Цифры отчёта", "Каждая цифра из HTML-отчёта: значение, формула, источник. Главный лист расшифровки."),
        ("02 Выручка по дням", "Дневная выручка по группам врачей и каналам оплаты + сверка с онлайн-кассой."),
        ("03 Выручка структура", "Итоги по группам врачей и каналам поступления, доли в выручке, комиссия эквайринга."),
        ("04 Расходы категории", "Счёт 60 в разрезе статей: медицинские + ИП и управленческие."),
        ("05 Расходы контрагенты", "Счёт 60 построчно по контрагентам с назначением платежа."),
        ("06 ДДС сч.51", "Обороты расчётного счёта: поступления и списания по корреспондирующим счетам."),
        ("07 Учредители", "Методика 50/50 по шагам: выручка, доли расходов, доход, остаток, резерв, к выплате."),
        ("08 Водопад", "От выручки к доходу месяца — те же шаги, что на графике «Куда уходит выручка»."),
        ("09 Контроль", "Контрольные равенства: сходится ли отчёт сам с собой и с данными бухгалтерии."),
    ]
    rr = top
    _header(ws, ["Лист", "Что внутри"], [26, 90], row=rr)
    rr += 1
    for name, desc in nav:
        rr = _row(ws, rr, [name, desc])
    rr += 1
    ws.cell(row=rr, column=1, value="Источники").font = Font(bold=True)
    rr += 1
    for s in [
        "БАНК — «Анализ счёта 51» за месяц + расчётный блок методики учредителей",
        "ВЫРУЧКА — дневная выручка по группам врачей и каналам оплаты, сверка с онлайн-кассой",
        "РАСХОДЫ — «Обороты счёта 60» за месяц по контрагентам",
    ]:
        ws.cell(row=rr, column=1, value=s)
        rr += 1

    # ---------------------------------------------------------------- 01
    ws = wb.create_sheet("01 Цифры отчёта")
    top = _title(ws, "Цифры отчёта: значение → формула → источник")
    _header(ws, ["Раздел отчёта", "Показатель", "Значение", "Ед.", "Как получено", "Источник"],
            [16, 42, 18, 6, 60, 46], row=top)
    rr = top + 1
    for section, name, value, unit, formula, source in _figures(c, r):
        rr = _row(ws, rr, [section, name, value, unit, formula, source], money_cols=(3,))

    # ---------------------------------------------------------------- 02
    ws = wb.create_sheet("02 Выручка по дням")
    top = _title(ws, "Выручка по дням", "Касса — наличные, Банк — эквайринг. Разница с онлайн-кассой — не ошибка, см. лист 09")
    cols = ["Дата"]
    for g in GROUPS:
        cols += [f"{GROUP_RU[g]} · касса", f"{GROUP_RU[g]} · банк"]
    cols += ["Всего за день", "Из кассового отчёта", "Разница (онлайн-касса)"]
    _header(ws, cols, [12] + [15] * (len(cols) - 1), row=top)
    rr = top + 1
    money = tuple(range(2, len(cols) + 1))
    for d in c["revenue"]["daily"]:
        vals = [d["date"]]
        for g in GROUPS:
            vals += [d["cash"].get(g, 0), d["card"].get(g, 0)]
        vals += [d["total"], d.get("kkt"), d.get("kkt_diff")]
        rr = _row(ws, rr, vals, money_cols=money)
    tot = ["ИТОГО"]
    for g in GROUPS:
        tot += [sum(d["cash"].get(g, 0) for d in c["revenue"]["daily"]),
                sum(d["card"].get(g, 0) for d in c["revenue"]["daily"])]
    tot += [sum(d["total"] for d in c["revenue"]["daily"]),
            sum(d.get("kkt") or 0 for d in c["revenue"]["daily"]),
            sum(d.get("kkt_diff") or 0 for d in c["revenue"]["daily"])]
    _row(ws, rr, tot, money_cols=money, bold=True, fill=TOT_FILL)

    # ---------------------------------------------------------------- 03
    ws = wb.create_sheet("03 Выручка структура")
    t = c["revenue"]["totals"]
    bg = c["revenue"]["by_group"]
    # каналы по группам известны не всегда: в августе управленка даёт только
    # итог по врачу, поэтому колонки касса/эквайринг показываем, если они заполнены
    chan = any(bg[g][k] for g in GROUPS for k in ("cash", "card", "bank_ind"))
    sep = bool(t.get("bank_ind"))
    top = _title(ws, "Структура выручки",
                 "Итог = касса + эквайринг + оплаты физлиц на р/с − возвраты" if sep
                 else "Итог = касса + эквайринг (карты и онлайн-оплаты) − возвраты")
    chan_head = (["Касса", "Эквайринг"] + (["Оплаты на р/с"] if sep else [])) if chan else []
    head = ["Группа"] + chan_head + ["Возвраты", "Итого", "Доля", "Комиссия эквайринга"]
    widths = [24] + [16] * len(chan_head) + [14, 18, 10, 20]
    _header(ws, head, widths, row=top)
    money = tuple(range(2, len(head) - 1)) + (len(head),)
    pct_col = len(head) - 1
    rr = top + 1
    for g in GROUPS:
        b = bg[g]
        vals = ([b["cash"], b["card"]] + ([b["bank_ind"]] if sep else [])) if chan else []
        rr = _row(ws, rr, [GROUP_RU[g]] + vals
                  + [-b["refund"], b["total"], r["revenue"]["share"][g] / 100,
                     (c["revenue"].get("acquiring_fee") or {}).get(g)], money_cols=money)
        ws.cell(row=rr - 1, column=pct_col).number_format = "0.0%"
    vals = ([t["cash"], t["card"]] + ([t["bank_ind"]] if sep else [])) if chan else []
    rr = _row(ws, rr, ["ИТОГО"] + vals
              + [-t["refund"], t["total"], 1.0, c["revenue"].get("acquiring_fee_total")],
              money_cols=money, bold=True, fill=TOT_FILL)
    ws.cell(row=rr - 1, column=pct_col).number_format = "0.0%"

    if not chan:                       # каналы отдельным блоком, раз по группам их нет
        rr += 1
        ws.cell(row=rr, column=1, value="Каналы поступления").font = Font(bold=True)
        rr += 1
        _header(ws, ["Канал", "Сумма", "Комментарий"], [24, 18, 56], row=rr)
        rr += 1
        sp = t.get("split") or {}
        online = (f" (в т.ч. онлайн-оплаты через сайт {sp['first_online']:,.0f} ₽)".replace(",", " ")
                  if sp.get("first_online") else "")
        for label, value, note in (
                ("Наличными", t["cash"], "касса клиники"),
                ("Безналичными", t["card"], "карты и онлайн-оплаты" + online),
                ("Оплаты физлиц на р/с", t["bank_ind"] if sep else None, ""),
                ("Возвраты пациентам", -t["refund"], "вычитаются из выручки"),
        ):
            if value is None:
                continue
            rr = _row(ws, rr, [label, value, note], money_cols=(2,))
        rr = _row(ws, rr, ["ИТОГО", t["total"], ""], money_cols=(2,), bold=True, fill=TOT_FILL)
    rr += 1
    ws.cell(row=rr, column=1, value="Динамика по месяцам (база: строка «Выручка за месяц», кассовый метод)").font = Font(bold=True)
    rr += 1
    _header(ws, ["Месяц", str(c["meta"]["year"] - 1), str(c["meta"]["year"]), "Изменение"], [16, 18, 18, 14], row=rr)
    rr += 1
    for i in range(12):
        p, cu = r["trend"]["prev"][i], r["trend"]["cur"][i]
        if p is None and cu is None:
            continue
        delta = (cu / p - 1) if (p and cu) else None
        rr = _row(ws, rr, [MONTHS_RU[i], p, cu, delta], money_cols=(2, 3))
        ws.cell(row=rr - 1, column=4).number_format = PCT

    # ---------------------------------------------------------------- 04
    ws = wb.create_sheet("04 Расходы категории")
    top = _title(ws, "Расходы по счёту 60 в разрезе статей")
    _header(ws, ["Блок", "Статья", "Сумма", "Доля в блоке"], [26, 52, 18, 14], row=top)
    rr = top + 1
    for block, items, total in (("Медицинские + ИП", c["expenses"]["medical"], r["expenses"]["medical_total"]),
                                ("Управленческие", c["expenses"]["management"], r["expenses"]["management_total"])):
        for it in items:
            rr = _row(ws, rr, [block, it["name"], it["amount"], (it["amount"] / total if total else None)],
                      money_cols=(3,))
            ws.cell(row=rr - 1, column=4).number_format = "0.0%"
        rr = _row(ws, rr, [block, "ИТОГО", total, 1.0], money_cols=(3,), bold=True, fill=TOT_FILL)
        ws.cell(row=rr - 1, column=4).number_format = "0.0%"
    _row(ws, rr, ["ВСЕГО ПО СЧЁТУ 60", "", r["expenses"]["acc60_total"], None],
         money_cols=(3,), bold=True, fill=SUB_FILL)

    # ---------------------------------------------------------------- 05
    ws = wb.create_sheet("05 Расходы контрагенты")
    top = _title(ws, "Счёт 60 построчно по контрагентам")
    _header(ws, ["Контрагент", "Сумма", "Назначение / комментарий"], [52, 18, 60], row=top)
    rr = top + 1
    for v in sorted(c["expenses"]["vendors"], key=lambda x: -x["amount"]):
        rr = _row(ws, rr, [v["name"], v["amount"], v.get("note", "")], money_cols=(2,))
    _row(ws, rr, ["ИТОГО", sum(v["amount"] for v in c["expenses"]["vendors"]), ""],
         money_cols=(2,), bold=True, fill=TOT_FILL)

    # ---------------------------------------------------------------- 06
    ws = wb.create_sheet("06 ДДС сч.51")
    top = _title(ws, "Движение денежных средств по расчётному счёту")
    _header(ws, ["Кор. счёт / статья", "Поступление", "Списание", "Комментарий"], [34, 20, 20, 52], row=top)
    rr = top + 1
    rr = _row(ws, rr, ["Начальное сальдо", r["cashflow"]["opening"], None, ""], money_cols=(2, 3), bold=True, fill=SUB_FILL)
    for fl in c["cashflow"]["flows"]:
        rr = _row(ws, rr, [fl["label"], fl.get("debit"), fl.get("credit"), fl.get("note", "")], money_cols=(2, 3))
    rr = _row(ws, rr, ["Оборот за месяц", r["cashflow"]["in_total"], r["cashflow"]["out_total"], ""],
              money_cols=(2, 3), bold=True, fill=TOT_FILL)
    _row(ws, rr, ["Конечное сальдо", r["cashflow"]["closing"], None, "= начальное + поступления − списания"],
         money_cols=(2, 3), bold=True, fill=SUB_FILL)

    # ---------------------------------------------------------------- 07
    ws = wb.create_sheet("07 Учредители")
    top = _title(ws, "Доход учредителей — методика по шагам",
                 "Доход = личная выручка + ½ выручки других хирургов + ½ косметологии и прочего − доля расходов")
    _header(ws, ["Шаг", "Показатель", GROUP_RU["mal"], GROUP_RU["pog"], "Всего", "Комментарий"],
            [6, 46, 18, 18, 18, 46], row=top)
    rr = top + 1
    f = r["founders"]
    steps = [
        ("1", "Личная выручка (свои операции)", "own_revenue", "выручка своей группы"),
        ("2", "+ ½ выручки других хирургов", "half_other", "выручка «Другие хирурги» / 2"),
        ("3", "+ ½ косметологии", "half_cosm", "выручка косметологии / 2"),
        ("4", "+ ½ прочих доходов", "half_misc", "прочие доходы / 2"),
        ("5", "+ ½ процентов по депозитам", "deposit_share", "проценты по депозитам / 2"),
    ]
    for n, label, key, note in steps:
        a, b = f["mal"].get(key, 0.0), f["pog"].get(key, 0.0)
        if not (a or b) and key == "deposit_share":
            continue
        rr = _row(ws, rr, [n, label, a, b, a + b, note], money_cols=(3, 4, 5))
    cost_labels = {
        "payroll_taxes": "налоги с ФОТ (68.90 + 69)", "salary": "зарплата по банку",
        "cash_expenses": "расходы по кассе", "bank_common": "услуги банка (общие)",
        "bank_acquiring": "услуги банка (эквайринг)", "alimony": "алименты (удержание из ЗП)",
        "credit_interest": "проценты по кредитной линии", "adjust": "дополнительные корректировки",
        "suppliers_bank": "поставщики (счёт 60) + Халва",
    }
    pers_note = ("½ всех расходов месяца; личные анализы (Маланичев {}, Погосян {}) "
                 "каждый несёт полностью — переносится половина разницы").format(
                     f"{f['mal'].get('personal_costs', 0):,.0f}".replace(",", " "),
                     f"{f['pog'].get('personal_costs', 0):,.0f}".replace(",", " ")) \
        if (f["mal"].get("personal_costs") or f["pog"].get("personal_costs")) else "½ всех расходов месяца"
    rr = _row(ws, rr, ["6", "− Доля расходов, в том числе:", -f["mal"]["costs"], -f["pog"]["costs"],
                       -(f["mal"]["costs"] + f["pog"]["costs"]), pers_note],
              money_cols=(3, 4, 5), bold=True)
    for k, label in cost_labels.items():
        a, b = f["mal"]["costs_detail"].get(k, 0), f["pog"]["costs_detail"].get(k, 0)
        rr = _row(ws, rr, ["", f"      {label}", -a, -b, -(a + b), ""], money_cols=(3, 4, 5))

    rr = _row(ws, rr, ["7", "ДОХОД ЗА МЕСЯЦ", f["mal"]["income"], f["pog"]["income"],
                       r["founders_total_income"], "сумма шагов выше"],
              money_cols=(3, 4, 5), bold=True, fill=TOT_FILL)
    rr = _row(ws, rr, ["8", "+ Остаток с прошлого месяца", f["mal"]["prev_balance"], f["pog"]["prev_balance"],
                       f["mal"]["prev_balance"] + f["pog"]["prev_balance"], "ручной ввод бухгалтерии"],
              money_cols=(3, 4, 5))
    step = 9
    for label, key, note in (("− Выплачены дивиденды", "dividends", "ручной ввод бухгалтерии"),
                             ("− Удержано в резерв (аренда)", "reserve", "ручной ввод бухгалтерии"),
                             ("+ Возврат остатка резерва", "reserve_return", "ручной ввод бухгалтерии")):
        a, b = f["mal"].get(key, 0.0), f["pog"].get(key, 0.0)
        if not (a or b):
            continue
        sign = 1 if key == "reserve_return" else -1
        rr = _row(ws, rr, [str(step), label, sign * a, sign * b, sign * (a + b), note],
                  money_cols=(3, 4, 5))
        step += 1
    _row(ws, rr, [str(step), "К ВЫПЛАТЕ НА КОНЕЦ МЕСЯЦА", f["mal"]["payout"], f["pog"]["payout"],
                  f["mal"]["payout"] + f["pog"]["payout"], "доход + остаток − дивиденды − резерв"],
         money_cols=(3, 4, 5), bold=True, fill=TOT_FILL)

    # ---------------------------------------------------------------- 08
    ws = wb.create_sheet("08 Водопад")
    top = _title(ws, "Куда уходит выручка", "Каждый шаг уменьшает накопленный остаток; последний столбец — доход месяца")
    _header(ws, ["Шаг", "Статья", "Сумма", "Накоплено после шага", "Копеек с рубля выручки"],
            [6, 42, 20, 22, 22], row=top)
    rr = top + 1
    total_rev = r["revenue"]["total"]
    running = 0.0
    for i, s in enumerate(r["waterfall"]):
        if i == 0:
            running = s["value"]
        elif i < len(r["waterfall"]) - 1:
            running += s["value"]
        share = abs(s["value"]) / total_rev * 100 if total_rev else 0
        rr = _row(ws, rr, [i + 1, s["name"], s["value"], running if i < len(r["waterfall"]) - 1 else s["value"],
                           round(share, 1)],
                  money_cols=(3, 4),
                  bold=(s["kind"] == "pos"),
                  fill=TOT_FILL if s["kind"] == "pos" else None)

    # ---------------------------------------------------------------- 09
    ws = wb.create_sheet("09 Контроль")
    top = _title(ws, "Контрольные сверки", "Проверки из брифа: отчёт должен сходиться сам с собой и с данными бухгалтерии")
    _header(ws, ["Проверка", "Слева", "Справа", "Разница", "Итог", "Комментарий"],
            [56, 20, 20, 16, 12, 52], row=top)
    rr = top + 1
    for k in r["checks"]:
        info = "разные базы" in k.get("note", "")
        verdict = "справочно" if info else ("сходится" if k["ok"] else "РАСХОЖДЕНИЕ")
        rr = _row(ws, rr, [k["name"], k["left"], k["right"], k["delta"], verdict, k.get("note", "")],
                  money_cols=(2, 3, 4))
        cell = ws.cell(row=rr - 1, column=5)
        cell.font = Font(bold=True, color=(BRONZE if info else (GREEN if k["ok"] else RED)))

    # ---------------------------------------------------------------- 10
    mig = c.get("migration")
    if mig:
        ws = wb.create_sheet("10 Переход управленки")
        top = _title(ws, "Переход на новую управленку",
                     "Как схлопнуты платежи дня перехода: что оставлено и что снято как дубль")
        _header(ws, ["Параметр", "Значение"], [46, 60], row=top)
        rr = top + 1
        for k, v in (("Дата перехода", mig.get("switch_date") or "не указана"),
                     ("Приоритетная система в день перехода", mig.get("priority")),
                     ("Платежей в выгрузках всего", mig.get("payments_total")),
                     ("Платежей учтено в отчёте", mig.get("payments_kept")),
                     ("Снято как дубли", len(mig.get("dropped", [])))):
            rr = _row(ws, rr, [k, v])
        rr += 1
        if mig.get("dropped"):
            ws.cell(row=rr, column=1, value="Снятые платежи (дубли дня перехода)").font = Font(bold=True)
            rr += 1
            _header(ws, ["ID платежа", "Дата", "Система", "Пациент", "Врач", "Сумма", "Причина"],
                    [18, 14, 14, 28, 28, 16, 52], row=rr)
            rr += 1
            for d in mig["dropped"]:
                rr = _row(ws, rr, [d.get("id"), d.get("date"), d.get("system"), d.get("patient"),
                                   d.get("doctor"), d.get("amount"), d.get("reason")], money_cols=(6,))
        else:
            ws.cell(row=rr, column=1,
                    value="Дублей не найдено — либо перехода в этом месяце не было, "
                          "либо платежи не пересекаются.")

    wb.save(path)
    return path
