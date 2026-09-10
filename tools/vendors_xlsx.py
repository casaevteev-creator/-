# -*- coding: utf-8 -*-
"""Классификация контрагентов счёта 60 с живыми формулами.

Три листа: контрагенты (что во что относится), статьи (SUMIF по контрагентам),
свод (SUMIF по статьям). Меняешь блок или статью у контрагента — пересчитывается всё.
"""
import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent.parent
MONEY = "#,##0.00;-#,##0.00;—"
INK, BRONZE, GREY = "1A1613", "8A6A3C", "807A70"
HEAD = PatternFill("solid", fgColor="EFEAE2")
TOT = PatternFill("solid", fgColor="F6F1E8")
THIN = Side(style="thin", color="D9D2C6")

BLOCKS = ["медицинские", "управленческие", "услуги банка", "за счёт кредитной линии"]


def _head(ws, row, titles, widths):
    for i, (t, w) in enumerate(zip(titles, widths), start=2):
        c = ws.cell(row=row, column=i, value=t)
        c.font = Font(bold=True, size=10, color=INK)
        c.fill = HEAD
        c.alignment = Alignment(vertical="center", wrap_text=True)
        c.border = Border(bottom=THIN)
        ws.column_dimensions[L(i)].width = w
    ws.row_dimensions[row].height = 26


def _title(ws, text, sub):
    ws["B2"] = text
    ws["B2"].font = Font(bold=True, size=14, color=INK)
    ws["B3"] = sub
    ws["B3"].font = Font(size=9, color=GREY)
    ws.column_dimensions["A"].width = 3
    ws.freeze_panes = "B6"
    ws.sheet_view.showGridLines = False


def build(canon, out):
    ex = canon["expenses"]
    med = {a["name"] for a in ex["medical"]}
    mgmt = {a["name"] for a in ex["management"]}
    credit_article = "Ремонт Самокатная 1 стр12"

    def block(article):
        if article in med:
            return BLOCKS[0]
        if article in mgmt:
            return BLOCKS[1]
        if article == "Услуги банка":
            return BLOCKS[2]
        return BLOCKS[3] if article == credit_article else BLOCKS[1]

    rows = sorted(ex["vendors"], key=lambda v: -v["amount"])
    wb = Workbook()

    # ---------------------------------------------------------------- контрагенты
    ws = wb.active
    ws.title = "Контрагенты"
    _title(ws, "Контрагенты счёта 60 — что к какой статье относится",
           "Меняете статью или блок в этой таблице — суммы на листах «Статьи» и «Свод» "
           "пересчитываются сами. Столбцы «Блок» — выпадающий список.")
    _head(ws, 5, ["Контрагент", "Оплачено", "Статья", "Блок"], [46, 18, 34, 24])
    first = 6
    for i, v in enumerate(rows):
        r = first + i
        art = v.get("note") or "Прочие управленческие"
        ws.cell(row=r, column=2, value=v["name"]).font = Font(size=10, color=INK)
        c = ws.cell(row=r, column=3, value=v["amount"])
        c.number_format = MONEY
        c.font = Font(size=10, color=INK)
        ws.cell(row=r, column=4, value=art).font = Font(size=10, color=INK)
        ws.cell(row=r, column=5, value=block(art)).font = Font(size=10, color=BRONZE)
        for col in range(2, 6):
            ws.cell(row=r, column=col).border = Border(bottom=THIN)
    last = first + len(rows) - 1
    r = last + 1
    ws.cell(row=r, column=2, value="ИТОГО ПО СЧЁТУ 60").font = Font(bold=True, size=10, color=INK)
    c = ws.cell(row=r, column=3, value=f"=SUM(C{first}:C{last})")
    c.number_format = MONEY
    c.font = Font(bold=True, size=10, color=INK)
    for col in range(2, 6):
        ws.cell(row=r, column=col).fill = TOT

    arts = sorted(med | mgmt | {"Услуги банка", credit_article})
    dv_b = DataValidation(type="list", formula1='"' + ",".join(BLOCKS) + '"', showDropDown=False)
    ws.add_data_validation(dv_b)
    dv_b.add(f"E{first}:E{last}")

    # ---------------------------------------------------------------- статьи
    ws2 = wb.create_sheet("Статьи")
    _title(ws2, "Статьи расходов", "Сумма каждой статьи — SUMIF по листу «Контрагенты». "
                                   "Формулу видно в ячейке.")
    _head(ws2, 5, ["Статья", "Блок", "Сумма", "Доля", "Контрагентов"], [40, 24, 18, 10, 14])
    src = f"Контрагенты!$D${first}:$D${last}"
    amt = f"Контрагенты!$C${first}:$C${last}"
    r2 = 6
    for blk in BLOCKS:
        items = sorted({a for a in arts if block(a) == blk})
        if not items:
            continue
        c = ws2.cell(row=r2, column=2, value=blk.upper())
        c.font = Font(bold=True, size=10, color=BRONZE)
        r2 += 1
        start = r2
        for art in items:
            c = ws2.cell(row=r2, column=2, value=art)
            c.font = Font(size=10, color=INK)
            c.alignment = Alignment(indent=2)
            ws2.cell(row=r2, column=3, value=blk).font = Font(size=9, color=GREY)
            c = ws2.cell(row=r2, column=4, value=f'=SUMIF({src},B{r2}&"",{amt})')
            c.number_format = MONEY
            c.font = Font(size=10, color=INK)
            c = ws2.cell(row=r2, column=5, value=f"=D{r2}/Контрагенты!$C${last + 1}")
            c.number_format = "0.0%"
            c.font = Font(size=9, color=GREY)
            c = ws2.cell(row=r2, column=6, value=f'=COUNTIF({src},B{r2}&"")')
            c.font = Font(size=9, color=GREY)
            for col in range(2, 7):
                ws2.cell(row=r2, column=col).border = Border(bottom=THIN)
            r2 += 1
        c = ws2.cell(row=r2, column=2, value=f"Итого «{blk}»")
        c.font = Font(bold=True, size=10, color=INK)
        c.alignment = Alignment(indent=2)
        c = ws2.cell(row=r2, column=4, value=f"=SUM(D{start}:D{r2 - 1})")
        c.number_format = MONEY
        c.font = Font(bold=True, size=10, color=INK)
        for col in range(2, 7):
            ws2.cell(row=r2, column=col).fill = TOT
        r2 += 2

    # ---------------------------------------------------------------- служебное
    # список для выпадающего меню: блоки, потом статьи. На отдельном листе,
    # потому что в названиях статей есть запятые — во встроенный список их нельзя
    wsl = wb.create_sheet("Списки")
    picks = BLOCKS + arts
    for i, v in enumerate(picks, start=1):
        wsl.cell(row=i, column=1, value=v)
    wsl.column_dimensions["A"].width = 40
    wsl.sheet_state = "hidden"

    # счётчик строк выборки на листе контрагентов: 1, 2, 3… у подходящих, пусто у прочих
    c = ws.cell(row=4, column=7, value="№ в выборке →")
    c.font = Font(size=9, color=GREY)
    c.alignment = Alignment(horizontal="center")
    ws.column_dimensions["G"].width = 13
    ws.cell(row=first - 1, column=7, value=0).font = Font(size=9, color="EFEAE2")  # затравка, не видна
    for r in range(first, last + 1):
        c = ws.cell(row=r, column=7,
                    value=f'=IF(OR($D{r}=Свод!$C$5,$E{r}=Свод!$C$5),MAX($G${first - 1}:$G{r - 1})+1,"")')
        c.font = Font(size=9, color=GREY)
        c.alignment = Alignment(horizontal="center")

    # ---------------------------------------------------------------- свод
    ws3 = wb.create_sheet("Свод")
    _title(ws3, "Свод и выборка",
           "Сверху итоги по блокам. Ниже выберите блок или статью в выпадающем списке — "
           "под ним развернётся список контрагентов, которые в неё входят.")
    blk_src = f"Контрагенты!$E${first}:$E${last}"
    art_src = f"Контрагенты!$D${first}:$D${last}"
    notes = {
        BLOCKS[0]: "медикаменты, импланты, расходка, гонорары ИП хирургов",
        BLOCKS[1]: "аренда, маркетинг, IT, хозяйственные",
        BLOCKS[2]: "комиссии банков, проведённые через счёт 60 — ошибка проводки",
        BLOCKS[3]: "оплачено заёмными, в расходы учредителей не входит",
    }

    # ---- выборка
    ws3.cell(row=5, column=2, value="Блок или статья:").font = Font(bold=True, size=11, color=INK)
    pick = ws3.cell(row=5, column=3, value=BLOCKS[0])
    pick.font = Font(bold=True, size=12, color=BRONZE)
    pick.fill = PatternFill("solid", fgColor="FFF6E6")
    pick.border = Border(bottom=Side(style="medium", color=BRONZE))
    dv = DataValidation(type="list", formula1=f"=Списки!$A$1:$A${len(picks)}", showDropDown=False)
    ws3.add_data_validation(dv)
    dv.add("C5")

    ws3.cell(row=6, column=2, value="Сумма:").font = Font(size=10, color=GREY)
    c = ws3.cell(row=6, column=3,
                 value=f'=SUMIF({blk_src},$C$5,{amt})+SUMIF({art_src},$C$5,{amt})')
    c.number_format = MONEY
    c.font = Font(bold=True, size=12, color=INK)
    ws3.cell(row=6, column=4, value="Контрагентов:").font = Font(size=10, color=GREY)
    c = ws3.cell(row=6, column=5,
                 value=f'=COUNTIF({blk_src},$C$5)+COUNTIF({art_src},$C$5)')
    c.font = Font(bold=True, size=12, color=INK)

    _head(ws3, 8, ["Контрагент", "Оплачено", "Статья", "Блок"], [46, 18, 34, 24])
    key = f"Контрагенты!$G${first}:$G${last}"
    for i in range(70):
        r = 9 + i
        for col, src_col in ((2, "B"), (3, "C"), (4, "D"), (5, "E")):
            c = ws3.cell(row=r, column=col, value=(
                f'=IFERROR(INDEX(Контрагенты!${src_col}${first}:${src_col}${last},'
                f'MATCH({i + 1},{key},0)),"")'))
            c.font = Font(size=10, color=INK if col < 4 else GREY)
            c.border = Border(bottom=THIN)
            if col == 3:
                c.number_format = MONEY
    r = 9 + 70
    ws3.cell(row=r, column=2, value="ИТОГО ПО ВЫБОРКЕ").font = Font(bold=True, size=10, color=INK)
    c = ws3.cell(row=r, column=3, value=f"=SUM(C9:C{r - 1})")
    c.number_format = MONEY
    c.font = Font(bold=True, size=10, color=INK)
    ws3.cell(row=r, column=4, value="должно совпасть с суммой выше").font = Font(size=9, color=GREY)
    for col in range(2, 6):
        ws3.cell(row=r, column=col).fill = TOT

    # ---- итоги по блокам, ниже выборки
    r += 2
    ws3.cell(row=r, column=2, value="Итоги по блокам").font = Font(bold=True, size=11, color=INK)
    r += 1
    _head(ws3, r, ["Блок", "Сумма", "Контрагентов", "Комментарий"], [46, 18, 34, 24])
    r += 1
    tot_first = r
    for blk in BLOCKS:
        ws3.cell(row=r, column=2, value=blk).font = Font(size=10, color=INK)
        c = ws3.cell(row=r, column=3, value=f'=SUMIF({blk_src},B{r},{amt})')
        c.number_format = MONEY
        c.font = Font(size=10, color=INK)
        c = ws3.cell(row=r, column=4, value=f'=COUNTIF({blk_src},B{r})')
        c.font = Font(size=9, color=GREY)
        ws3.cell(row=r, column=5, value=notes[blk]).font = Font(size=9, color=GREY)
        for col in range(2, 6):
            ws3.cell(row=r, column=col).border = Border(bottom=THIN)
        r += 1
    ws3.cell(row=r, column=2, value="ИТОГО ПО СЧЁТУ 60").font = Font(bold=True, size=10, color=INK)
    c = ws3.cell(row=r, column=3, value=f"=SUM(C{tot_first}:C{r - 1})")
    c.number_format = MONEY
    c.font = Font(bold=True, size=10, color=INK)
    for col in range(2, 6):
        ws3.cell(row=r, column=col).fill = TOT
    r += 2
    ws3.cell(row=r, column=2, value="В расходы учредителей").font = Font(bold=True, size=11, color=INK)
    c = ws3.cell(row=r, column=3, value=f"=C{tot_first}+C{tot_first + 1}")
    c.number_format = MONEY
    c.font = Font(bold=True, size=12, color=BRONZE)
    ws3.cell(row=r, column=5,
             value="медицинские + управленческие").font = Font(size=9, color=GREY)

    wb.save(out)
    return out


if __name__ == "__main__":
    canon = json.loads((ROOT / "data/canonical/2026-08.json").read_text(encoding="utf-8"))
    print("готово:", build(canon, ROOT / "out" / "Классификация-расходов-Август-2026.xlsx"))
