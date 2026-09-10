# -*- coding: utf-8 -*-
"""Классификация расходов по контрагентам — источник только файл бухгалтерии.

Читает лист «Расходы по контрагентам» присланного файла и собирает книгу
с живыми формулами: контрагенты → статьи → свод, плюс выборка по статье.
"""
import json
import sys
import tempfile
import zipfile
from pathlib import Path

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent.parent
MONEY = "#,##0.00;-#,##0.00;—"
INK, BRONZE, GREY, RED = "1A1613", "8A6A3C", "807A70", "9B2C2C"
HEAD = PatternFill("solid", fgColor="EFEAE2")
TOT = PatternFill("solid", fgColor="F6F1E8")
PICKF = PatternFill("solid", fgColor="FFF6E6")
THIN = Side(style="thin", color="D9D2C6")
BLOCKS = ["медицинские", "управленческие"]

# статьи, которых нет в списке присланного файла — назначены по смыслу
EXTRA = {
    "Мед.инструмент": "медицинские",
    "Мед.мебель": "медицинские",
    "Медоборудование 320588 и расходники 194940": "медицинские",
    "Стиральный порошок": "медицинские",
    "Обслуживание 1СПредприятие": "управленческие",
    "Обслуживание ККТ": "управленческие",
    "Ремонт керамогранит": "управленческие",
}


def read_source(path):
    """Файл 1С не открывается openpyxl напрямую: у него SharedStrings с заглавной S."""
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except KeyError:
        z = zipfile.ZipFile(path)
        tmp = Path(tempfile.mkdtemp()) / "f.xlsx"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
            for it in z.infolist():
                out.writestr(it.filename.replace("SharedStrings", "sharedStrings"),
                             z.read(it.filename))
        wb = openpyxl.load_workbook(tmp, data_only=True)
    ws = wb["Расходы по контрагентам"]
    rows, credit = [], None
    for r in range(5, ws.max_row + 1):
        name = ws.cell(row=r, column=2).value
        amount = ws.cell(row=r, column=3).value
        art = str(ws.cell(row=r, column=4).value or "").strip()
        if not name or not isinstance(amount, (int, float)):
            continue
        name = str(name).strip()
        if name.upper().startswith("ИТОГО"):
            continue
        if "кредитной линии" in art.lower():
            credit = (name, amount, art)
        else:
            rows.append((name, amount, art))
    blocks = dict(EXTRA)
    ws2 = wb["Статьи расходов"]
    for r in range(6, ws2.max_row + 1):
        a, b = ws2.cell(row=r, column=2).value, ws2.cell(row=r, column=5).value
        if a and b:
            blocks[str(a).strip()] = str(b).strip()
    return rows, credit, blocks


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
    ws.sheet_view.showGridLines = False


def build(rows, credit, blocks, adjust, out):
    wb = Workbook()

    # ---------------------------------------------------------------- контрагенты
    ws = wb.active
    ws.title = "Контрагенты"
    _title(ws, "Расходы по контрагентам, счёт 60",
           "Данные из файла бухгалтерии. Меняете статью или блок — суммы на листах "
           "«Статьи» и «Свод» пересчитываются сами.")
    _head(ws, 5, ["Контрагент", "Оплачено", "Статья", "Блок"], [46, 18, 40, 20])
    first = 6
    for i, (name, amount, art) in enumerate(rows):
        r = first + i
        ws.cell(row=r, column=2, value=name).font = Font(size=10, color=INK)
        c = ws.cell(row=r, column=3, value=amount)
        c.number_format = MONEY
        c.font = Font(size=10, color=INK)
        ws.cell(row=r, column=4, value=art).font = Font(size=10, color=INK)
        blk = blocks.get(art, BLOCKS[1])
        c = ws.cell(row=r, column=5, value=blk)
        c.font = Font(size=10, color=BRONZE if art in blocks else RED)
        for col in range(2, 6):
            ws.cell(row=r, column=col).border = Border(bottom=THIN)
    last = first + len(rows) - 1
    ws.freeze_panes = "B6"

    r = last + 1
    ws.cell(row=r, column=2, value="Сумма по контрагентам").font = Font(bold=True, size=10, color=INK)
    c = ws.cell(row=r, column=3, value=f"=SUM(C{first}:C{last})")
    c.number_format = MONEY
    c.font = Font(bold=True, size=10, color=INK)
    adj_first = r + 1
    for label, value in adjust:
        r += 1
        ws.cell(row=r, column=2, value=label).font = Font(size=10, color=RED)
        c = ws.cell(row=r, column=3, value=value)
        c.number_format = MONEY
        c.font = Font(size=10, color=RED)
        ws.cell(row=r, column=4, value="ручная поправка бухгалтерии").font = Font(size=9, color=GREY)
    r += 1
    total_row = r
    ws.cell(row=r, column=2, value="ИТОГО ПО СЧЁТУ 60").font = Font(bold=True, size=11, color=INK)
    c = ws.cell(row=r, column=3, value=f"=C{last + 1}+SUM(C{adj_first}:C{r - 1})")
    c.number_format = MONEY
    c.font = Font(bold=True, size=11, color=INK)
    for col in range(2, 6):
        ws.cell(row=r, column=col).fill = TOT
    if credit:
        r += 2
        ws.cell(row=r, column=2, value=credit[0]).font = Font(size=10, color=INK)
        c = ws.cell(row=r, column=3, value=credit[1])
        c.number_format = MONEY
        c.font = Font(size=10, color=INK)
        ws.cell(row=r, column=4, value=credit[2]).font = Font(size=10, color=BRONZE)
        ws.cell(row=r, column=5, value="в расходы учредителей не входит").font = Font(size=9, color=GREY)

    dv_b = DataValidation(type="list", formula1='"' + ",".join(BLOCKS) + '"', showDropDown=False)
    ws.add_data_validation(dv_b)
    dv_b.add(f"E{first}:E{last}")

    # счётчик выборки
    c = ws.cell(row=4, column=7, value="№ в выборке →")
    c.font = Font(size=9, color=GREY)
    c.alignment = Alignment(horizontal="center")
    ws.column_dimensions["G"].width = 13
    ws.cell(row=first - 1, column=7, value=0).font = Font(size=9, color="EFEAE2")
    for rr in range(first, last + 1):
        c = ws.cell(row=rr, column=7,
                    value=f'=IF(OR($D{rr}=Свод!$C$5,$E{rr}=Свод!$C$5),MAX($G${first - 1}:$G{rr - 1})+1,"")')
        c.font = Font(size=9, color=GREY)
        c.alignment = Alignment(horizontal="center")

    amt = f"Контрагенты!$C${first}:$C${last}"
    art_src = f"Контрагенты!$D${first}:$D${last}"
    blk_src = f"Контрагенты!$E${first}:$E${last}"
    arts = sorted({a for _, _, a in rows})

    # ---------------------------------------------------------------- списки
    wsl = wb.create_sheet("Списки")
    picks = BLOCKS + arts
    for i, v in enumerate(picks, start=1):
        wsl.cell(row=i, column=1, value=v)
    wsl.column_dimensions["A"].width = 44
    wsl.sheet_state = "hidden"

    # ---------------------------------------------------------------- статьи
    ws2 = wb.create_sheet("Статьи")
    _title(ws2, "Статьи расходов",
           "Сумма статьи — SUMIF по листу «Контрагенты». База для доли переключается в C4.")
    ws2.cell(row=4, column=2, value="Доля считается от суммы по контрагентам:").font = Font(size=10, color=GREY)
    c = ws2.cell(row=4, column=3, value=f"=Контрагенты!$C${last + 1}")
    c.number_format = MONEY
    c.font = Font(bold=True, size=10, color=BRONZE)
    ws2.cell(row=4, column=4,
             value="ручные поправки в базу не входят — они не разнесены по статьям").font = \
        Font(size=9, color=GREY)
    _head(ws2, 6, ["Статья", "Блок", "Сумма", "Доля", "Контрагентов"], [40, 20, 18, 10, 14])
    r2 = 7
    tot_rows = []
    for blk in BLOCKS:
        items = [a for a in arts if blocks.get(a, BLOCKS[1]) == blk]
        if not items:
            continue
        c = ws2.cell(row=r2, column=2, value=blk.upper())
        c.font = Font(bold=True, size=10, color=BRONZE)
        r2 += 1
        start = r2
        for art in items:
            c = ws2.cell(row=r2, column=2, value=art)
            c.font = Font(size=10, color=INK)
            c.alignment = Alignment(indent=1)
            ws2.cell(row=r2, column=3, value=blk).font = Font(size=9, color=GREY)
            c = ws2.cell(row=r2, column=4, value=f'=SUMIF({art_src},B{r2}&"",{amt})')
            c.number_format = MONEY
            c.font = Font(size=10, color=INK)
            c = ws2.cell(row=r2, column=5, value=f"=IF($C$4=0,0,D{r2}/$C$4)")
            c.number_format = "0.0%"
            c.font = Font(size=9, color=GREY)
            c = ws2.cell(row=r2, column=6, value=f'=COUNTIF({art_src},B{r2}&"")')
            c.font = Font(size=9, color=GREY)
            for col in range(2, 7):
                ws2.cell(row=r2, column=col).border = Border(bottom=THIN)
            r2 += 1
        c = ws2.cell(row=r2, column=2, value=f"Итого «{blk}»")
        c.font = Font(bold=True, size=10, color=INK)
        c = ws2.cell(row=r2, column=4, value=f"=SUM(D{start}:D{r2 - 1})")
        c.number_format = MONEY
        c.font = Font(bold=True, size=10, color=INK)
        c = ws2.cell(row=r2, column=5, value=f"=IF($C$4=0,0,D{r2}/$C$4)")
        c.number_format = "0.0%"
        c.font = Font(bold=True, size=9, color=BRONZE)
        for col in range(2, 7):
            ws2.cell(row=r2, column=col).fill = TOT
        tot_rows.append(r2)
        r2 += 2
    ws2.cell(row=r2, column=2, value="ВСЕГО ПО СТАТЬЯМ").font = Font(bold=True, size=11, color=INK)
    c = ws2.cell(row=r2, column=4, value="+".join(f"D{x}" for x in tot_rows))
    c.value = "=" + "+".join(f"D{x}" for x in tot_rows)
    c.number_format = MONEY
    c.font = Font(bold=True, size=11, color=INK)
    c = ws2.cell(row=r2, column=5, value=f"=IF($C$4=0,0,D{r2}/$C$4)")
    c.number_format = "0.0%"
    c.font = Font(bold=True, size=11, color=BRONZE)
    ws2.cell(row=r2, column=6, value="сходится с суммой по контрагентам, доля 100,0 %").font = Font(size=9, color=GREY)
    ws2.freeze_panes = "B7"

    # ---------------------------------------------------------------- свод
    ws3 = wb.create_sheet("Свод")
    _title(ws3, "Свод и выборка",
           "Выберите блок или статью в списке — ниже развернётся состав.")
    ws3.cell(row=5, column=2, value="Блок или статья:").font = Font(bold=True, size=11, color=INK)
    pick = ws3.cell(row=5, column=3, value=BLOCKS[0])
    pick.font = Font(bold=True, size=12, color=BRONZE)
    pick.fill = PICKF
    pick.border = Border(bottom=Side(style="medium", color=BRONZE))
    dv = DataValidation(type="list", formula1=f"=Списки!$A$1:$A${len(picks)}", showDropDown=False)
    ws3.add_data_validation(dv)
    dv.add("C5")
    ws3.cell(row=6, column=2, value="Сумма:").font = Font(size=10, color=GREY)
    c = ws3.cell(row=6, column=3, value=f'=SUMIF({blk_src},$C$5,{amt})+SUMIF({art_src},$C$5,{amt})')
    c.number_format = MONEY
    c.font = Font(bold=True, size=12, color=INK)
    ws3.cell(row=6, column=4, value="Контрагентов:").font = Font(size=10, color=GREY)
    c = ws3.cell(row=6, column=5, value=f'=COUNTIF({blk_src},$C$5)+COUNTIF({art_src},$C$5)')
    c.font = Font(bold=True, size=12, color=INK)
    _head(ws3, 8, ["Контрагент", "Оплачено", "Статья", "Блок"], [46, 18, 40, 20])
    key = f"Контрагенты!$G${first}:$G${last}"
    n_rows = len(rows) + 5
    for i in range(n_rows):
        rr = 9 + i
        for col, sc in ((2, "B"), (3, "C"), (4, "D"), (5, "E")):
            c = ws3.cell(row=rr, column=col, value=(
                f'=IFERROR(INDEX(Контрагенты!${sc}${first}:${sc}${last},'
                f'MATCH({i + 1},{key},0)),"")'))
            c.font = Font(size=10, color=INK if col < 4 else GREY)
            c.border = Border(bottom=THIN)
            if col == 3:
                c.number_format = MONEY
    rr = 9 + n_rows
    ws3.cell(row=rr, column=2, value="ИТОГО ПО ВЫБОРКЕ").font = Font(bold=True, size=10, color=INK)
    c = ws3.cell(row=rr, column=3, value=f"=SUM(C9:C{rr - 1})")
    c.number_format = MONEY
    c.font = Font(bold=True, size=10, color=INK)
    ws3.cell(row=rr, column=4, value="сходится с суммой выше").font = Font(size=9, color=GREY)
    for col in range(2, 6):
        ws3.cell(row=rr, column=col).fill = TOT

    rr += 2
    ws3.cell(row=rr, column=2, value="Итоги").font = Font(bold=True, size=11, color=INK)
    rr += 1
    _head(ws3, rr, ["Показатель", "Сумма", "Контрагентов", "Комментарий"], [46, 18, 40, 20])
    rr += 1
    base = rr
    for blk in BLOCKS:
        ws3.cell(row=rr, column=2, value=blk).font = Font(size=10, color=INK)
        c = ws3.cell(row=rr, column=3, value=f'=SUMIF({blk_src},B{rr},{amt})')
        c.number_format = MONEY
        c.font = Font(size=10, color=INK)
        c = ws3.cell(row=rr, column=4, value=f'=COUNTIF({blk_src},B{rr})')
        c.font = Font(size=9, color=GREY)
        for col in range(2, 6):
            ws3.cell(row=rr, column=col).border = Border(bottom=THIN)
        rr += 1
    ws3.cell(row=rr, column=2, value="Сумма по контрагентам").font = Font(bold=True, size=10, color=INK)
    c = ws3.cell(row=rr, column=3, value=f"=SUM(C{base}:C{rr - 1})")
    c.number_format = MONEY
    c.font = Font(bold=True, size=10, color=INK)
    for col in range(2, 6):
        ws3.cell(row=rr, column=col).fill = TOT
    for label, value in adjust:
        rr += 1
        ws3.cell(row=rr, column=2, value=label).font = Font(size=10, color=RED)
        c = ws3.cell(row=rr, column=3, value=f"=Контрагенты!C{adj_first + [a[0] for a in adjust].index(label)}")
        c.number_format = MONEY
        c.font = Font(size=10, color=RED)
        ws3.cell(row=rr, column=5, value="ручная поправка бухгалтерии").font = Font(size=9, color=GREY)
    rr += 1
    founders = rr
    ws3.cell(row=rr, column=2, value="В РАСХОДЫ УЧРЕДИТЕЛЕЙ").font = Font(bold=True, size=12, color=INK)
    c = ws3.cell(row=rr, column=3, value=f"=Контрагенты!C{total_row}")
    c.number_format = MONEY
    c.font = Font(bold=True, size=12, color=BRONZE)
    for col in range(2, 6):
        ws3.cell(row=rr, column=col).fill = TOT
    if credit:
        rr += 2
        ws3.cell(row=rr, column=2, value=credit[0]).font = Font(size=10, color=GREY)
        c = ws3.cell(row=rr, column=3, value=credit[1])
        c.number_format = MONEY
        c.font = Font(size=10, color=GREY)
        ws3.cell(row=rr, column=5, value="за суммой, в расходы не входит").font = Font(size=9, color=GREY)
    ws3.freeze_panes = "B9"

    wb.move_sheet("Свод", offset=-1)
    wb.save(out)
    return out


if __name__ == "__main__":
    src = sys.argv[1]
    rows, credit, blocks = read_source(src)
    adjust = [("Корректировка 1", -169644.37), ("Корректировка 2", -6000.0),
              ("Корректировка 3", -44711.0)]
    out = build(rows, credit, blocks, adjust,
                ROOT / "out" / "Расходы-по-контрагентам-Август-2026.xlsx")
    print(f"контрагентов {len(rows)}, сумма {sum(a for _, a, _ in rows):,.2f}")
    print("готово:", out)
