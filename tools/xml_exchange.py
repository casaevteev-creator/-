# -*- coding: utf-8 -*-
"""Разбор XML-выгрузки «Выгрузка в 1С:Бухгалтерию предприятия» и сверка с ОФД.

Что делает:

  * читает файл обмена (формат «Конвертации данных 2.0» — <ФайлОбмена><Объект Тип=...>,
    либо универсальный формат EnterpriseData — <Документ.ОказаниеУслуг> и т.п.);
  * печатает состав выгрузки: сколько документов каждого вида и на какую сумму;
  * разбивает деньги по дням: наличные (ПКО, отчёт о розничных продажах),
    карты (оплата платёжной картой), возвраты;
  * ищет задвоенные документы услуг — один день, один клиент, одна сумма,
    несколько документов (баг управленки: карточку услуги открывает второй
    сотрудник и программа создаёт документ заново);
  * сверяет день в день с отчётом «Такском-Касса» (фискальные данные ОФД).

    python3 tools/xml_exchange.py выгрузка.xml
    python3 tools/xml_exchange.py выгрузка.xml --taksk Такском-сентябрь.xlsx
    python3 tools/xml_exchange.py выгрузка.xml --structure   # что вообще внутри файла

Если файл окажется в незнакомой структуре — `--structure` покажет дерево тегов
и примеры значений, по нему разбор допиливается за минуту.
"""
import argparse
import re
import sys
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from pathlib import Path

DATE_RE = re.compile(r"(\d{4})-(\d{2})-(\d{2})")
DATE_RU_RE = re.compile(r"(\d{2})\.(\d{2})\.(\d{4})")

# поля, в которых 1С держит сумму документа — в порядке доверия
SUM_KEYS = ("СуммаДокумента", "СуммаПлатежа", "Сумма")
# поля с контрагентом/пациентом
PARTY_KEYS = ("Контрагент", "Клиент", "Покупатель", "Пациент")

# как называть виды документов по-человечески
HUMAN = {
    "ОказаниеУслуг": "Оказание услуг",
    "РеализацияТоваровУслуг": "Реализация товаров и услуг",
    "ПриходныйКассовыйОрдер": "ПКО (наличные)",
    "РасходныйКассовыйОрдер": "РКО",
    "ОперацияПоПлатежнойКарте": "Оплата платёжной картой",
    "ОплатаПлатежнойКартой": "Оплата платёжной картой",
    "ОтчетОРозничныхПродажах": "Отчёт о розничных продажах",
    "ВозвратТоваровОтПокупателя": "Возврат от покупателя",
    "ПоступлениеТоваровУслуг": "Поступление товаров и услуг",
    "ТребованиеНакладная": "Требование-накладная",
}
# что считаем наличной выручкой, что картами, что возвратом
CASH_TYPES = ("ПриходныйКассовыйОрдер", "ОтчетОРозничныхПродажах")
CARD_TYPES = ("ОперацияПоПлатежнойКарте", "ОплатаПлатежнойКартой")
REFUND_TYPES = ("ВозвратТоваровОтПокупателя",)
# ПКО с такой операцией — не выручка: деньги пришли не от пациента.
# В сверку с ОФД они не идут, иначе касса «перевыполняет» фискальные данные.
NOT_REVENUE_PKO = ("ПолучениеНаличныхВБанке", "ВозвратОтПоставщика",
                   "ПрочийПриход", "РасчетыПоКредитамИЗаймам",
                   "ВозвратДенежныхСредствПодотчетником")
# РКО — движение денег, а не возврат выручки (инкассация, выплаты)
CASHOUT_TYPES = ("РасходныйКассовыйОрдер",)
# документы услуг — на них смотрим в поиске дублей
SERVICE_TYPES = ("ОказаниеУслуг", "РеализацияТоваровУслуг")


def money(v):
    return f"{v:,.2f}".replace(",", " ").replace(".", ",")


def _num(s):
    if s is None:
        return None
    s = str(s).strip().replace("\xa0", "").replace(" ", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def _day(s):
    """Дата документа -> '2026-09-10'. 1С пишет ISO, отчёты — по-русски."""
    if not s:
        return ""
    m = DATE_RE.search(str(s))
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    m = DATE_RU_RE.search(str(s))
    if m:
        return f"{m.group(3)}-{m.group(2)}-{m.group(1)}"
    return ""


def _short_type(raw):
    """'ДокументСсылка.ОказаниеУслуг' и 'Документ.ОказаниеУслуг' -> 'ОказаниеУслуг'."""
    return str(raw).split(".")[-1] if raw else ""


def _tag(el):
    return el.tag.split("}")[-1]


def _text(el):
    """Всё содержимое узла одной строкой — годится и для ссылок с вложенностью."""
    return " ".join(t.strip() for t in el.itertext() if t and t.strip())


class Doc:
    __slots__ = ("type", "props", "rows", "rule")

    def __init__(self, type_):
        self.type = type_
        self.props = {}
        self.rows = []
        self.rule = ""

    @property
    def day(self):
        return _day(self.props.get("Дата"))

    @property
    def number(self):
        return self.props.get("Номер", "")

    @property
    def party(self):
        for k in PARTY_KEYS:
            v = self.props.get(k)
            if v:
                return v
        return ""

    def tab_total(self, name, field):
        """Итог по колонке табличной части документа."""
        return sum(_num(r.get(field)) or 0 for r in self.rows if r.get("__тч__") == name)

    @property
    def operation(self):
        return self.props.get("ВидОперации", "")

    @property
    def is_revenue(self):
        """Выручка ли это — то, что должно быть пробито по кассе и лежать в ОФД."""
        if self.type == "ПриходныйКассовыйОрдер":
            return self.operation not in NOT_REVENUE_PKO
        return True

    @property
    def total(self):
        """Сумма документа: сначала своё поле, иначе сумма табличной части."""
        for k in SUM_KEYS:
            v = _num(self.props.get(k))
            if v is not None:
                return v
        if self.type == "ОтчетОРозничныхПродажах":
            return (self.tab_total("Товары", "Сумма")
                    + self.tab_total("Предоплаты", "Сумма"))
        s = 0.0
        for row in self.rows:
            for k in ("СуммаПлатежа", "Сумма", "СуммаСНДС"):
                v = _num(row.get(k))
                if v is not None:
                    s += v
                    break
        return s


def _props_kd2(node, doc, into):
    """КД2: <Свойство Имя="Дата"><Значение>...</Значение></Свойство>."""
    for p in node:
        if _tag(p) != "Свойство":
            continue
        name = p.get("Имя") or ""
        val = p.find("Значение")
        into[name] = _text(val) if val is not None else _text(p)


def read_kd2(root):
    """Формат «Конвертации данных 2.0»: плоский список <Объект>."""
    docs = []
    for obj in root.iter():
        if _tag(obj) != "Объект":
            continue
        t = _short_type(obj.get("Тип"))
        if not t:
            continue
        d = Doc(t)
        # ключевые реквизиты (дата, организация, комментарий) 1С кладёт
        # внутрь <Ссылка> — это поля поиска объекта в приёмнике
        for ref in obj:
            if _tag(ref) == "Ссылка":
                _props_kd2(ref, d, d.props)
                break
        _props_kd2(obj, d, d.props)
        d.rule = obj.get("ИмяПравила") or ""
        for tab in obj:
            if _tag(tab) != "ТабличнаяЧасть":
                continue
            for rec in tab:
                row = {"__тч__": tab.get("Имя") or ""}
                _props_kd2(rec, d, row)
                if len(row) > 1:
                    d.rows.append(row)
        docs.append(d)
    return docs


def read_enterprise(root):
    """Универсальный формат: <Документ.ОказаниеУслуг><Дата>...</Дата>."""
    docs = []
    for node in root.iter():
        tag = _tag(node)
        if "." not in tag or not tag.startswith(("Документ", "Справочник")):
            continue
        d = Doc(_short_type(tag))
        for child in node:
            ct = _tag(child)
            kids = list(child)
            if kids and all("." not in _tag(k) for k in kids) and len(kids) > 1:
                # похоже на строку табличной части
                d.rows.append({_tag(k): _text(k) for k in kids})
            else:
                d.props[ct] = _text(child)
        docs.append(d)
    return docs


def read(path):
    root = ET.parse(path).getroot()
    docs = read_kd2(root)
    if not docs:
        docs = read_enterprise(root)
    return root, docs


def structure(path, limit=3):
    """Дерево тегов с примерами значений — чтобы понять незнакомый файл."""
    root = ET.parse(path).getroot()
    seen = Counter()
    samples = defaultdict(list)

    def walk(el, prefix=""):
        p = prefix + "/" + _tag(el)
        seen[p] += 1
        if el.attrib and len(samples[p]) < limit:
            samples[p].append(dict(list(el.attrib.items())[:4]))
        elif el.text and el.text.strip() and len(samples[p]) < limit:
            samples[p].append(el.text.strip()[:40])
        for child in el:
            walk(child, p)

    walk(root)
    print(f"корень: {_tag(root)}, атрибуты: {dict(root.attrib)}\n")
    for p, n in seen.most_common(60):
        ex = "; ".join(str(s) for s in samples[p][:limit])
        print(f"{n:>6}  {p}" + (f"   ← {ex[:90]}" if ex else ""))


def duplicates(docs):
    """Задвоенные услуги: один день + один клиент + одна сумма, разные номера."""
    groups = defaultdict(list)
    for d in docs:
        if d.type not in SERVICE_TYPES:
            continue
        groups[(d.day, d.party, round(d.total, 2))].append(d)
    return {k: v for k, v in groups.items() if len(v) > 1}


def by_day(docs):
    """Раскладка денег по дням так, как их видит бухгалтер.

    cash/card    — живые деньги: наличные в кассу и эквайринг. Их и сверяем с ОФД.
    advance_used — зачёт ранее полученного аванса: выручка есть, денег нет.
    prepay       — полученная предоплата: деньги есть, выручки ещё нет.
    revenue      — реализация (Кт 90.01).
    other/out    — не выручка (получение налички в банке) и инкассация.
    """
    days = defaultdict(lambda: {"cash": 0.0, "card": 0.0, "refund": 0.0,
                                "other": 0.0, "out": 0.0, "advance_used": 0.0,
                                "prepay": 0.0, "revenue": 0.0})
    for d in docs:
        if not d.day:
            continue
        v = days[d.day]
        if d.type == "ОтчетОРозничныхПродажах":
            goods = d.tab_total("Товары", "Сумма")
            goods_back = d.tab_total("Возвраты", "Сумма")
            prepay = d.tab_total("Предоплаты", "Сумма")
            prepay_back = d.tab_total("Предоплаты", "СуммаВозврат")
            card = d.tab_total("Оплата", "СуммаОплаты")
            card_back = d.tab_total("ВозвратОплаты", "СуммаОплаты")
            adv = d.tab_total("ЗачетАвансов", "СуммаЗачета")
            # Наличных в ОРП нет отдельной строкой — это то, что осталось
            # от суммы документа после всех прочих способов её закрыть:
            #   товары − возвраты товаров + предоплата − возврат предоплаты
            #   − эквайринг (за вычетом возвратов по карте) − зачёт аванса
            # Проверено на 01-14.09: сходится с ОФД все четырнадцать дней.
            v["cash"] += (goods - goods_back + prepay - prepay_back
                          - (card - card_back) - adv)
            v["card"] += card - card_back
            # возврат предоплаты и возврат оплаты — одна операция,
            # показанная в двух табличных частях: деньги считаем один раз
            v["refund"] += max(card_back, prepay_back)
            v["advance_used"] += adv
            v["prepay"] += prepay
            v["revenue"] += goods - goods_back
        elif d.type in REFUND_TYPES:
            v["refund"] += abs(d.total)
        elif d.type in CASHOUT_TYPES:
            v["out"] += d.total
        elif d.type in CASH_TYPES:
            v["cash" if d.is_revenue else "other"] += d.total
        elif d.type in CARD_TYPES:
            v["card"] += d.total
        elif d.type in SERVICE_TYPES:
            v["revenue"] += d.total
    return days


def read_taksk(path):
    """Отчёт «Такском-Касса» по фискальным документам -> живые деньги по дням.

    Живые деньги = наличные + безнал. Зачёт ранее внесённого аванса новых
    поступлений не даёт, поэтому в сверку он не идёт.
    """
    import openpyxl

    # без read_only: в коротких выгрузках Такском объявленный диапазон листа
    # врёт, и быстрый режим обрезает таблицу на первых строках
    wb = openpyxl.load_workbook(path, data_only=True)
    for name in wb.sheetnames:
        ws = wb[name]
        rows = list(ws.iter_rows(values_only=True))
        head = None
        for i, row in enumerate(rows[:40]):
            vals = [str(c).strip() if c is not None else "" for c in row]
            if "Наличными" in vals and "Сумма" in vals:
                head = (i, vals)
                break
        if not head:
            continue
        i, h = head
        ci = {k: (h.index(k) if k in h else -1) for k in
              ("Дата и время", "Документ", "Тип операции", "Наличными",
               "Безналичными", "Аванс", "Сумма")}
        t = {"cash": 0.0, "card": 0.0, "adv": 0.0, "refund": 0.0,
             "checks": 0, "byDay": defaultdict(float),
             "days": defaultdict(lambda: {"cash": 0.0, "card": 0.0, "adv": 0.0})}
        for row in rows[i + 1:]:
            if ci["Документ"] < 0 or str(row[ci["Документ"]]).strip() != "Кассовый чек":
                continue
            s = _num(row[ci["Сумма"]]) or 0
            if not s:
                continue
            t["checks"] += 1
            cash = _num(row[ci["Наличными"]]) or 0
            card = _num(row[ci["Безналичными"]]) or 0
            t["cash"] += cash
            t["card"] += card
            if ci["Аванс"] >= 0:
                t["adv"] += _num(row[ci["Аванс"]]) or 0
            op = str(row[ci["Тип операции"]] or "")
            if "озврат" in op:
                t["refund"] += abs(s)
            day = _day(row[ci["Дата и время"]])
            if day:
                t["byDay"][day] += cash + card
                t["days"][day]["cash"] += cash
                t["days"][day]["card"] += card
                t["days"][day]["adv"] += _num(row[ci["Аванс"]]) or 0 if ci["Аванс"] >= 0 else 0
        wb.close()
        return t
    wb.close()
    return None


def report(path, taksk_path=None):
    _, docs = read(path)
    if not docs:
        print("В файле не нашлось объектов. Посмотрите структуру: --structure")
        return 1

    print(f"Файл: {Path(path).name}")
    print(f"Объектов: {len(docs)}\n")

    print("Состав выгрузки")
    print(f"{'документ':<34}{'штук':>7}{'сумма, ₽':>18}")
    tot = Counter()
    cnt = Counter()
    for d in docs:
        tot[d.type] += d.total
        cnt[d.type] += 1
    for t, n in cnt.most_common():
        print(f"{HUMAN.get(t, t):<34}{n:>7}{money(tot[t]):>18}")

    days = by_day(docs)
    if days:
        print("\nДеньги и выручка по дням (из выгрузки)")
        print(f"{'день':<12}{'реализация':>15}{'предоплата':>15}{'зачёт аванса':>15}"
              f"{'наличные':>15}{'карты':>15}")
        agg = defaultdict(float)
        for day in sorted(days):
            v = days[day]
            for k in ("revenue", "prepay", "advance_used", "cash", "card",
                      "refund", "other", "out"):
                agg[k] += v[k]
            print(f"{day:<12}{money(v['revenue']):>15}{money(v['prepay']):>15}"
                  f"{money(v['advance_used']):>15}{money(v['cash']):>15}"
                  f"{money(v['card']):>15}")
        print(f"{'итого':<12}{money(agg['revenue']):>15}{money(agg['prepay']):>15}"
              f"{money(agg['advance_used']):>15}{money(agg['cash']):>15}"
              f"{money(agg['card']):>15}")
        if agg["refund"] or agg["other"] or agg["out"]:
            print(f"\n  возвраты покупателям: {money(agg['refund'])} ₽")
            print(f"  приход не от пациентов (в ОФД не должно быть): "
                  f"{money(agg['other'])} ₽")
            print(f"  инкассация (Дт 57.01 Кт 50.01): {money(agg['out'])} ₽")

    dups = duplicates(docs)
    print("\nЗадвоенные документы услуг")
    if not dups:
        print("не найдено")
    else:
        extra = sum(len(v) - 1 for v in dups.values())
        lost = sum(round(v[0].total, 2) * (len(v) - 1) for v in dups.values())
        print(f"групп: {len(dups)}, лишних документов: {extra}, "
              f"завышение выручки: {money(lost)} ₽\n")
        for (day, party, amount), v in sorted(dups.items())[:20]:
            nums = ", ".join(d.number or "б/н" for d in v)
            print(f"  {day}  {party[:32]:<32} {money(amount):>14}  ×{len(v)}  [{nums}]")
        if len(dups) > 20:
            print(f"  … ещё {len(dups) - 20} групп")

    if taksk_path:
        t = read_taksk(taksk_path)
        if not t:
            print(f"\nОтчёт Такском не распознан: {taksk_path}")
            return 1
        print(f"\nСверка с ОФД (Такском): чеков {t['checks']}")
        print(f"{'день':<12}{'':<14}{'выгрузка':>15}{'ОФД':>15}{'разница':>13}")
        all_days = sorted(set(days) | set(t["days"]))
        bad = 0
        for day in all_days:
            ours = days.get(day, {})
            th = t["days"].get(day, {"cash": 0.0, "card": 0.0, "adv": 0.0})
            for key, label in (("cash", "наличные"), ("card", "карты"),
                               ("advance_used", "зачёт аванса")):
                o = ours.get(key, 0.0)
                x = th["adv" if key == "advance_used" else key]
                diff = round(o - x, 2)
                mark = "" if abs(diff) < 1 else "  ←"
                if abs(diff) >= 1:
                    bad += 1
                print(f"{day if label == 'наличные' else '':<12}{label:<14}"
                      f"{money(o):>15}{money(x):>15}{money(diff):>13}{mark}")
        return 1 if bad else 0
    return 0


def main():
    ap = argparse.ArgumentParser(description="Разбор XML-выгрузки УМЦ -> 1С:Бухгалтерия")
    ap.add_argument("xml", help="файл обмена, выгруженный из управленки")
    ap.add_argument("--taksk", help="отчёт Такском-Касса (xlsx) за тот же период")
    ap.add_argument("--structure", action="store_true",
                    help="показать дерево тегов файла вместо разбора")
    a = ap.parse_args()
    if a.structure:
        structure(a.xml)
        return 0
    return report(a.xml, a.taksk)


if __name__ == "__main__":
    sys.exit(main())
