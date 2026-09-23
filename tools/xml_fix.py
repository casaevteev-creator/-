# -*- coding: utf-8 -*-
"""Доводка XML-выгрузки «Выгрузка в 1С:Бухгалтерию предприятия» до загрузки.

Берёт файл обмена, выгруженный из БИТ:УМЦ, проверяет его по бухгалтерскому
чек-листу и пишет рядом исправленный файл, готовый к загрузке в БП 3.0.

    python3 tools/xml_fix.py выгрузка.xml
    python3 tools/xml_fix.py выгрузка.xml --out готово.xml
    python3 tools/xml_fix.py выгрузка.xml --проверка     # только протокол, без правки

Что правит:

  склад          пустой «Склад» в отчёте о розничных продажах — подставляется
                 по подразделению из комментария документа;
  статья ДДС     пустая статья движения денежных средств в ПКО и РКО —
                 подставляется по виду операции (у ОРП её ставит сама БП);
  контрагент     пустой контрагент в ПКО «Оплата покупателя»;
  НДС            ставка НДС в строках услуг — для медуслуг всегда «Без НДС»
                 (освобождение по пп. 2 п. 2 ст. 149 НК РФ);
  номер чека     мусорное значение «-1» в реквизите «Номер чека ККМ».

О чём предупреждает, но молча не трогает:

  дата 12:00:00  документ проведён ровно в полдень — почти всегда это дата,
                 проставленная руками, и она разъезжается с датой чека ОФД;
  дубли услуг    одинаковые строки в табличной части «Товары» — след того,
                 что услугу завели несколько сотрудников;
  зачёт аванса   зачёт без чека предоплаты — нарушение 54-ФЗ, правится в кассе,
                 а не в файле.

Настройки — в JSON (см. --config и tools/obmen-nastroyki.json). Названия
справочников должны совпадать с теми, что заведены в БП, символ в символ:
поиск при загрузке идёт по наименованию, и опечатка создаст новый элемент.
"""
import argparse
import json
import sys
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from pathlib import Path

# как 1С пишет ссылку на справочник: поиск по наименованию
DIRECTORY_TYPES = {
    "Склад": ("СправочникСсылка.Склады", "Склады"),
    "СтатьяДвиженияДенежныхСредств": ("СправочникСсылка.СтатьиДвиженияДенежныхСредств",
                                      "СтатьиДвиженияДенежныхСредств"),
    "Контрагент": ("СправочникСсылка.Контрагенты", "Контрагенты"),
}

DEFAULTS = {
    "склад_по_подразделению": {},
    "склад_по_умолчанию": "",
    "статья_ддс": {
        "ОплатаПокупателя": "Розничная выручка",
        "ПолучениеНаличныхВБанке": "Получение наличных в банке",
        "Инкассация": "Внутреннее перемещение денежных средств",
        "ВозвратПокупателю": "Розничная выручка",
    },
    "контрагент_розница": "",
    "ставка_ндс_услуг": "БезНДС",
}


def _tag(el):
    return el.tag.split("}")[-1]


def _value(prop):
    """Значение свойства: текст, «пусто» или краткое описание ссылки."""
    v = prop.find("Значение")
    if v is not None:
        return (v.text or "").strip()
    if prop.find("Пусто") is not None:
        return ""
    ref = prop.find("Ссылка")
    if ref is not None:
        for p in ref:
            if _tag(p) == "Свойство" and p.get("Имя") == "Наименование":
                return (p.findtext("Значение") or "").strip()
        return "<ссылка>"
    return ""


def _find_prop(node, name):
    for p in node:
        if _tag(p) == "Свойство" and p.get("Имя") == name:
            return p
    return None


def _set_directory(prop, name, value):
    """Заменить <Пусто/> на ссылку на элемент справочника, найденный по имени."""
    type_, rule = DIRECTORY_TYPES[name]
    prop.set("Тип", type_)
    prop.set("ИмяПКО", rule)
    for child in list(prop):
        prop.remove(child)
    ref = ET.SubElement(prop, "Ссылка")
    sub = ET.SubElement(ref, "Свойство", {"Имя": "Наименование", "Тип": "Строка"})
    ET.SubElement(sub, "Значение").text = value


def _set_value(prop, value):
    for child in list(prop):
        prop.remove(child)
    if value == "":
        ET.SubElement(prop, "Пусто")
    else:
        ET.SubElement(prop, "Значение").text = value


def nomenclature_map(root):
    """uid номенклатуры -> (наименование, услуга ли это).

    Нужно, чтобы не поставить «Без НДС» товару: освобождение по
    пп. 2 п. 2 ст. 149 НК РФ распространяется на медуслуги, а бельё,
    кремы и перчатки клиника продаёт с налогом на общих основаниях.
    """
    out = {}
    for obj in root.iter():
        if _tag(obj) != "Объект":
            continue
        if "СправочникСсылка.Номенклатура" not in (obj.get("Тип") or ""):
            continue
        ref = next((c for c in obj if _tag(c) == "Ссылка"), None)
        if ref is None:
            continue
        key = name = ""
        for p in ref:
            if _tag(p) != "Свойство":
                continue
            if p.get("Имя") == "{УникальныйИдентификатор}":
                key = _value(p)
            elif p.get("Имя") == "Наименование":
                name = _value(p)
        услуга = ""
        for p in obj:
            if _tag(p) == "Свойство" and p.get("Имя") == "Услуга":
                услуга = _value(p)
        if key:
            out[key] = (name, услуга == "true")
    return out


def _nomenclature_uid(rec):
    prop = _find_prop(rec, "Номенклатура")
    if prop is None:
        return ""
    ref = prop.find("Ссылка")
    if ref is None:
        return ""
    for p in ref:
        if _tag(p) == "Свойство" and p.get("Имя") == "{УникальныйИдентификатор}":
            return _value(p)
    return ""


class Fixer:
    def __init__(self, cfg, nomenclature=None):
        self.cfg = cfg
        self.nom = nomenclature or {}
        self.fixed = []     # что поправили
        self.warned = []    # на что смотреть руками

    def _fix(self, doc, what, was, now):
        self.fixed.append((doc, what, was, now))

    def _warn(self, doc, what, detail):
        self.warned.append((doc, what, detail))

    # --- отдельные проверки -------------------------------------------------

    def склад(self, name, obj, ref, props):
        """Отчёт о розничных продажах без склада в БП не проводится."""
        prop = (_find_prop(ref, "Склад") if ref is not None else None)
        if prop is None:
            prop = _find_prop(obj, "Склад")
        if prop is None or _value(prop):
            return
        подразделение = (props.get("_Комментарий", "") or "").split(".")[0].strip()
        склад = (self.cfg["склад_по_подразделению"].get(подразделение)
                 or self.cfg["склад_по_умолчанию"])
        if not склад:
            self._warn(name, "склад не заполнен",
                       f"подразделение «{подразделение}» — задайте склад в настройках")
            return
        _set_directory(prop, "Склад", склад)
        self._fix(name, "склад", "<пусто>", склад)

    def статья_ддс(self, name, node, вид):
        prop = _find_prop(node, "СтатьяДвиженияДенежныхСредств")
        if prop is None or _value(prop):
            return
        статья = self.cfg["статья_ддс"].get(вид)
        if not статья:
            self._warn(name, "статья ДДС не заполнена",
                       f"вид операции «{вид}» не описан в настройках")
            return
        _set_directory(prop, "СтатьяДвиженияДенежныхСредств", статья)
        self._fix(name, "статья ДДС", "<пусто>", статья)

    def контрагент(self, name, node, вид):
        if вид != "ОплатаПокупателя":
            return
        prop = _find_prop(node, "Контрагент")
        if prop is None or _value(prop):
            return
        кто = self.cfg["контрагент_розница"]
        if not кто:
            self._warn(name, "контрагент не заполнен",
                       "ПКО «Оплата покупателя» без контрагента не проведётся")
            return
        _set_directory(prop, "Контрагент", кто)
        self._fix(name, "контрагент", "<пусто>", кто)

    def ндс(self, name, tab_name, rec):
        """Пустую ставку заполняем, проставленную — не трогаем.

        Медуслуги освобождены от НДС (пп. 2 п. 2 ст. 149 НК РФ), но товары,
        которые клиника продаёт попутно, облагаются на общих основаниях.
        Поэтому строку с явной ставкой переписывать нельзя: это занизит налог.
        """
        prop = _find_prop(rec, "СтавкаНДС")
        if prop is None:
            return
        было = _value(prop)
        надо = self.cfg["ставка_ндс_услуг"]
        if было == надо:
            return
        сумма = ""
        for k in ("Сумма", "СуммаОплаты"):
            q = _find_prop(rec, k)
            if q is not None:
                сумма = _value(q)
                break
        поз, услуга = self.nom.get(_nomenclature_uid(rec), ("", True))
        if было:
            self._warn(name, f"строка с НДС в «{tab_name}»",
                       f"{поз[:34] or '?'} — ставка {было}, сумма {сумма}")
            return
        if not услуга:
            # товар с незаполненной ставкой: ставить «Без НДС» нельзя,
            # это занизило бы налог. Ставку надо завести в карточке товара.
            self._warn(name, "товар без ставки НДС",
                       f"{поз[:34] or '?'} — сумма {сумма}, заполните ставку в карточке")
            return
        _set_value(prop, надо)
        self._fix(name, f"ставка НДС ({tab_name})", "<пусто>", надо)

    def номер_чека(self, name, node):
        prop = _find_prop(node, "НомерЧекаККМ")
        if prop is not None and _value(prop) == "-1":
            _set_value(prop, "")
            self._fix(name, "номер чека ККМ", "-1", "<пусто>")

    def дата_руками(self, name, дата):
        if дата.endswith("T12:00:00"):
            self._warn(name, "дата проставлена вручную",
                       f"{дата} — ровно полдень, сверьте с датой чека ОФД")

    def дубли_услуг(self, name, rows):
        c = Counter((r.get("Номенклатура", ""), r.get("Цена", ""), r.get("Сумма", ""))
                    for r in rows)
        лишние = [(k, n) for k, n in c.items() if n > 1 and k[0]]
        if лишние:
            всего = sum(n - 1 for _, n in лишние)
            self._warn(name, "повторяющиеся строки услуг",
                       f"{len(лишние)} позиц. повторяются, лишних строк {всего}")


def props_of(node):
    out = {}
    for p in node:
        if _tag(p) == "Свойство":
            out[p.get("Имя")] = _value(p)
    return out


def process(path, out_path, cfg, only_check=False):
    tree = ET.parse(path)
    root = tree.getroot()
    f = Fixer(cfg, nomenclature_map(root))
    counts = Counter()

    for obj in root.iter():
        if _tag(obj) != "Объект":
            continue
        тип = (obj.get("Тип") or "").split(".")[-1]
        if not тип or "Документ" not in (obj.get("Тип") or ""):
            continue
        ref = next((c for c in obj if _tag(c) == "Ссылка"), None)
        p = {}
        if ref is not None:
            p.update(props_of(ref))
        p.update(props_of(obj))
        p["_Комментарий"] = p.get("Комментарий", "")
        вид = p.get("ВидОперации", "")
        дата = p.get("Дата", "")
        name = f"{тип} от {дата[:10] or '?'}"
        counts[тип] += 1

        if тип == "ОтчетОРозничныхПродажах":
            f.склад(name, obj, ref, p)
        if тип in ("ПриходныйКассовыйОрдер", "РасходныйКассовыйОрдер"):
            узел = obj if _find_prop(obj, "СтатьяДвиженияДенежныхСредств") is not None else ref
            if узел is not None:
                f.статья_ддс(name, узел, вид)
            узел_к = obj if _find_prop(obj, "Контрагент") is not None else ref
            if узел_к is not None:
                f.контрагент(name, узел_к, вид)
            f.номер_чека(name, obj)
        f.дата_руками(name, дата)

        for tab in obj:
            if _tag(tab) != "ТабличнаяЧасть":
                continue
            имя = tab.get("Имя") or ""
            строки = []
            for rec in tab:
                if имя in ("Товары", "Возвраты", "Предоплаты", "АгентскиеУслуги"):
                    f.ндс(name, имя, rec)
                строки.append(props_of(rec))
            if имя == "Товары":
                f.дубли_услуг(name, строки)

    print(f"Файл: {Path(path).name}")
    print("Документов: " + ", ".join(f"{k} — {v}" for k, v in counts.most_common()))

    print(f"\nИсправлено: {len(f.fixed)}")
    свод = defaultdict(list)
    for doc, what, was, now in f.fixed:
        свод[what].append((doc, was, now))
    for what, items in свод.items():
        print(f"  {what}: {len(items)}")
        for doc, was, now in items[:5]:
            print(f"      {doc}: {was} → {now}")
        if len(items) > 5:
            print(f"      … ещё {len(items) - 5}")

    print(f"\nСмотреть руками: {len(f.warned)}")
    свод_w = defaultdict(list)
    for doc, what, detail in f.warned:
        свод_w[what].append((doc, detail))
    for what, items in свод_w.items():
        print(f"  {what}: {len(items)}")
        for doc, detail in items[:5]:
            print(f"      {doc}: {detail}")
        if len(items) > 5:
            print(f"      … ещё {len(items) - 5}")

    if only_check:
        print("\nРежим проверки: файл не изменён.")
        return 0
    if not f.fixed:
        print("\nПравить нечего — файл можно грузить как есть.")
        return 0
    tree.write(out_path, encoding="UTF-8", xml_declaration=True)
    print(f"\nИсправленный файл: {out_path}")
    return 0


def load_config(path):
    cfg = json.loads(json.dumps(DEFAULTS))
    if path and Path(path).exists():
        user = json.loads(Path(path).read_text(encoding="utf-8"))
        for k, v in user.items():
            if isinstance(v, dict) and isinstance(cfg.get(k), dict):
                cfg[k].update(v)
            else:
                cfg[k] = v
    return cfg


def main():
    ap = argparse.ArgumentParser(description="Доводка XML-выгрузки УМЦ до загрузки в БП")
    ap.add_argument("xml")
    ap.add_argument("--out", help="куда писать исправленный файл")
    ap.add_argument("--config", default="tools/obmen-nastroyki.json")
    ap.add_argument("--проверка", action="store_true", dest="check",
                    help="только протокол, файл не трогать")
    a = ap.parse_args()
    out = a.out or str(Path(a.xml).with_name(Path(a.xml).stem + "-исправленный.xml"))
    return process(a.xml, out, load_config(a.config), a.check)


if __name__ == "__main__":
    sys.exit(main())
