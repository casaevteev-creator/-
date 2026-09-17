/* Проверка приложения на эталоне: 01–14.09.2026.
 *
 * Открывает собранный out/obmen.html настоящим браузером, кладёт в него четыре
 * файла из data/source/2026-09 и сверяет то, что приложение показало, с
 * цифрами, которые тот же период даёт tools/xml_exchange.py и tools/xml_fix.py.
 *
 *     npm install xlsx@0.18.5 playwright@1.49.1 --no-save
 *     python3 tools/obmen_app/build.py
 *     node tools/obmen_app/test.mjs
 */
import pkg from "playwright";
import { fileURLToPath } from "url";
import { dirname, join } from "path";
import { existsSync, copyFileSync, mkdtempSync } from "fs";
import { tmpdir } from "os";

const { chromium } = pkg;
const ROOT = join(dirname(fileURLToPath(import.meta.url)), "..", "..");
const SRC = join(ROOT, "data", "source", "2026-09");
const APP = join(ROOT, "out", "obmen.html");

/* эталон: то же, что печатают питоновские инструменты на этом периоде */
const ЖДЁМ = {
  "ВЫРУЧКА": "43 549 220",
  "ПРЕДОПЛАТА": "13 429 900",
  "ЗАЧЁТ АВАНСА": "11 413 940",
  "НАЛИЧНЫЕ": "22 221 300",
  "ЭКВАЙРИНГ": "21 755 880",
  "ИНКАССАЦИЯ": "25 468 358",
  "ГРУПП-КАНДИДАТОВ": "5",
  "УЖЕ ПОГАШЕНО": "0",
  "НЕ ПРОВЕРЕНО": "3",
  "ЗАДВОЕНО ПО ДЕНЬГАМ": "4 500"
};
const ПРАВОК = 68;
/* эквайринг по терминалам: сентябрь сходится, кроме 8 000 ₽ на облачной кассе */
const ТЕРМИНАЛЫ = [
  ["Корпус 1 / Терминал", "20 512 380", "20 512 380"],
  ["Онлайн-касса Карта", "1 074 500", "1 066 500"],
  ["Корпус 2 / Терминал ВТБ", "169 000", "169 000"]
];

/* браузер не принимает файлы с кириллицей в имени — кладём копии с латинскими */
const TMP = mkdtempSync(join(tmpdir(), "obmen-"));
const ascii = (имя, как) => {
  const to = join(TMP, как);
  copyFileSync(join(SRC, имя), to);
  return to;
};
const ФАЙЛЫ = [
  ascii("Выгрузка-в-БП-01-14.09.2026.xml", "obmen.xml"),
  ascii("Такском-фискальные-документы-01-14.09.2026.xlsx", "ofd.xlsx"),
  ascii("Услуги-реестр-01-14.09.2026.xlsx", "uslugi.xlsx"),
  ascii("Оплаты-детально-по-документам-01-14.09.2026.xlsx", "oplaty.xlsx")
];

const BIN = ["/opt/pw-browsers/chromium-1194/chrome-linux/chrome",
             "/opt/pw-browsers/chromium/chrome-linux/chrome"].find(existsSync);

const browser = await chromium.launch(BIN ? { executablePath: BIN } : {});
const page = await browser.newPage();
const errors = [];
page.on("pageerror", e => errors.push("ошибка страницы: " + e.message));
page.on("console", m => { if (m.type() === "error") errors.push("консоль: " + m.text()); });

await page.goto("file://" + APP);
const inputs = await page.locator(".slot input[type=file]").all();
if (inputs.length !== 4) throw new Error("ячеек для файлов не четыре, а " + inputs.length);
for (let i = 0; i < 4; i++) await inputs[i].setInputFiles(ФАЙЛЫ[i]);
for (let i = 0; i < 4; i++)
  if (!await inputs[i].evaluate(e => e.files.length))
    throw new Error("браузер не принял файл " + ФАЙЛЫ[i]);
await page.click("#run");
await page.waitForSelector("#out section", { timeout: 60000 });

const беда = [];
const наЭкране = await page.locator("#err").innerText();
if (наЭкране.trim()) беда.push("приложение показало ошибку: " + наЭкране.trim());

/* toLocaleString разделяет разряды неразрывным пробелом — сравниваем по обычному */
const ровно = s => s.replace(/[   ]/g, " ").trim();
const карточки = {};
for (const t of await page.locator("#out .card").allInnerTexts()){
  const [l, v] = t.split("\n");
  if (карточки[l] === undefined) карточки[l] = ровно(v);   // первая карточка с этим названием
}
for (const [l, v] of Object.entries(ЖДЁМ)){
  if (карточки[l] !== v) беда.push(`${l}: показано «${карточки[l]}», ждали «${v}»`);
}

const всё = ровно(await page.locator("#out").innerText());
if (!всё.includes("исправлено " + ПРАВОК)) беда.push("правок не " + ПРАВОК);
/* по умолчанию перемещения оформляются через 57.01: приходному ордеру меняется
   вид операции на «Прочий приход», и обработчик правила подставит ему счёт */
if (!всё.includes("Перемещения между кассами: 6 пар на 40 500 ₽ — оформлены через 57.01"))
  беда.push("перемещения между кассами не оформлены через 57.01");
const видыОпераций = await page.evaluate(() => {
  const x = new DOMParser().parseFromString(FIXED_TEXT, "application/xml");
  let прочий = 0;
  for (const o of x.getElementsByTagName("Объект")){
    if (o.getAttribute("Тип") !== "ДокументСсылка.ПриходныйКассовыйОрдер") continue;
    for (const s of o.children)
      if (s.getAttribute && s.getAttribute("Имя") === "ВидОперации" &&
          s.textContent.trim() === "ПрочийПриход") { прочий++; break; }
  }
  return прочий;
});
if (видыОпераций !== 6)
  беда.push("приходных ордеров с «Прочий приход» " + видыОпераций + ", ждали 6");
if (!всё.includes("Наличные сошлись полностью")) беда.push("наличные не сошлись с ОФД");
/* наличные должны сходиться не только в сумме, но и по каждой кассе: иначе
   выручку записали не на тот аппарат, а общий итог этого не покажет */
if (!всё.includes("Каждая касса сходится со своим аппаратом"))
  беда.push("наличные по кассам не сошлись с ОФД");
for (const [подр, сумма] of [["Форма строение 12", "22 180 800"], ["Форма строение 1", "40 500"]]){
  const строка = (await page.locator("#out tr").allInnerTexts()).map(ровно)
    /* «Форма строение 1» — префикс «Форма строение 12», сверяем с разделителем */
    .find(x => x.startsWith(подр + "\t"));
  if (!строка || !строка.includes(сумма))
    беда.push("касса «" + подр + "»: ждали " + сумма + ", видно «" + строка + "»");
}
if (!всё.includes("Наличные сходятся с фискальными данными до рубля"))
  беда.push("в итоге наверху нет строки про сошедшиеся наличные");
if (!всё.includes("Задвоено по деньгам 4 500 ₽ — 2 группы"))
  беда.push("в итоге нет разбора задвоения или сломано склонение");
if (!всё.includes("Как загрузить файл в бухгалтерию")) беда.push("нет памятки по загрузке");

/* разрез эквайринга по терминалам — то, чем ловится «не тот терминал» */
const строки = await page.locator("#out tr").allInnerTexts();
for (const [имя, наш, ofd] of ТЕРМИНАЛЫ){
  const r = строки.map(ровно).find(x => x.startsWith(имя));
  if (!r) { беда.push("нет строки терминала «" + имя + "»"); continue; }
  if (!r.includes(наш) || !r.includes(ofd))
    беда.push("терминал «" + имя + "»: ждали " + наш + " против " + ofd + ", показано «" + r + "»");
}

/* исправленный файл должен скачиваться и оставаться валидным XML */
const ждём = page.waitForEvent("download", { timeout: 15000 });
await page.click("#save");
const dl = await ждём;
const path = await dl.path();
const text = await (await import("fs/promises")).readFile(path, "utf8");
if ((text.match(/<\?xml/g) || []).length !== 1) беда.push("в файле не одно объявление XML");
if (!text.includes("<ФайлОбмена")) беда.push("в скачанном файле нет корневого узла");
if ((text.match(/<Правило>/g) || []).length !== 48) беда.push("правила обмена повреждены");

/* Второй прогон: август. Там приложение не просто дозаполняет реквизиты,
   а двигает деньги между терминалами — самая рискованная правка, поэтому
   проверяем её на настоящих данных и следим, чтобы итог не поехал. */
const АВГ = join(ROOT, "data", "source", "2026-08");
const авг = (имя, как) => { const to = join(TMP, как); copyFileSync(join(АВГ, имя), to); return to; };
const стр2 = await browser.newPage();
стр2.on("pageerror", e => errors.push("август, ошибка страницы: " + e.message));
await стр2.goto("file://" + APP);
const in2 = await стр2.locator(".slot input[type=file]").all();
await in2[0].setInputFiles(авг("Выгрузка-в-БП-11-31.08.2026.xml", "a.xml"));
await in2[1].setInputFiles(авг("Такском-фискальные-документы-11-31.08.2026.xlsx", "a-ofd.xlsx"));
await стр2.click("#run");
await стр2.waitForSelector("#out section", { timeout: 60000 });

const текст2 = ровно(await стр2.locator("#out").innerText());
if (!текст2.includes("1 548 500 ₽ с «Онлайн-касса Карта» на «Корпус 1 / Терминал»"))
  беда.push("август: перенос эквайринга 13.08 не выполнен");
if (!текст2.includes("В самой УМЦ это не исправлено"))
  беда.push("август: нет предупреждения, что в программе ошибка осталась");

/* считаем деньги прямо в исправленном файле: итог не должен измениться ни на рубль */
const свод = await стр2.evaluate(() => {
  const x = new DOMParser().parseFromString(FIXED_TEXT, "application/xml");
  const имена = {};
  for (const o of x.getElementsByTagName("Объект"))
    if (o.getAttribute("Тип") === "СправочникСсылка.ВидыОплатОрганизаций")
      for (const s of o.children)
        if (s.getAttribute && s.getAttribute("Имя") === "Наименование")
          имена[o.getAttribute("Нпп")] = s.textContent.trim();
  let итог = 0; const день13 = {};
  for (const o of x.getElementsByTagName("Объект")){
    if (o.getAttribute("Тип") !== "ДокументСсылка.ОтчетОРозничныхПродажах") continue;
    const ссыл = o.querySelector("Ссылка");
    let дата = "";
    for (const s of (ссыл || o).children)
      if (s.getAttribute("Имя") === "Дата") дата = s.textContent.trim().slice(0, 10);
    for (const t of o.getElementsByTagName("ТабличнаяЧасть")){
      const знак = t.getAttribute("Имя") === "Оплата" ? 1
                 : t.getAttribute("Имя") === "ВозвратОплаты" ? -1 : 0;
      if (!знак) continue;
      for (const rec of t.children){
        let npp = "", сум = 0;
        for (const s of rec.children){
          if (s.getAttribute("Имя") === "ВидОплаты"){
            const r = s.querySelector("Ссылка"); npp = r ? r.getAttribute("Нпп") : "";
          }
          if (s.getAttribute("Имя") === "СуммаОплаты") сум = parseFloat(s.textContent) || 0;
        }
        итог += знак * сум;
        if (дата === "2026-08-13" && знак > 0)
          день13[имена[npp] || "?"] = (день13[имена[npp] || "?"] || 0) + сум;
      }
    }
  }
  return {итог, день13};
});
/* 39 495 440 в исходнике плюс 20 000: две строки возврата зачтённого аванса
   ушли из «Возврата оплаты», где им не место, — эквайринг на них не уменьшается */
if (Math.round(свод.итог) !== 39515440)
  беда.push("август: итог эквайринга поехал — стало " + свод.итог + ", ждали 39515440");
if (Math.round(свод.день13["Онлайн-касса Карта"] || 0) !== 55000)
  беда.push("август 13.08: облачная касса " + свод.день13["Онлайн-касса Карта"] + ", ждали 55000 (столько в ОФД)");
if (Math.round(свод.день13["Корпус 1 / Терминал"] || 0) !== 5381300)
  беда.push("август 13.08: Штрих " + свод.день13["Корпус 1 / Терминал"] + ", ждали 5381300");

/* Третий прогон: август с дописыванием кассы из ОФД. Правка добавляет в файл
   деньги, поэтому следим за главным — наличные не должны сдвинуться. */
const стр3 = await browser.newPage();
стр3.on("pageerror", e => errors.push("дописывание, ошибка страницы: " + e.message));
await стр3.goto("file://" + APP);
await стр3.locator("details.set summary").first().click();
await стр3.check("#c-dopisat");
const in3 = await стр3.locator(".slot input[type=file]").all();
await in3[0].setInputFiles(авг("Выгрузка-в-БП-11-31.08.2026.xml", "b.xml"));
await in3[1].setInputFiles(авг("Такском-фискальные-документы-11-31.08.2026.xlsx", "b-ofd.xlsx"));
await стр3.click("#run");
await стр3.waitForSelector("#out section", { timeout: 60000 });

const текст3 = ровно(await стр3.locator("#out").innerText());
if (!текст3.includes("Дописано из ОФД 2 073 500 ₽ авансов"))
  беда.push("дописывание: не добавлены авансы облачной кассы");
if (!текст3.includes("Наличные сходятся с фискальными данными до рубля"))
  беда.push("дописывание: наличные разъехались с ОФД — правка тронула не то");
if (!текст3.includes("Проведено из ОФД 350 000 ₽ возвратов"))
  беда.push("дописывание: не проведён возврат 24.08, пробитый кассой");
if (!текст3.includes("Возврат зачёта аванса на 20 000 ₽ убран"))
  беда.push("не разобраны строки возврата без терминала (возврат зачёта аванса)");
if (!текст3.includes("Терминал проставлен по данным ОФД на 2 000 ₽"))
  беда.push("не проставлен терминал в строке оплаты 11.08");
if (!текст3.includes("Файл можно загружать"))
  беда.push("август после всех правок всё ещё не готов к загрузке");

/* ни одной строки оплаты без терминала и ни одного задвоенного «ВидОплаты»:
   с такой строкой документ в бухгалтерии не проведётся */
const строкиОплаты = await стр3.evaluate(() => {
  const x = new DOMParser().parseFromString(FIXED_TEXT, "application/xml");
  let сирот = 0, дубль = 0;
  for (const o of x.getElementsByTagName("Объект")){
    if (o.getAttribute("Тип") !== "ДокументСсылка.ОтчетОРозничныхПродажах") continue;
    for (const t of o.getElementsByTagName("ТабличнаяЧасть")){
      const имя = t.getAttribute("Имя");
      if (имя !== "Оплата" && имя !== "ВозвратОплаты") continue;
      for (const rec of t.children){
        const vo = [...rec.children].filter(s => s.getAttribute("Имя") === "ВидОплаты");
        if (vo.length > 1) дубль++;
        const r = vo[0] && vo[0].querySelector("Ссылка");
        if (!r || !r.getAttribute("Нпп")) сирот++;
      }
    }
  }
  return {сирот, дубль};
});
if (строкиОплаты.сирот) беда.push("строк оплаты без терминала: " + строкиОплаты.сирот);
if (строкиОплаты.дубль) беда.push("строк с двумя «ВидОплаты»: " + строкиОплаты.дубль);

const баланс = await стр3.evaluate(() => {
  const x = new DOMParser().parseFromString(FIXED_TEXT, "application/xml");
  const итог = {товары:0, предоплата:0, зачёт:0, карты:0, наличные:0};
  const сум = (o, таб, поле) => {
    let s = 0;
    for (const t of o.getElementsByTagName("ТабличнаяЧасть")){
      if (t.getAttribute("Имя") !== таб) continue;
      for (const rec of t.children)
        for (const p of rec.children)
          if (p.getAttribute("Имя") === поле) s += parseFloat(p.textContent) || 0;
    }
    return s;
  };
  for (const o of x.getElementsByTagName("Объект")){
    if (o.getAttribute("Тип") !== "ДокументСсылка.ОтчетОРозничныхПродажах") continue;
    const g = сум(o, "Товары", "Сумма"), gb = сум(o, "Возвраты", "Сумма"),
          pp = сум(o, "Предоплаты", "Сумма"), pb = сум(o, "Предоплаты", "СуммаВозврат"),
          c = сум(o, "Оплата", "СуммаОплаты"), cb = сум(o, "ВозвратОплаты", "СуммаОплаты"),
          a = сум(o, "ЗачетАвансов", "СуммаЗачета");
    итог.товары += g - gb; итог.предоплата += pp; итог.зачёт += a;
    итог.карты += c - cb; итог.наличные += g - gb + pp - pb - (c - cb) - a;
  }
  return итог;
});
/* выручка меньше исходной на 350 000 — дописан возврат, пробитый 24.08 */
const ЖДЁМ3 = {товары:61752240, предоплата:20182500, зачёт:15710400,
               карты:41238940, наличные:24293400};
for (const [k, v] of Object.entries(ЖДЁМ3))
  if (Math.round(баланс[k]) !== v)
    беда.push(`дописывание, ${k}: стало ${Math.round(баланс[k])}, ждали ${v}`);

/* Четвёртый прогон: один день против месячного отчёта ОФД — так проверяют
   день за днём. Фискальные данные должны обрезаться по периоду файла, иначе
   месяц кассы сравнивается с сутками выгрузки и всё расходится на миллионы. */
const день = join(TMP, "den.xml");
copyFileSync(join(ROOT, "data", "test", "день-13.08.xml"), день);
const стр4 = await browser.newPage();
стр4.on("pageerror", e => errors.push("один день, ошибка страницы: " + e.message));
await стр4.goto("file://" + APP);
const in4 = await стр4.locator(".slot input[type=file]").all();
await in4[0].setInputFiles(день);
await in4[1].setInputFiles(авг("Такском-фискальные-документы-11-31.08.2026.xlsx", "d-ofd.xlsx"));
await стр4.click("#run");
await стр4.waitForSelector("#out section", { timeout: 60000 });
const текст4 = ровно(await стр4.locator("#out").innerText());
if (!текст4.includes("Период 2026-08-13 — 2026-08-13"))
  беда.push("один день: период файла прочитан неверно");
if (!текст4.includes("Наличные сходятся с фискальными данными до рубля"))
  беда.push("один день: наличные не сошлись — фискальные данные не обрезаны по периоду");
if (!текст4.includes("Эквайринг на 1 548 500 ₽ переставлен"))
  беда.push("один день: не сработал перенос эквайринга 13.08");
if (!текст4.includes("Не объяснено 55 000 ₽"))
  беда.push("один день: остаток не 55 000 — считает не по этому дню");

/* Пятый прогон: деление по кассам. Приложение дописывает документам реквизит
   «Подразделение», которого в выгрузке нет, и добавляет в правила обмена
   недостающее правило для справочника подразделений. Проверяем и то, и другое:
   правил должно стать 49, у каждого кассового документа — своё подразделение,
   а у ссылки на него — владелец, иначе загрузка не найдёт элемент. */
const стр5 = await browser.newPage();
стр5.on("pageerror", e => errors.push("кассы, ошибка страницы: " + e.message));
await стр5.goto("file://" + APP);
await стр5.locator("details.set summary").first().click();
await стр5.selectOption("#c-podr", "все");
const in5 = await стр5.locator(".slot input[type=file]").all();
await in5[0].setInputFiles(авг("Выгрузка-в-БП-11-31.08.2026.xml", "c.xml"));
await in5[1].setInputFiles(авг("Такском-фискальные-документы-11-31.08.2026.xlsx", "c-ofd.xlsx"));
await стр5.click("#run");
await стр5.waitForSelector("#out section", { timeout: 60000 });

const текст5 = ровно(await стр5.locator("#out").innerText());
if (!текст5.includes("Деление по кассам: подразделение проставлено"))
  беда.push("кассы: в отчёте нет блока про деление по кассам");
/* названия подставляются из самой выгрузки — их же и ждём в документах */
for (const п of ["Форма строение 12", "Форма строение 1"])
  if (!текст5.includes(п + " — "))
    беда.push("кассы: в разбивке нет подразделения «" + п + "»");

const кассыXML = await стр5.evaluate(() => {
  const x = new DOMParser().parseFromString(FIXED_TEXT, "application/xml");
  const правил = x.getElementsByTagName("Правило").length;
  const поТипам = {}, нпп = {};
  let безВладельца = 0, всего = 0;
  for (const o of x.getElementsByTagName("Объект")){
    const тип = (o.getAttribute("Тип") || "").replace("ДокументСсылка.", "");
    const p = [...o.children].find(s => s.getAttribute("Имя") === "Подразделение");
    if (!p) continue;
    всего++;
    if (p.getAttribute("ИмяПКО") !== "Подразделения") безВладельца += 1000;
    const ссылка = p.querySelector("Ссылка");
    const имя = [...ссылка.children]
      .find(s => s.getAttribute("Имя") === "Наименование").textContent.trim();
    if (!ссылка.querySelector('[Имя="Владелец"] Ссылка')) безВладельца++;
    нпп[имя] = ссылка.getAttribute("Нпп");
    поТипам[тип] = (поТипам[тип] || 0) + 1;
  }
  return {правил, поТипам, безВладельца, всего, нпп};
});
if (кассыXML.правил !== 49)
  беда.push("кассы: правил в файле " + кассыXML.правил + ", ждали 49 (48 + подразделения)");
if (кассыXML.безВладельца)
  беда.push("кассы: ссылок на подразделение без владельца — " + кассыXML.безВладельца);
for (const тип of ["ПриходныйКассовыйОрдер", "РасходныйКассовыйОрдер", "ОтчетОРозничныхПродажах"])
  if (!кассыXML.поТипам[тип])
    беда.push("кассы: подразделение не проставлено в «" + тип + "»");
/* у каждого названия свой номер и он один на весь файл — так ссылки на один
   элемент справочника не размножатся в разные объекты */
const номера = Object.values(кассыXML.нпп);
if (номера.length !== new Set(номера).size)
  беда.push("кассы: два подразделения с одним Нпп");
if (номера.some(n => !n)) беда.push("кассы: ссылка на подразделение без Нпп");
if (!текст5.includes("Наличные сходятся с фискальными данными до рубля"))
  беда.push("кассы: деление тронуло деньги — наличные разъехались с ОФД");

/* Шестой прогон: месяц, обрезанный до одного дня. Выгружать день за днём
   из УМЦ утомительно — проще взять месяц и резать его двумя датами. Срез
   обязан совпасть с отдельным дневным файлом: те же документы, тот же период
   в шапке, те же деньги. Иначе резать месяц нельзя и об этом надо знать. */
const стр6 = await browser.newPage();
стр6.on("pageerror", e => errors.push("срез, ошибка страницы: " + e.message));
await стр6.goto("file://" + APP);
await стр6.locator("details.set summary").first().click();
await стр6.fill("#c-ot", "2026-08-13");
await стр6.fill("#c-do", "2026-08-13");
const in6 = await стр6.locator(".slot input[type=file]").all();
await in6[0].setInputFiles(авг("Выгрузка-в-БП-11-31.08.2026.xml", "e.xml"));
await in6[1].setInputFiles(авг("Такском-фискальные-документы-11-31.08.2026.xlsx", "e-ofd.xlsx"));
await стр6.click("#run");
await стр6.waitForSelector("#out section", { timeout: 60000 });

const текст6 = ровно(await стр6.locator("#out").innerText());
if (!текст6.includes("Период 2026-08-13 — 2026-08-13"))
  беда.push("срез: период в шапке не пересчитан по датам отсечки");
/* те же цифры, что даёт отдельный дневной файл в четвёртом прогоне */
for (const [имя, знач] of [["НАЛИЧНЫЕ", "1 303 000"], ["ЭКВАЙРИНГ", "4 337 800"],
                           ["ВЫРУЧКА", "2 904 300"], ["ЗАЧЁТ АВАНСА", "451 000"]]){
  const c = (await стр6.locator("#out .card").allInnerTexts()).map(ровно)
    .find(x => x.startsWith(имя + "\n"));
  if (!c || !c.includes(знач))
    беда.push("срез, " + имя + ": ждали " + знач + ", видно «" + c + "»");
}
if (!текст6.includes("Наличные сходятся с фискальными данными до рубля"))
  беда.push("срез: наличные не сошлись — фискальные данные не подрезались под срез");
const срезXML = await стр6.evaluate(() => {
  const x = new DOMParser().parseFromString(FIXED_TEXT, "application/xml");
  let док = 0;
  for (const o of x.getElementsByTagName("Объект"))
    if ((o.getAttribute("Тип") || "").indexOf("Документ") === 0) док++;
  const r = x.documentElement;
  return {док, от:r.getAttribute("НачалоПериодаВыгрузки"), до:r.getAttribute("ОкончаниеПериодаВыгрузки")};
});
if (срезXML.док !== 2)
  беда.push("срез: документов в файле " + срезXML.док + ", ждали 2 (столько 13.08)");
if (!/^2026-08-13/.test(срезXML.от) || !/^2026-08-13/.test(срезXML.до))
  беда.push("срез: период в самом файле не переписан — " + срезXML.от + " … " + срезXML.до);

await browser.close();
беда.push(...errors);

if (беда.length){
  console.error("НЕ СОШЛОСЬ:");
  беда.forEach(b => console.error("  · " + b));
  process.exit(1);
}
console.log("всё сошлось: сентябрь — карточки, " + ПРАВОК + " правок, наличные и терминалы " +
            "против ОФД, итог, памятка, скачивание; август — перенос эквайринга 13.08 " +
            "без изменения итога, дописывание авансов без сдвига наличных; один день против месячного ОФД; " +
            "деление по кассам через подразделения; месяц, обрезанный до одного дня");
