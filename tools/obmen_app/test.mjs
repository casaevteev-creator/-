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
if (!всё.includes("Наличные сошлись полностью")) беда.push("наличные не сошлись с ОФД");
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

await browser.close();
беда.push(...errors);

if (беда.length){
  console.error("НЕ СОШЛОСЬ:");
  беда.forEach(b => console.error("  · " + b));
  process.exit(1);
}
console.log("всё сошлось: сентябрь — карточки, " + ПРАВОК + " правок, наличные и терминалы " +
            "против ОФД, итог, памятка, скачивание; август — перенос эквайринга 13.08 " +
            "без изменения итога, дописывание авансов без сдвига наличных; один день против месячного ОФД");
