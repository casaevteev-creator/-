/* Проверка приложения на эталоне: 01–14.09.2026.
 *
 * Открывает собранный out/obmen.html настоящим браузером, кладёт в него три
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
  "РЕАЛИЗАЦИЯ": "43 549 220",
  "ПРЕДОПЛАТА": "13 429 900",
  "ЗАЧЁТ АВАНСА": "11 413 940",
  "НАЛИЧНЫЕ": "22 221 300",
  "ЭКВАЙРИНГ": "21 755 880",
  "ИНКАССАЦИЯ": "25 468 358",
  "ГРУПП": "4",
  "ЛИШНИХ ПЛАТЕЖЕЙ": "4",
  "СУММА ЛИШНИХ": "14 500"
};
const ПРАВОК = 68;

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
for (let i = 0; i < 3; i++) await inputs[i].setInputFiles(ФАЙЛЫ[i]);
for (let i = 0; i < 3; i++)
  if (!await inputs[i].evaluate(e => e.files.length))
    throw new Error("браузер не принял файл " + ФАЙЛЫ[i]);
await page.click("#run");
await page.waitForSelector("#out section", { timeout: 60000 });

const беда = [];
const наЭкране = await page.locator("#err").innerText();
if (наЭкране.trim()) беда.push("приложение показало ошибку: " + наЭкране.trim());

/* toLocaleString разделяет разряды неразрывным пробелом — сравниваем по обычному */
const ровно = s => s.replace(/[\u00A0\u202F\u2009]/g, " ").trim();
const карточки = {};
for (const t of await page.locator("#out .card").allInnerTexts()){
  const [l, v] = t.split("\n");
  if (карточки[l] === undefined) карточки[l] = ровно(v);   // первая карточка с этим названием
}
for (const [l, v] of Object.entries(ЖДЁМ)){
  if (карточки[l] !== v) беда.push(`${l}: показано «${карточки[l]}», ждали «${v}»`);
}

const свод = (await page.locator("#out .sb").allInnerTexts()).join(" ");
if (!свод.includes("Исправлено " + ПРАВОК)) беда.push("правок не " + ПРАВОК + ": " + свод);

const заметки = (await page.locator("#out .note, #out .why").allInnerTexts()).join(" ");
if (!заметки.includes("Наличные сошлись полностью"))
  беда.push("наличные не сошлись с ОФД");

/* исправленный файл должен скачиваться и оставаться валидным XML */
const ждём = page.waitForEvent("download", { timeout: 15000 });
await page.click("#save");
const dl = await ждём;
const path = await dl.path();
const text = await (await import("fs/promises")).readFile(path, "utf8");
if ((text.match(/<\?xml/g) || []).length !== 1) беда.push("в файле не одно объявление XML");
if (!text.includes("<ФайлОбмена")) беда.push("в скачанном файле нет корневого узла");
if ((text.match(/<Правило>/g) || []).length !== 48) беда.push("правила обмена повреждены");

await browser.close();
беда.push(...errors);

if (беда.length){
  console.error("НЕ СОШЛОСЬ:");
  беда.forEach(b => console.error("  · " + b));
  process.exit(1);
}
console.log("всё сошлось: карточки, " + ПРАВОК + " правок, наличные против ОФД, скачивание файла");
