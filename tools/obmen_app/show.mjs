/* Печатает в консоль то, что приложение показало бы на экране.
 *
 * Нужен, когда разбираешься с периодом: открывать браузер и листать глазами
 * долго, а так весь разбор виден текстом и по нему можно грепать.
 * Развёрнутые блоки («дни, где терминал не сошёлся», «погашенные пары»)
 * раскрываются сами — иначе их содержимое в текст не попадёт.
 *
 *     npm install xlsx@0.18.5 playwright@1.49.1 --no-save
 *     python3 tools/obmen_app/build.py
 *     node tools/obmen_app/show.mjs 2026-08
 *     node tools/obmen_app/show.mjs 2026-09 | sed -n '/§ 3/,/§ 4/p'
 *
 * Берёт из data/source/<период> всё, что там лежит: xml обязателен,
 * xlsx-файлы приложение само разбирает по содержимому.
 */
import pkg from "playwright";
import { fileURLToPath } from "url";
import { dirname, join, extname } from "path";
import { existsSync, readdirSync, copyFileSync, mkdtempSync } from "fs";
import { tmpdir } from "os";

const { chromium } = pkg;
const ROOT = join(dirname(fileURLToPath(import.meta.url)), "..", "..");
const период = process.argv[2];
if (!период){
  console.error("укажите период, например: node tools/obmen_app/show.mjs 2026-08");
  process.exit(2);
}
const SRC = join(ROOT, "data", "source", период);
const APP = join(ROOT, "out", process.env.APP || "obmen.html");
if (!existsSync(SRC)) { console.error("нет папки " + SRC); process.exit(2); }
if (!existsSync(APP)) { console.error("нет " + APP + " — соберите: python3 tools/obmen_app/build.py"); process.exit(2); }

/* В папке периода лежат не только файлы обмена, но и выгрузки из самой
   бухгалтерии — их приложению давать нельзя, оно на них ругается. Поэтому
   берём по одному файлу каждого нужного вида, узнавая их по имени. */
const ВИДЫ = [
  [/\.xml$/i,                       "обмен"],
  [/такском|фискальн/i,             "ОФД"],
  [/услуг/i,                        "реестр услуг"],
  [/оплат/i,                        "реестр оплат"]
];
const TMP = mkdtempSync(join(tmpdir(), "obmen-show-"));
/* В папке лежат и файлы из старой программы, и выгрузки самой бухгалтерии.
   Отодвигаем их в конец, чтобы по имени выбирался файл ровно того периода. */
const второсорт = f => /^(1С-|Новая-|Старая-)/i.test(f) ? 1 : 0;
const все = readdirSync(SRC).filter(f => /\.(xml|xlsx)$/i.test(f))
  .sort((a, b) => второсорт(a) - второсорт(b) || a.localeCompare(b, "ru"));
const файлы = [], занято = new Set();
for (const [маска, что] of ВИДЫ){
  const f = все.find(x => !занято.has(x) && маска.test(x) &&
    (что === "обмен" ? true : /\.xlsx$/i.test(x)));
  if (!f){ if (что !== "обмен") console.error("нет файла: " + что); continue; }
  занято.add(f);
  const to = join(TMP, "f" + файлы.length + extname(f));
  copyFileSync(join(SRC, f), to);
  файлы.push(to);
  console.error("· " + что + ": " + f);
}
if (!файлы.length){ console.error("в " + SRC + " нет ни файла обмена, ни таблиц"); process.exit(2); }

const BIN = ["/opt/pw-browsers/chromium-1194/chrome-linux/chrome",
             "/opt/pw-browsers/chromium/chrome-linux/chrome"].find(existsSync);
const browser = await chromium.launch(BIN ? { executablePath: BIN } : {});
const page = await browser.newPage();
page.on("pageerror", e => console.error("ошибка страницы: " + e.message));

await page.goto("file://" + APP);
const ячейки = await page.locator(".slot input[type=file]").all();
for (let i = 0; i < файлы.length && i < ячейки.length; i++)
  await ячейки[i].setInputFiles(файлы[i]);
await page.click("#run");
await page.waitForSelector("#out section", { timeout: 60000 });

const ошибка = await page.locator("#err").innerText();
if (ошибка.trim()) console.error("приложение показало ошибку: " + ошибка.trim());
for (const d of await page.locator("#out details").all()) await d.evaluate(e => e.open = true);
console.log(await page.locator("#out").innerText());
await browser.close();
