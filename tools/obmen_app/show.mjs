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

/* браузер не принимает файлы с кириллицей в имени — кладём копии с латинскими */
const TMP = mkdtempSync(join(tmpdir(), "obmen-show-"));
const все = readdirSync(SRC);
const взять = маска => все.filter(f => маска.test(f))
  .map((f, i) => { const to = join(TMP, "f" + i + extname(f)); copyFileSync(join(SRC, f), to); return to; });
const файлы = [...взять(/\.xml$/i), ...взять(/\.xlsx$/i)].slice(0, 4);
if (!файлы.length){ console.error("в " + SRC + " нет ни одного файла"); process.exit(2); }

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
