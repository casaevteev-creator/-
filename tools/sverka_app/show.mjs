/* Печатает в консоль то, что приложение показало бы на экране.
 *
 * Нужен, когда разбираешься с периодом: открывать браузер и листать глазами
 * долго, а так весь разбор виден текстом и по нему можно грепать.
 * Находки раскрываются сами — иначе их содержимое в текст не попадёт.
 *
 *     npm install xlsx@0.18.5 playwright@1.49.1 --no-save
 *     python3 tools/sverka_app/build.py
 *     node tools/sverka_app/show.mjs 2026-08
 *     node tools/sverka_app/show.mjs 2026-09 | sed -n '/Находки/,/Кассы/p'
 *
 * Файлы берутся из data/source/<период> по именам.
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
  console.error("укажите период, например: node tools/sverka_app/show.mjs 2026-08");
  process.exit(2);
}
const SRC = join(ROOT, "data", "source", период);
const APP = join(ROOT, "out", process.env.APP || "sverka.html");
if (!existsSync(SRC)) { console.error("нет папки " + SRC); process.exit(2); }
if (!existsSync(APP)) { console.error("нет " + APP + " — соберите: python3 tools/sverka_app/build.py"); process.exit(2); }

/* В папке периода лежат выгрузки и прошлых задач тоже — отодвигаем их
   в конец, чтобы по имени выбирался файл ровно этого периода. */
export const ВИДЫ = [
  [/такском|фискальн/i, "отчёт Такскома"],
  [/оплат/i,            "реестр оплат"],
  [/услуг/i,            "реестр услуг"]
];
const второсорт = f => /^(1С-|Новая-|Старая-)/i.test(f) ? 1 : 0;

export function файлыПериода(dir){
  const все = readdirSync(dir).filter(f => /\.xlsx$/i.test(f))
    .sort((a, b) => второсорт(a) - второсорт(b) || a.localeCompare(b, "ru"));
  const занято = new Set(), из = [];
  for (const [маска, что] of ВИДЫ){
    const f = все.find(x => !занято.has(x) && маска.test(x));
    if (!f){ console.error("нет файла: " + что); continue; }
    занято.add(f);
    из.push({путь: join(dir, f), имя: f, что});
  }
  return из;
}

const файлы = файлыПериода(SRC);
if (!файлы.length){ console.error("в " + SRC + " нет подходящих таблиц"); process.exit(2); }
const TMP = mkdtempSync(join(tmpdir(), "sverka-show-"));
const копии = файлы.map((ф, i) => {
  const to = join(TMP, "f" + i + extname(ф.имя));
  copyFileSync(ф.путь, to);
  console.error("· " + ф.что + ": " + ф.имя);
  return to;
});

const BIN = ["/opt/pw-browsers/chromium-1194/chrome-linux/chrome",
             "/opt/pw-browsers/chromium/chrome-linux/chrome"].find(existsSync);
const browser = await chromium.launch(BIN ? { executablePath: BIN } : {});
const page = await browser.newPage();
page.on("pageerror", e => console.error("ошибка страницы: " + e.message));

await page.goto("file://" + APP);
const ячейки = await page.locator(".slot input[type=file]").all();
for (let i = 0; i < копии.length && i < ячейки.length; i++)
  await ячейки[i].setInputFiles(копии[i]);
await page.click("#run");
await page.waitForSelector("#out section", { timeout: 60000 });

const ошибка = await page.locator("#err").innerText();
if (ошибка.trim()) console.error("приложение показало ошибку: " + ошибка.trim());
for (const d of await page.locator("#out details").all()) await d.evaluate(e => e.open = true);
console.log(await page.locator("#out").innerText());
await browser.close();
