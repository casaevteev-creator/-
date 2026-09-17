# -*- coding: utf-8 -*-
"""Сборка приложения «Обмен УМЦ → Бухгалтерия».

Две сборки из одного исходника — разница только в том, как отдаётся
исправленный файл и откуда берётся SheetJS.

    python3 tools/obmen_app/build.py              -> out/obmen.html
    python3 tools/obmen_app/build.py --artifact   -> out/obmen-artifact.html

Офлайн-версия кладёт SheetJS внутрь файла и сохраняет XML обычной ссылкой:
бухгалтер открывает её двойным щелчком рядом с 1С, интернет не нужен.

Версия для артефакта берёт SheetJS с cdnjs и сохраняет через песочницу.
Расширение .xml там не в списке разрешённых, поэтому файл отдаётся архивом
из одной записи — внутри XML с правильным именем.
"""
import argparse
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
HERE = Path(__file__).resolve().parent
SRC = HERE / "index.html"
SHELL = HERE / "artifact-shell.html"
LIB = ROOT / "node_modules" / "xlsx" / "dist" / "xlsx.full.min.js"

HEAD = """<!doctype html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<style>html{color-scheme:light}body{margin:0}img{max-width:100%}[hidden]{display:none!important}</style>
"""

# сохранение в песочнице артефактов: архив из одной записи, без сжатия
ARTIFACT_SAVE = r"""
/* В песочнице страница не скачивает файл сама — его отдаёт возможность
   downloads, а .xml не входит в список разрешённых расширений. Поэтому
   заворачиваем в архив из одной записи: внутри XML с нужным именем. */
let DOWNLOADS = null;
(async () => {
  try { DOWNLOADS = window.claude && claude.use ? await claude.use("downloads") : null; }
  catch (e){ DOWNLOADS = null; }
  $("#zip-hint").hidden = !DOWNLOADS;
  if (!DOWNLOADS) $("#save").hidden = true;
})();

function crc32(bytes){
  let table = crc32.t;
  if (!table){
    table = crc32.t = new Int32Array(256);
    for (let n = 0; n < 256; n++){
      let c = n;
      for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      table[n] = c;
    }
  }
  let crc = -1;
  for (let i = 0; i < bytes.length; i++) crc = (crc >>> 8) ^ table[(crc ^ bytes[i]) & 0xFF];
  return (crc ^ -1) >>> 0;
}
function zipOne(name, text){
  const enc = new TextEncoder();
  const data = enc.encode(text), nb = enc.encode(name), crc = crc32(data);
  const out = new Uint8Array(30 + nb.length + data.length + 46 + nb.length + 22);
  const v = new DataView(out.buffer);
  let o = 0;
  const w32 = x => { v.setUint32(o, x, true); o += 4; };
  const w16 = x => { v.setUint16(o, x, true); o += 2; };
  /* локальный заголовок: без сжатия, имя в UTF-8 (флаг 0x0800) */
  w32(0x04034b50); w16(20); w16(0x0800); w16(0); w16(0); w16(0);
  w32(crc); w32(data.length); w32(data.length); w16(nb.length); w16(0);
  out.set(nb, o); o += nb.length;
  out.set(data, o); o += data.length;
  const central = o;
  w32(0x02014b50); w16(20); w16(20); w16(0x0800); w16(0); w16(0); w16(0);
  w32(crc); w32(data.length); w32(data.length);
  w16(nb.length); w16(0); w16(0); w16(0); w16(0); w32(0); w32(0);
  out.set(nb, o); o += nb.length;
  const size = o - central;
  /* конец центрального каталога: одна запись */
  w32(0x06054b50); w16(0); w16(0); w16(1); w16(1); w32(size); w32(central); w16(0);
  return out.slice(0, o);
}

$("#save").addEventListener("click", async () => {
  if (!DOWNLOADS) return;
  try {
    await DOWNLOADS.save({
      filename: FIXED_NAME.replace(/\.xml$/i, "") + ".zip",
      data: new Blob([zipOne(FIXED_NAME, FIXED_TEXT)])
    });
  } catch (e){
    const code = (e && e.code) || "";
    const текст = code === "declined" ? "Сохранение отменено."
      : code === "rate_limited" ? "Окно сохранения уже открыто — закройте его и нажмите ещё раз."
      : "Файл сохранить не удалось: " + esc((e && e.message) || code);
    $("#err").innerHTML = '<div class="err">' + текст + "</div>";
  }
});
"""


def логика():
    """Код приложения из index.html — он общий для обеих сборок."""
    html = SRC.read_text(encoding="utf-8")
    js = html.rsplit("<script>", 1)[1].rsplit("</script>", 1)[0]
    return html, js


def офлайн():
    html, _ = логика()
    if not LIB.exists():
        raise SystemExit(f"нет SheetJS: {LIB}\nпоставьте его: npm install xlsx@0.18.5 --no-save")
    html = html.replace(
        "__SHEETJS__",
        "/* SheetJS 0.18.5, Apache-2.0 — вшит в файл, чтобы работало без интернета */\n"
        + LIB.read_text(encoding="utf-8"))
    head, rest = html.split("</style>", 1)
    return HEAD + head + "</style>\n</head>\n<body>\n" + rest + "\n</body>\n</html>\n"


def артефакт():
    _, js = логика()
    js = re.sub(r"/\*СОХРАНЕНИЕ\*/.*?/\*КОНЕЦ СОХРАНЕНИЯ\*/", ARTIFACT_SAVE.strip(),
                js, flags=re.S)
    if "DOWNLOADS" not in js:
        raise SystemExit("в index.html не нашлась разметка /*СОХРАНЕНИЕ*/")
    return SHELL.read_text(encoding="utf-8").replace("__ЛОГИКА__", js)


def main():
    ap = argparse.ArgumentParser(description="Сборка приложения обмена")
    ap.add_argument("--artifact", action="store_true",
                    help="версия для публикации артефактом")
    a = ap.parse_args()
    out = ROOT / "out" / ("obmen-artifact.html" if a.artifact else "obmen.html")
    doc = артефакт() if a.artifact else офлайн()
    out.parent.mkdir(exist_ok=True)
    out.write_text(doc, encoding="utf-8")
    print(f"собрано: {out}  ({len(doc):,} символов)".replace(",", " "))


if __name__ == "__main__":
    main()
