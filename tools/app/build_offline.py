# -*- coding: utf-8 -*-
"""Автономная сборка: один html-файл, работает без интернета."""
import re, io, os, subprocess

html = open('/home/user/-/out/upravlenka.html', encoding='utf-8').read()

# 1. шрифты Google — убираем; в CSS уже прописаны системные запасные
html = re.sub(r'<link rel="preconnect"[^>]*>\s*', '', html)
html = re.sub(r'<link rel="stylesheet" href="https://fonts\.googleapis[^>]*>\s*', '', html)

# 2. SheetJS кладём внутрь файла
lib = open('node_modules/xlsx/dist/xlsx.full.min.js', encoding='utf-8').read()
html = html.replace(
    '<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>',
    '<script>/* SheetJS 0.18.5, Apache-2.0 — вшит в файл, чтобы работало без интернета */\n'
    + lib + '\n</script>')

# 3. сохранение файлов — обычной ссылкой, без платформы
start = html.index('/*OFFER*/'); end = html.index('/*ENDOFFER*/') + len('/*ENDOFFER*/')
html = html[:start] + '''function offer(filename, blob){
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
  return Promise.resolve();
}''' + html[end:]

# 4. подпись в боковой заметке
html = html.replace(
    'Файлы не покидают компьютер — разбор идёт в браузере. Сохраняется только справочник статей.',
    'Автономная версия. Интернет не нужен, файлы никуда не уходят. Справочник хранится в этом браузере — кнопка «Сохранить справочник» переносит его на другой компьютер.')

# 4б. в автономной версии ссылки нет, а справочник не общий
html = html.replace('<button class="btn ghost sm" id="copyLink">Ссылка для учредителей</button>\n      ', '')
html = html.replace('id="copyLink"', 'id="copyLink" hidden')
html = html.replace(
    'Справочник статей, состав ассистентов и подрядчики по стройке сохраняются между месяцами и общие для всех, кто открывает эту ссылку. Выгрузки нигде не сохраняются.',
    'Справочник статей, подрядчики по стройке и закрытые месяцы хранятся в этом браузере и переживают перезапуск. Выгрузки нигде не сохраняются.')

# 5. полноценный документ
doc = ('<!doctype html>\n<html lang="ru">\n<head>\n<meta charset="utf-8">\n'
       '<meta name="viewport" content="width=device-width,initial-scale=1">\n'
       '<style>html{color-scheme:light dark}body{margin:0}img{max-width:100%}'
       '[hidden]{display:none!important}</style>\n'
       + html.split('<style>', 1)[0]
       + '<style>' + html.split('<style>', 1)[1].split('</style>', 1)[0] + '</style>\n</head>\n<body>\n'
       + html.split('</style>', 1)[1] + '\n</body>\n</html>\n')

out = '/home/user/-/out/Управленка-Формы-офлайн.html'
open(out, 'w', encoding='utf-8').write(doc)
print('записано:', out, round(len(doc.encode('utf-8'))/1048576, 2), 'МБ')
