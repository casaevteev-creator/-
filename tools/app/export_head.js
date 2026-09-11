"use strict";
/* Самодостаточный отчёт: цифры месяца заморожены в файле,
   разбор работает так же, как в программе. */
const FROZEN = window.FROZEN;
const $ = id => document.getElementById(id);
const S = v => v == null ? "" : String(v).replace(/\s+/g, " ").trim();
const N = v => typeof v === "number" ? v : 0;
const R2 = x => Math.round((x + Number.EPSILON) * 100) / 100;
const fmt = x => (x == null || !isFinite(x)) ? "—"
  : (Math.abs(x) < 0.005 ? "0,00" : x.toLocaleString("ru-RU", {minimumFractionDigits:2, maximumFractionDigits:2}));
const fmt0 = x => (x == null || !isFinite(x)) ? "—" : Math.round(x).toLocaleString("ru-RU");
const pct = x => (!isFinite(x) ? "—" : (x * 100).toLocaleString("ru-RU", {minimumFractionDigits:1, maximumFractionDigits:1}) + " %");
const esc = s => S(s).replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const MONTHS = ["январь","февраль","март","апрель","май","июнь","июль","август","сентябрь","октябрь","ноябрь","декабрь"];
const monthName = v => { const m = /(\d{4})-(\d{2})/.exec(v || ""); return m ? MONTHS[+m[2]-1] + " " + m[1] : v; };
const shiftMonth = (m, d) => {
  const x = /(\d{4})-(\d{2})/.exec(m); if (!x) return "";
  let y = +x[1], mm = +x[2] + d;
  while (mm < 1){ mm += 12; y--; } while (mm > 12){ mm -= 12; y++; }
  return y + "-" + String(mm).padStart(2, "0");
};
const st = {params:FROZEN.params, psrc:FROZEN.psrc, taksk:FROZEN.taksk, groups:FROZEN.groups,
  vendors:FROZEN.vendors, month:FROZEN.month, acc:FROZEN.acc, periods:{}, isExample:false};
let book = FROZEN.book;
let hist = FROZEN.hist;
function calc(){ return FROZEN.c; }
window.__calc = FROZEN.c;
