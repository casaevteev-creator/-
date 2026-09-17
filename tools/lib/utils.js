/* Мелочи, на которых держатся оба приложения: приведение к строке и числу,
 * округление до копейки, формат денег, экранирование, склонение.
 * Разбирать выгрузки без них нельзя, а объявлять дважды — тем более.
 */
const S = v => v == null ? "" : String(v).replace(/\s+/g, " ").trim();
const num = v => {
  if (typeof v === "number") return v;
  const t = S(v).replace(/\s/g, "").replace(",", ".");
  const x = parseFloat(t);
  return isFinite(x) ? x : 0;
};
const fmt = x => (x == null || !isFinite(x)) ? "—"
  : x.toLocaleString("ru-RU", {minimumFractionDigits:2, maximumFractionDigits:2});
const fmt0 = x => (x == null || !isFinite(x)) ? "—" : Math.round(x).toLocaleString("ru-RU");
const esc = s => S(s).replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const R2 = x => Math.round((x + Number.EPSILON) * 100) / 100;
/* «1 группа, 2 группы, 5 групп» — иначе итог читается как машинный */
const skl0 = n => скл(n, "документ", "документа", "документов");
const скл = (n, одна, две, много) => {
  const a = Math.abs(n) % 100, b = a % 10;
  return n + " " + (a > 10 && a < 20 ? много : b === 1 ? одна : b > 1 && b < 5 ? две : много);
};
