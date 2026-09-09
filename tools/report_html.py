# -*- coding: utf-8 -*-
"""Генератор HTML-отчёта учредителям из расчётной модели.

Один самодостаточный файл, mobile-first, палитра и структура — по брифу.
Ни одна цифра не зашита в шаблон: всё приходит из model.compute().
"""
import json
from html import escape

from model import FOUNDERS, GROUP_RU, GROUPS, MONTHS_RU

NBSP = " "


def mln(v, dec=2):
    if v is None:
        return "—"
    return f"{v/1e6:,.{dec}f}".replace(",", NBSP).replace(".", ",")


def rub(v):
    if v is None:
        return "—"
    return f"{round(v):,}".replace(",", NBSP)


def pct(v, dec=1, sign=True):
    if v is None:
        return "—"
    s = f"{v:+.{dec}f}" if sign else f"{v:.{dec}f}"
    return s.replace(".", ",") + "%"


def arrow(v):
    if v is None:
        return ""
    return ("▲ " if v >= 0 else "▼ ") + pct(v)


def _donut(parts, total, size=260):
    """SVG-кольцо: parts = [(подпись, значение, цвет)]."""
    r, sw = size / 2 - 22, 34
    cx = cy = size / 2
    circ = 2 * 3.141592653589793 * r
    segs, off = [], 0.0
    for label, value, color in parts:
        frac = (value / total) if total else 0
        dash = circ * frac
        segs.append(
            f'<circle class="dseg" cx="{cx}" cy="{cy}" r="{r:.1f}" fill="none" stroke="{color}" '
            f'stroke-width="{sw}" stroke-dasharray="{dash:.2f} {circ - dash:.2f}" '
            f'stroke-dashoffset="{-off:.2f}" transform="rotate(-90 {cx} {cy})">'
            f'<title>{escape(label)} — {mln(value)} млн ₽</title></circle>'
        )
        off += dash
    return (
        f'<svg class="donut" viewBox="0 0 {size} {size}" role="img" aria-label="Структура выручки">'
        + "".join(segs)
        + f'<text x="{cx}" y="{cy-4}" text-anchor="middle" class="dcv">{mln(total)}</text>'
        + f'<text x="{cx}" y="{cy+18}" text-anchor="middle" class="dcl">млн ₽ выручка</text>'
        + "</svg>"
    )


def _summary(r, n):
    """5 тезисов резюме. Текст можно переопределить в canonical['narrative']['summary']."""
    if n.get("summary"):
        return n["summary"]
    meta, rev, tr, f = r["meta"], r["revenue"], r["trend"], r["founders"]
    mon = meta["month_ru"].lower()
    out = [{
        "h": f"Выручка месяца — {mln(rev['total'])} млн ₽",
        "p": (f"К {mon}ю {meta['year']-1} — <b>{arrow(tr['yoy_month'])}</b>."
              if tr.get("yoy_month") is not None else
              "Сравнение год к году недоступно: нет данных за тот же месяц прошлого года."),
    }, {
        "h": f"С начала года — {mln(sum(v for v in tr['cur'][:meta['month']] if v))} млн ₽",
        "p": (f"Рост <b>{arrow(tr['ytd_growth'])}</b> к тому же периоду прошлого года "
              f"({mln(tr['ytd_prev'])} млн ₽)." if tr.get("ytd_growth") is not None else
              "Сравнение с прошлым годом неполное: за "
              + ", ".join(tr.get("missing_prev") or []) + " данных в архиве нет."),
    }, {
        "h": f"Денежная позиция — {mln(r['cashflow']['closing'])} млн ₽ на счёте",
        "p": (f"Против {mln(r['cashflow']['opening'])} млн ₽ на начало месяца. "
              f"Проценты по депозитам за месяц — <b>{rub(r['manual'].get('deposit_interest'))} ₽</b>."),
    }, {
        "h": f"Расходы по счёту 60 — {mln(r['expenses']['acc60_total'])} млн ₽",
        "p": (f"Это {pct(r['expenses']['share_of_revenue'], 1, sign=False)} выручки. "
              f"Медицинские и ИП — {mln(r['expenses']['medical_total'])} млн, "
              f"управленческие — {mln(r['expenses']['management_total'])} млн."),
    }, {
        "h": f"Доход учредителей — {mln(r['founders_total_income'])} млн ₽",
        "p": (f"{f['pog']['name']} — <b>{mln(f['pog']['income'])} млн</b>, "
              f"{f['mal']['name']} — <b>{mln(f['mal']['income'])} млн</b>. "
              f"С учётом накопленных остатков к выплате: {mln(f['pog']['payout'])} и "
              f"{mln(f['mal']['payout'])} млн ₽."),
    }]
    return out


CSS = """
:root{--ink:#1C1B18;--ink2:#423E37;--mute:#807A70;--bg:#FAF9F6;--panel:#FFFFFF;--tint:#F1EFEA;
--line:#E7E3D9;--green:#2E5E4E;--green2:#4A7A68;--green3:#7BA08F;--bronze:#8A7350;--bronze2:#B9895A;
--red:#A3423B;--ghost:#D9D4C9}
*{margin:0;padding:0;box-sizing:border-box}
html{scroll-behavior:smooth}
body{background:var(--bg);color:var(--ink);font-family:'Inter',system-ui,sans-serif;font-size:16px;line-height:1.6;-webkit-font-smoothing:antialiased;overflow-x:hidden}
.mono{font-family:'IBM Plex Mono',monospace}
.rv{opacity:0;transform:translateY(24px);transition:opacity .8s cubic-bezier(.2,.6,.2,1),transform .8s cubic-bezier(.2,.6,.2,1)}
.rv.on{opacity:1;transform:none}
.rv.d1{transition-delay:.1s}.rv.d2{transition-delay:.2s}.rv.d3{transition-delay:.3s}
@media (prefers-reduced-motion:reduce){html{scroll-behavior:auto}.rv{opacity:1;transform:none;transition:none}}
nav{position:fixed;right:22px;top:50%;transform:translateY(-50%);z-index:50;display:flex;flex-direction:column;gap:13px}
nav a{width:8px;height:8px;border-radius:50%;background:var(--ghost);display:block;transition:all .3s;position:relative}
nav a.act{background:var(--bronze2);transform:scale(1.5)}
nav a:hover::after{content:attr(data-t);position:absolute;right:20px;top:50%;transform:translateY(-50%);background:var(--ink);color:#fff;font-size:11px;padding:4px 10px;border-radius:6px;white-space:nowrap}
@media(max-width:1000px){nav{display:none}}
.wrap{max-width:1060px;margin:0 auto;padding:0 20px}
@media(min-width:700px){.wrap{padding:0 32px}}
section{padding:64px 0}
@media(min-width:700px){section{padding:88px 0}}
.kick{font-family:'IBM Plex Mono',monospace;font-size:11px;letter-spacing:.26em;text-transform:uppercase;color:var(--bronze);font-weight:600}
h2{font-family:'Cormorant Garamond',Cambria,serif;font-weight:600;font-size:clamp(30px,6vw,50px);line-height:1.08;margin:12px 0 8px}
.sub{font-family:'Cormorant Garamond',Cambria,serif;font-style:italic;font-size:18px;color:var(--mute);margin-bottom:32px}
.hero{background:var(--ink);color:#fff;min-height:88vh;display:flex;flex-direction:column;justify-content:center;position:relative;overflow:hidden;padding:80px 0}
.hero::before{content:"";position:absolute;right:-180px;top:-180px;width:560px;height:560px;border-radius:50%;border:1px solid rgba(185,137,90,.22)}
.hero::after{content:"";position:absolute;right:-80px;top:-80px;width:360px;height:360px;border-radius:50%;border:1px solid rgba(185,137,90,.36)}
.hero .brand{font-size:13px;font-weight:700;letter-spacing:.55em;margin-bottom:10px}
.hero .tag{font-family:'Cormorant Garamond',Cambria,serif;font-style:italic;color:#9C968A;font-size:16px;max-width:32ch}
.hero h1{font-family:'Cormorant Garamond',Cambria,serif;font-weight:600;font-size:clamp(42px,10vw,92px);line-height:1;margin:36px 0 6px}
.hero .yr{font-family:'Cormorant Garamond',Cambria,serif;font-style:italic;font-weight:500;font-size:clamp(30px,7vw,60px);color:var(--bronze2)}
.hero .strip{display:grid;grid-template-columns:1fr 1fr;gap:24px 32px;margin-top:48px;max-width:640px}
@media(min-width:760px){.hero .strip{display:flex;gap:44px;flex-wrap:wrap;max-width:none}}
.hero .strip .v{font-family:'IBM Plex Mono',monospace;font-size:clamp(20px,5vw,25px);font-weight:600}
.hero .strip .l{font-size:12px;color:#9C968A;margin-top:4px}
.tez{display:grid;grid-template-columns:44px 1fr;gap:12px 20px;padding:22px 0;border-bottom:1px solid var(--line);align-items:start}
@media(min-width:860px){.tez{grid-template-columns:64px 320px 1fr;gap:24px}}
.tez:last-child{border-bottom:none}
.tez .n{font-family:'Cormorant Garamond',Cambria,serif;font-size:30px;font-weight:700;color:var(--bronze);line-height:1}
.tez h3{font-size:16px;font-weight:700;line-height:1.35}
.tez p{grid-column:1/-1;font-family:'Cormorant Garamond',Cambria,serif;font-size:18px;line-height:1.5;color:var(--ink2)}
@media(min-width:860px){.tez p{grid-column:auto}}
.tez p b{font-family:'Inter',sans-serif;font-size:15px;font-weight:700;color:var(--green)}
.kpi{display:grid;grid-template-columns:1fr;gap:14px}
@media(min-width:620px){.kpi{grid-template-columns:repeat(auto-fit,minmax(270px,1fr));gap:18px}}
.card{background:var(--panel);border:1px solid var(--line);border-radius:18px;padding:24px 26px}
.card .v{font-family:'IBM Plex Mono',monospace;font-size:clamp(28px,7vw,36px);font-weight:600;letter-spacing:-.02em}
.card .v small{font-size:15px;color:var(--mute);font-weight:400}
.card .l{font-size:11px;font-weight:700;letter-spacing:.13em;text-transform:uppercase;margin-top:10px}
.card .d{font-size:13px;margin-top:6px;color:var(--mute)}
.up{color:var(--green)}.neg{color:var(--red)}
.chart{background:var(--panel);border:1px solid var(--line);border-radius:18px;padding:24px 16px 14px}
@media(min-width:700px){.chart{padding:30px 26px 18px}}
.bars{display:flex;align-items:flex-end;gap:0;height:250px;border-bottom:1.5px solid var(--ink);padding:0 4px}
@media(min-width:700px){.bars{height:300px}}
.bgrp{flex:1;display:flex;align-items:flex-end;justify-content:center;gap:4px;height:100%}
.bar{width:min(6vw,34px);border-radius:5px 5px 0 0;position:relative;height:0;transition:height 1.1s cubic-bezier(.2,.7,.2,1)}
.bar.g25{background:var(--ghost)}.bar.g26{background:var(--green)}.bar.acc{background:var(--bronze2)}
.bar .bv{position:absolute;top:-22px;left:50%;transform:translateX(-50%);font-family:'IBM Plex Mono',monospace;font-size:11px;white-space:nowrap;opacity:0;transition:opacity .5s .8s}
.on .bar .bv{opacity:1}
@media(max-width:700px){.bar .bv{display:none}.bar{width:min(4.5vw,34px)}.bgrp{gap:2px}.bars{padding:0 2px}}
.bx{display:flex;padding:10px 4px 0}
.bx span{flex:1;text-align:center;font-family:'IBM Plex Mono',monospace;font-size:11px;color:var(--mute)}
.leg{display:flex;gap:18px;padding:14px 4px 4px;font-size:13px;color:var(--mute);flex-wrap:wrap}
.dot{display:inline-block;width:11px;height:11px;border-radius:3.5px;margin-right:7px;vertical-align:-1px}
.hl{background:var(--tint);border-radius:18px;padding:24px 26px;margin-top:18px;display:flex;gap:24px;flex-wrap:wrap;align-items:center}
.hl .big{font-family:'Cormorant Garamond',Cambria,serif;font-size:46px;font-weight:700;color:var(--green);line-height:1}
.hl p{font-family:'Cormorant Garamond',Cambria,serif;font-size:18px;color:var(--ink2);max-width:520px}
.dsplit{display:grid;grid-template-columns:1fr;gap:22px;align-items:center}
@media(min-width:820px){.dsplit{grid-template-columns:300px 1fr;gap:32px}}
.donut{width:100%;max-width:300px;margin:0 auto;display:block}
.dseg{stroke-dasharray:0 9999;transition:stroke-dasharray 1.1s cubic-bezier(.2,.7,.2,1)}
.dcv{font-family:'IBM Plex Mono',monospace;font-size:30px;font-weight:600;fill:var(--ink)}
.dcl{font-size:11px;fill:var(--mute);letter-spacing:.1em;text-transform:uppercase}
.dail{display:flex;align-items:flex-end;gap:2px;height:160px;border-bottom:1.5px solid var(--ink);padding:0 2px}
.dail div{flex:1;background:var(--green);border-radius:3px 3px 0 0;height:0;transition:height .9s cubic-bezier(.2,.7,.2,1);position:relative;cursor:pointer}
.dail div.act::after,.dail div:hover::after{content:attr(data-v);position:absolute;bottom:calc(100% + 8px);left:50%;transform:translateX(-50%);background:var(--ink);color:#fff;font-size:11px;font-family:'IBM Plex Mono',monospace;padding:4px 9px;border-radius:6px;white-space:nowrap;z-index:5}
.dx{display:flex;gap:2px;padding-top:8px}
.dx span{flex:1;text-align:center;font-family:'IBM Plex Mono',monospace;font-size:10px;color:var(--mute)}
.split{display:grid;grid-template-columns:1fr;gap:18px}
@media(min-width:860px){.split{grid-template-columns:1fr 1fr;gap:20px}}
.stackbar{display:flex;height:52px;border-radius:14px;overflow:hidden;margin:6px 0 20px}
.stackbar div{display:flex;align-items:center;justify-content:center;color:#fff;font-family:'IBM Plex Mono',monospace;font-size:12px;white-space:nowrap;width:0;transition:width 1.2s cubic-bezier(.2,.7,.2,1)}
.tw{overflow-x:auto}
table{width:100%;border-collapse:collapse;background:var(--panel);border:1px solid var(--line);border-radius:16px;overflow:hidden}
th{font-size:10.5px;text-transform:uppercase;letter-spacing:.11em;color:var(--mute);font-weight:700;text-align:left;padding:13px 16px;border-bottom:1px solid var(--line);background:#FCFBF8}
td{padding:11px 16px;border-bottom:1px solid var(--line);font-size:14px}
tr:last-child td{border-bottom:none}
td.num,th.num{text-align:right;font-family:'IBM Plex Mono',monospace;font-size:13px;white-space:nowrap}
tr.tot td{font-weight:700;background:#EDF1EE}
.wf{display:flex;align-items:flex-end;gap:6px;height:300px;border-bottom:1.5px solid var(--ink);padding:0 2px;position:relative}
@media(min-width:700px){.wf{gap:14px;height:360px}}
.wcol{flex:1;position:relative;height:100%}
.wbar{position:absolute;left:0;right:0;border-radius:4px;opacity:0;transform:scaleY(.6);transform-origin:bottom;transition:opacity .6s,transform .6s cubic-bezier(.2,.7,.2,1)}
.on .wbar{opacity:1;transform:none}
.wbar.pos{background:var(--green)}.wbar.exp{background:#CCC5B7}.wbar.inv{background:var(--bronze2)}
.wbridge{position:absolute;left:0;right:-8px;border-top:1px dashed var(--ghost);opacity:0;transition:opacity .5s}
.on .wbridge{opacity:1}
.wv{position:absolute;left:50%;transform:translateX(-50%);font-family:'IBM Plex Mono',monospace;font-size:11px;font-weight:600;white-space:nowrap;opacity:0;transition:opacity .5s}
@media(min-width:700px){.wv{font-size:13px}}
.on .wv{opacity:1}
.wl{display:flex;gap:6px;padding:12px 2px 0}
@media(min-width:700px){.wl{gap:14px}}
.wl span{flex:1;min-width:0;overflow-wrap:break-word;text-align:center;font-size:10px;color:var(--ink2);line-height:1.25}
.wfscroll{overflow-x:auto;-webkit-overflow-scrolling:touch}
.wfinner{min-width:660px}
@media(min-width:720px){.wfinner{min-width:0}}
.bx span,.dx span{min-width:0}
@media(min-width:700px){.wl span{font-size:12px}}
.kop{margin-top:22px;background:var(--tint);border-radius:18px;padding:24px 26px}
.kop p{font-family:'Cormorant Garamond',Cambria,serif;font-size:18px;color:var(--ink2)}
.kop b{color:var(--green)}
.dark{background:var(--ink);color:#fff}
.dark h2{color:#fff}.dark .sub{color:#9C968A}
.own{display:grid;grid-template-columns:1fr;gap:18px}
@media(min-width:860px){.own{grid-template-columns:1fr 1fr;gap:22px}}
.ocard{background:rgba(255,255,255,.045);border:1px solid rgba(217,212,201,.16);border-radius:20px;padding:28px 24px 24px}
@media(min-width:700px){.ocard{padding:34px 34px 28px}}
.ocard h3{font-family:'Cormorant Garamond',Cambria,serif;font-size:30px;font-weight:600;margin-bottom:18px}
.orow{display:flex;justify-content:space-between;gap:14px;padding:9px 0;border-bottom:1px solid rgba(217,212,201,.12);font-size:14px;color:#CFC9BD}
.orow .n{font-family:'IBM Plex Mono',monospace;font-weight:600;color:#fff;white-space:nowrap}
.orow .n.neg{color:#D98A82}
.oinc{display:flex;justify-content:space-between;align-items:baseline;gap:14px;padding:16px 0 8px}
.oinc .t{font-size:11px;font-weight:700;letter-spacing:.15em;text-transform:uppercase;color:#7BA08F}
.oinc .n{font-family:'IBM Plex Mono',monospace;font-size:24px;font-weight:600;color:#8FBFA9}
.opay{margin-top:16px;background:var(--bronze2);color:var(--ink);border-radius:14px;padding:16px 20px;display:flex;justify-content:space-between;align-items:center;gap:14px}
.opay .t{font-size:11px;font-weight:700;letter-spacing:.13em;text-transform:uppercase}
.opay .n{font-family:'IBM Plex Mono',monospace;font-size:23px;font-weight:600;white-space:nowrap}
.opay .t{max-width:11ch}
.onote{font-family:'Cormorant Garamond',Cambria,serif;font-style:italic;font-size:16px;color:#9C968A;margin-top:24px;max-width:860px}
.note{background:#EDF1EE;border-radius:16px;padding:20px 22px;margin-top:20px;font-size:14px;color:var(--ink2)}
.note b{color:var(--green)}
.foc{display:grid;grid-template-columns:1fr;gap:16px;margin-top:32px}
@media(min-width:820px){.foc{grid-template-columns:1fr 1fr;gap:20px}}
.fcard{border:1px solid rgba(217,212,201,.16);border-radius:18px;padding:24px 26px}
.fcard .fn{font-family:'Cormorant Garamond',Cambria,serif;font-size:28px;font-weight:700;color:var(--bronze2)}
.fcard h3{font-family:'Cormorant Garamond',Cambria,serif;font-size:22px;margin:8px 0 8px}
.fcard p{font-size:14px;color:#A8A296;line-height:1.6}
footer{background:var(--ink);color:#6E6960;font-family:'IBM Plex Mono',monospace;font-size:11.5px;padding:28px 0;border-top:1px solid rgba(217,212,201,.12)}
footer .wrap{display:flex;justify-content:space-between;flex-wrap:wrap;gap:10px}
"""

JS = """
const reduced=matchMedia('(prefers-reduced-motion: reduce)').matches;
function countUp(el){
  const t=parseFloat(el.dataset.cnt),d=parseInt(el.dataset.dec||0),s=el.dataset.suf||'',p=el.dataset.pre||'';
  const fmt=v=>p+v.toFixed(d).replace('.',',').replace(/\\B(?=(\\d{3})+(?!\\d))/g,'\\u00a0')+s;
  if(reduced){el.textContent=fmt(t);return}
  const t0=performance.now();
  (function tick(now){const k=Math.min((now-t0)/1300,1),e=1-Math.pow(1-k,3);
    el.textContent=fmt(t*e); if(k<1)requestAnimationFrame(tick)})(t0);
}
const io=new IntersectionObserver(es=>{es.forEach(e=>{
  if(!e.isIntersecting)return; const el=e.target; el.classList.add('on');
  el.querySelectorAll('.bar').forEach(b=>b.style.height=b.dataset.h+'%');
  el.querySelectorAll('.dail div').forEach(b=>b.style.height=b.dataset.h+'%');
  el.querySelectorAll('[data-w]').forEach(b=>b.style.width=b.dataset.w+'%');
  el.querySelectorAll('.dseg').forEach(s=>s.style.strokeDasharray=s.getAttribute('stroke-dasharray'));
  el.querySelectorAll('[data-cnt]').forEach(c=>{if(!c.dataset.done){c.dataset.done=1;countUp(c)}});
  io.unobserve(el);});},{threshold:.2});
document.querySelectorAll('.rv,.chart,.stackbar,.hero .strip,#wfC').forEach(el=>io.observe(el));
// тап по столбцу дня — подпись, сама прячется
let tapT=null;
document.querySelectorAll('.dail div').forEach(b=>b.addEventListener('click',()=>{
  document.querySelectorAll('.dail div.act').forEach(x=>x.classList.remove('act'));
  b.classList.add('act'); clearTimeout(tapT); tapT=setTimeout(()=>b.classList.remove('act'),2200);
}));
const nav=document.getElementById('nav');
if(nav){const nio=new IntersectionObserver(es=>{es.forEach(e=>{if(e.isIntersecting){
  nav.querySelectorAll('a').forEach(a=>a.classList.toggle('act',a.getAttribute('href')==='#'+e.target.id));}});},{threshold:.4});
  document.querySelectorAll('section[id],header[id]').forEach(s=>nio.observe(s));}
"""


def render(r, narrative=None):
    n = narrative or {}
    meta, rev, tr, ex, f = r["meta"], r["revenue"], r["trend"], r["expenses"], r["founders"]
    cf, man = r["cashflow"], r["manual"]
    period = meta["period"]
    import calendar
    last_day = f"{calendar.monthrange(meta['year'], meta['month'])[1]:02d}.{meta['month']:02d}"
    P = []
    A = P.append

    secs = [("s0", "Титул"), ("s1", "Резюме"), ("s2", "Показатели"), ("s3", "Динамика"),
            ("s4", "Структура"), ("s5", "По дням"), ("s6", "Расходы"), ("s7", "Экономика"),
            ("s8", "Учредители"), ("s9", "Деньги"), ("s10", "Фокус")]

    A(f"""<!DOCTYPE html><html lang="ru"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{escape(meta['brand'].upper())} · Отчёт учредителям · {escape(period)}</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,500;0,600;0,700;1,500;1,600&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap" rel="stylesheet">
<style>{CSS}</style></head><body>
<nav id="nav">""" + "".join(f'<a href="#{i}" data-t="{escape(t)}"></a>' for i, t in secs) + "</nav>")

    # ---------------------------------------------------------------- титул
    A(f"""<header class="hero" id="s0"><div class="wrap">
<div class="brand">{' '.join(meta['brand'].upper())}</div>
<div class="tag">клиника пластической хирургии и эстетической косметологии</div>
<h1>Отчёт учредителям</h1><div class="yr">{escape(period)}</div>
<div class="strip">
<div class="it"><div class="v mono" data-cnt="{rev['total']/1e6:.2f}" data-dec="2">0</div><div class="l">млн ₽ выручка</div></div>
<div class="it"><div class="v mono" style="color:{'#8FBFA9' if r['founders_total_income'] >= 0 else '#D98A82'}" data-cnt="{r['founders_total_income']/1e6:.2f}" data-dec="2">0</div><div class="l">млн ₽ доход учредителей</div></div>
<div class="it"><div class="v mono" data-cnt="{cf['closing']/1e6:.2f}" data-dec="2">0</div><div class="l">млн ₽ на счёте на {last_day}</div></div>
{f'<div class="it"><div class="v mono" style="color:#B9895A" data-cnt="{tr["ytd_growth"]:.1f}" data-dec="1" data-suf="%">0</div><div class="l">рост с начала года</div></div>' if tr.get("ytd_growth") is not None else f'<div class="it"><div class="v mono" style="color:#B9895A" data-cnt="{sum(v for v in tr["cur"][:meta["month"]] if v)/1e6:.1f}" data-dec="1">0</div><div class="l">млн ₽ с начала года</div></div>'}
</div></div></header>""")

    # ---------------------------------------------------------------- резюме
    A(f"""<section id="s1"><div class="wrap"><div class="kick rv">01 · Executive Summary</div>
<h2 class="rv d1">Резюме месяца</h2><div class="sub rv d1">Пять тезисов, которые описывают {escape(meta['month_ru'].lower())} полностью</div>""")
    for i, t in enumerate(_summary(r, n), start=1):
        A(f'<div class="tez rv"><div class="n">{i:02d}</div><h3>{t["h"]}</h3><p>{t["p"]}</p></div>')
    A("</div></section>")

    # ---------------------------------------------------------------- KPI
    kpis = [
        (mln(rev["total"]), "млн ₽", "Выручка за месяц",
         (arrow(tr["yoy_month"]) + f" к {meta['month_ru'].lower()}ю {meta['year']-1}") if tr.get("yoy_month") is not None
         else "нет данных за прошлый год", "up" if (tr.get("yoy_month") or 0) >= 0 else "neg"),
        (mln(sum(v for v in tr["cur"][:meta["month"]] if v)), "млн ₽", "Выручка с начала года",
         (arrow(tr["ytd_growth"]) + " к прошлому году") if tr.get("ytd_growth") is not None
         else "сравнение неполное: нет августа 2025",
         "up" if (tr.get("ytd_growth") or 0) >= 0 else "neg"),
        (mln(rev.get("avg_day")), "млн ₽", "Средняя выручка в день",
         f"пик {rev['best_day']['day']} — {mln(rev['best_day']['total'])} млн", ""),
        (mln(ex["acc60_total"]), "млн ₽", "Оплаты поставщикам · сч. 60",
         f"{pct(ex['share_of_revenue'],1,sign=False)} от выручки", ""),
        (mln(r["payroll"]["salary"] + r["payroll"]["taxes"]), "млн ₽", "ФОТ с налогами по банку",
         f"ЗП {mln(r['payroll']['salary'])} + налоги {mln(r['payroll']['taxes'])}", ""),
        (mln(cf["closing"]), "млн ₽", "Остаток на р/с на конец месяца",
         f"{arrow((cf['delta']/cf['opening']*100) if cf['opening'] else None)} — {mln(cf['delta'])} млн за месяц"
         if cf["opening"] else f"{mln(cf['delta'])} млн за месяц",
         "up" if cf["delta"] >= 0 else "neg"),
    ]
    A("""<section id="s2" style="padding-top:0"><div class="wrap"><div class="kick rv">02 · Показатели</div>
<h2 class="rv d1">Ключевые показатели</h2><div class="sub rv d1">Шесть чисел, по которым можно судить о месяце</div><div class="kpi">""")
    for i, (v, u, label, note, cls) in enumerate(kpis):
        A(f'<div class="card rv d{i%3}"><div class="v">{v} <small>{u}</small></div>'
          f'<div class="l">{escape(label)}</div><div class="d {cls}">{note}</div></div>')
    A("</div></div></section>")

    # ---------------------------------------------------------------- динамика
    vals = [v for v in (tr["cur"] + tr["prev"]) if v]
    top = (max(vals) * 1.16) if vals else 1
    bars = []
    for i in range(meta["month"]):
        p, cu = tr["prev"][i], tr["cur"][i]
        acc = "acc" if i == meta["month"] - 1 else "g26"
        bars.append(
            f'<div class="bgrp"><div class="bar g25" data-h="{(p or 0)/top*100:.1f}"></div>'
            f'<div class="bar {acc}" data-h="{(cu or 0)/top*100:.1f}">'
            f'<span class="bv">{mln(cu,1) if cu else ""}</span></div></div>')
    A(f"""<section id="s3" style="padding-top:0"><div class="wrap"><div class="kick rv">03 · Динамика</div>
<h2 class="rv d1">Динамика выручки</h2>
<div class="sub rv d1">Январь – {escape(meta['month_ru'].lower())}: {meta['year']} год против {meta['year']-1}-го, млн ₽</div>
<div class="chart rv"><div class="bars">{''.join(bars)}</div>
<div class="bx">{''.join(f'<span>{MONTHS_RU[i][:3]}</span>' for i in range(meta['month']))}</div>
<div class="leg"><span><span class="dot" style="background:var(--ghost)"></span>{meta['year']-1}</span>
<span><span class="dot" style="background:var(--green)"></span>{meta['year']}</span>
<span><span class="dot" style="background:#B9895A"></span>{meta['year']} · отчётный месяц</span></div></div>""")
    if tr.get("ytd_growth") is not None:
        A(f"""<div class="hl rv d1"><div class="big">{pct(tr['ytd_growth'])}</div>
<p>рост выручки с начала года: {mln(tr['ytd_cur'])} млн против {mln(tr['ytd_prev'])} млн годом ранее.</p></div>""")
    if tr.get("missing_prev"):
        A(f'<div class="note rv"><b>Нет данных:</b> за {", ".join(tr["missing_prev"])} {meta["year"]-1} '
          f'года в источнике нет выручки — сравнение за эти месяцы не строится.</div>')
    A("</div></section>")

    # ---------------------------------------------------------------- структура (donut)
    colors = {"other": "#2E5E4E", "pog": "#4A7A68", "mal": "#7BA08F", "cosm": "#B9895A", "misc": "#D9D4C9"}
    order = sorted(GROUPS, key=lambda g: -rev["by_group"][g])
    parts = [(GROUP_RU[g], rev["by_group"][g], colors[g]) for g in order if rev["by_group"][g] > 0]
    rows = "".join(f'<tr><td><span class="dot" style="background:{colors[g]}"></span>{GROUP_RU[g]}</td>'
                   f'<td class="num">{rub(rev["by_group"][g])}</td>'
                   f'<td class="num">{pct(rev["share"][g],1,sign=False)}</td></tr>' for g in order)
    ch_rows = [("Эквайринг (оплата картами)", rev["card"]), ("Касса (наличные)", rev["cash"]),
               ("Оплата на р/с (физлица)", rev["bank_ind"])]
    A(f"""<section id="s4" style="padding-top:0"><div class="wrap"><div class="kick rv">04 · Выручка</div>
<h2 class="rv d1">Структура выручки</h2>
<div class="sub rv d1">{rub(rev['total'])} ₽ — по направлениям и каналам оплаты</div>
<div class="dsplit rv"><div class="chart">{_donut(parts, rev['total'])}</div>
<div class="tw"><table><tr><th>Направление</th><th class="num">Сумма, ₽</th><th class="num">Доля</th></tr>
{rows}<tr class="tot"><td>Итого выручка</td><td class="num">{rub(rev['total'])}</td><td class="num">100%</td></tr></table></div></div>
<div class="tw rv d1" style="margin-top:18px"><table><tr><th>Канал поступления</th><th class="num">Сумма, ₽</th></tr>
{''.join(f'<tr><td>{escape(t)}</td><td class="num">{rub(v)}</td></tr>' for t,v in ch_rows)}
<tr><td>Возвраты пациентам</td><td class="num" style="color:var(--red)">−{rub(rev['refund'])}</td></tr>
<tr class="tot"><td>Итого</td><td class="num">{rub(rev['total'])}</td></tr></table></div>
<p class="rv" style="font-family:'Cormorant Garamond',serif;font-style:italic;color:var(--mute);margin-top:14px;font-size:16px">
Комиссия эквайринга — {rub(rev.get('acquiring_fee'))} ₽ ({pct(rev.get('acquiring_fee_pct'),2,sign=False)} оборота по картам).</p>
</div></section>""")

    # ---------------------------------------------------------------- по дням
    daily = rev["daily"]
    dmax = (max((d["total"] for d in daily), default=0) or 1) * 1.05
    dbars = "".join(
        f'<div data-h="{d["total"]/dmax*100:.1f}" data-v="{d["day"]} {meta["month_ru"].lower()} · {mln(d["total"])} млн ₽" '
        f'style="opacity:{0.5+0.5*d["total"]/(dmax or 1):.2f}"></div>' for d in daily)
    dx = "".join(f'<span>{d["day"] if (d["day"]%5==0 or d["day"]==1) else ""}</span>' for d in daily)
    A(f"""<section id="s5" style="padding-top:0"><div class="wrap"><div class="kick rv">05 · Выручка</div>
<h2 class="rv d1">Выручка по дням</h2>
<div class="sub rv d1">млн ₽ в день, по данным кассовых отчётов · нажмите на столбец</div>
<div class="chart rv"><div class="dail">{dbars}</div><div class="dx">{dx}</div></div>
<p class="rv" style="font-family:'Cormorant Garamond',serif;font-size:18px;color:var(--ink2);margin-top:16px">
<b style="font-family:'Inter';font-size:15px">Средний день — {mln(rev['avg_day'])} млн ₽.</b>
Сильнейший — {rev['best_day']['day']} {meta['month_ru'].lower()} ({mln(rev['best_day']['total'])} млн),
самый тихий — {rev['worst_day']['day']} {meta['month_ru'].lower()} ({mln(rev['worst_day']['total'])} млн).</p>
</div></section>""")

    # ---------------------------------------------------------------- расходы
    med_share = (ex["medical_total"] / ex["acc60_total"] * 100) if ex["acc60_total"] else 0
    def top_rows(items, limit=7):
        s = sorted(items, key=lambda x: -x["amount"])
        head, tail = s[:limit], s[limit:]
        out = "".join(f'<tr><td>{escape(i["name"])}</td><td class="num">{mln(i["amount"])}</td></tr>' for i in head)
        if tail:
            out += f'<tr><td>Прочие статьи ({len(tail)})</td><td class="num">{mln(sum(i["amount"] for i in tail))}</td></tr>'
        return out
    A(f"""<section id="s6" style="padding-top:0"><div class="wrap"><div class="kick rv">06 · Расходы</div>
<h2 class="rv d1">Расходы: оплаты поставщикам</h2>
<div class="sub rv d1">{rub(ex['acc60_total'])} ₽ по счёту 60 — {pct(ex['share_of_revenue'],1,sign=False)} выручки месяца</div>
<div class="stackbar rv"><div style="background:#2E5E4E" data-w="{med_share:.0f}">Медицинские + ИП · {med_share:.0f}%</div>
<div style="background:#8A7350" data-w="{100-med_share:.0f}">Управленческие · {100-med_share:.0f}%</div></div>
<div class="split"><div class="tw"><table><tr><th>Медицинские расходы + ИП</th><th class="num">млн ₽</th></tr>
{top_rows(ex['medical'])}<tr class="tot"><td>Итого</td><td class="num">{mln(ex['medical_total'])}</td></tr></table></div>
<div class="tw"><table><tr><th>Управленческие расходы</th><th class="num">млн ₽</th></tr>
{top_rows(ex['management'])}<tr class="tot"><td>Итого</td><td class="num">{mln(ex['management_total'])}</td></tr></table></div></div>
</div></section>""")

    # ---------------------------------------------------------------- водопад
    wf = r["waterfall"]
    K = 100.0 / rev["total"] if rev["total"] else 0
    cols, labels = [], []
    for i, s in enumerate(wf):
        color = {"pos": "var(--green)", "inv": "var(--bronze2)"}.get(s["kind"], "var(--mute)")
        lbl = ("+" if s["kind"] == "pos" and i else "") + mln(s["value"])
        top_pct = max(s["to"], s["from"]) * K
        inside = top_pct > 88
        lbl_pos = (f'bottom:{top_pct-9:.2f}%' if inside else f'bottom:calc({top_pct:.2f}% + 6px)')
        if inside:
            color = "#fff" if s["kind"] == "pos" else "var(--ink)"
        bridge = ("" if i == len(wf) - 1 else
                  f'<div class="wbridge" style="bottom:{max(s["to"],s["from"])*K:.2f}%;transition-delay:{i*0.14+0.4:.2f}s"></div>')
        cols.append(
            f'<div class="wcol"><div class="wbar {s["kind"]}" style="bottom:{min(s["from"],s["to"])*K:.2f}%;'
            f'height:{max(abs(s["to"]-s["from"])*K,1.0):.2f}%;transition-delay:{i*0.14:.2f}s"></div>'
            f'<div class="wv" style="{lbl_pos};color:{color};'
            f'transition-delay:{i*0.14+0.3:.2f}s">{lbl}</div>{bridge}</div>')
        labels.append(f'<span>{escape(s.get("short") or s["name"])}</span>')
    kop = {s["name"]: abs(s["value"]) / rev["total"] * 100 for s in wf[1:-1]} if rev["total"] else {}
    margin = r["founders_total_income"] / rev["total"] * 100 if rev["total"] else 0
    A(f"""<section id="s7" style="padding-top:0"><div class="wrap"><div class="kick rv">07 · Экономика</div>
<h2 class="rv d1">Куда уходит выручка</h2>
<div class="sub rv d1">От {mln(rev['total'])} млн ₽ выручки — к доходу месяца</div>
<div class="chart rv" id="wfC"><div class="wfscroll"><div class="wfinner"><div class="wf">{''.join(cols)}</div><div class="wl">{''.join(labels)}</div></div></div></div>
<div class="kop rv d1"><p>Из каждого рубля выручки <b>{kop.get('Гонорары ИП хирургов',0):.0f} копеек</b> забирают гонорары хирургов,
<b>{kop.get('ФОТ, налоги, алименты',0):.0f}</b> — персонал и налоги,
<b>{kop.get('Медицинские расходы',0):.0f}</b> — медицинские закупки,
<b>{kop.get('Управленческие расходы',0)+kop.get('Ремонт нового корпуса',0):.0f}</b> — управление и ремонт.
Учредителям остаётся <b>{margin:.0f} копеек</b> — маржа месяца {pct(margin,1,sign=False)}.</p></div>
</div></section>""")

    # ---------------------------------------------------------------- учредители
    A(f"""<section class="dark" id="s8"><div class="wrap"><div class="kick rv">08 · Учредители</div>
<h2 class="rv d1">Доход учредителей</h2>
<div class="sub rv d1">Методика 50/50: личная выручка + ½ общих доходов − доля расходов</div><div class="own">""")
    for i, key in enumerate(FOUNDERS):
        d = f[key]
        A(f"""<div class="ocard rv d{i}"><h3>{escape(d['name'])}</h3>
<div class="orow"><span>Личная выручка (операции)</span><span class="n">{mln(d['own_revenue'])}</span></div>
<div class="orow"><span>½ выручки других хирургов</span><span class="n">+{mln(d['half_other'])}</span></div>
<div class="orow"><span>½ косметологии и прочего</span><span class="n">+{mln(d['half_cosm']+d['half_misc'])}</span></div>
<div class="orow"><span>Доля расходов</span><span class="n neg">−{mln(d['costs'])}</span></div>
<div class="oinc"><span class="t">Доход за месяц</span><span class="n" style="color:{'#8FBFA9' if d['income'] >= 0 else '#D98A82'}">{mln(d['income'])} млн ₽</span></div>
<div class="orow"><span>Остаток с прошлого месяца</span><span class="n">+{mln(d['prev_balance'])}</span></div>
{f'<div class="orow"><span>Выплачены дивиденды</span><span class="n neg">−{mln(d["dividends"])}</span></div>' if d.get('dividends') else ''}
{f'<div class="orow"><span>Отложено на аренду (депозит)</span><span class="n neg">−{mln(d["reserve"])}</span></div>' if d.get('reserve') else ''}
{f'<div class="orow"><span>Возврат остатка резерва</span><span class="n">+{mln(d["reserve_return"])}</span></div>' if d.get('reserve_return') else ''}
<div class="opay"><span class="t">К выплате на {last_day}</span><span class="n">{mln(d['payout'])} млн ₽</span></div></div>""")
    A(f"""</div><p class="onote rv">Суммы в млн ₽. Доля расходов: ½ всех расходов месяца плюс
личные анализы учредителя. Полная расшифровка каждой строки — в Excel-приложении, лист «07 Учредители».</p>
</div></section>""")

    # ---------------------------------------------------------------- ДДС
    flows = cf["flows"]
    ins = sorted([x for x in flows if x.get("debit")], key=lambda x: -x["debit"])[:6]
    outs = sorted([x for x in flows if x.get("credit")], key=lambda x: -x["credit"])[:6]
    A(f"""<section id="s9"><div class="wrap"><div class="kick rv">09 · Деньги</div>
<h2 class="rv d1">Движение денежных средств</h2>
<div class="sub rv d1">Расчётный счёт, сч. 51 — {escape(period.lower())}</div>
<div class="kpi" style="margin-bottom:20px">
<div class="card rv"><div class="v">{mln(cf['opening'])} <small>млн</small></div><div class="l">Остаток на начало</div></div>
<div class="card rv d1"><div class="v up">+{mln(cf['in_total'])} <small>млн</small></div><div class="l">Поступления</div></div>
<div class="card rv d2"><div class="v neg">−{mln(cf['out_total'])} <small>млн</small></div><div class="l">Списания</div></div>
<div class="card rv d3"><div class="v">{mln(cf['closing'])} <small>млн</small></div><div class="l">Остаток на конец</div></div></div>
<div class="split">
<div class="tw"><table><tr><th style="color:var(--green)">Ключевые поступления</th><th class="num">млн ₽</th></tr>
{''.join(f'<tr><td>{escape(x["note"] or x["label"])}</td><td class="num">{mln(x["debit"])}</td></tr>' for x in ins)}</table></div>
<div class="tw"><table><tr><th style="color:var(--red)">Ключевые списания</th><th class="num">млн ₽</th></tr>
{''.join(f'<tr><td>{escape(x["note"] or x["label"])}</td><td class="num">{mln(x["credit"])}</td></tr>' for x in outs)}</table></div></div>
<div class="note rv"><b>Справочно:</b> проценты по депозитам за месяц — {rub(man.get('deposit_interest'))} ₽;
зарезервировано на аренду — {rub(man.get('reserve_total'))} ₽;
долг по кредитной линии — {rub(man.get('credit_debt'))} ₽;
остаток в кассе на конец месяца — {rub(man.get('cash_on_hand'))} ₽.
<span style="color:var(--mute)">Гонорары ИП хирургов входят в оплаты поставщикам по сч. 60 и отдельной строкой не суммируются.</span></div>
</div></section>""")

    # ---------------------------------------------------------------- фокус
    focus = n.get("focus") or [
        {"h": "Сверка после перехода на новую управленку",
         "p": "Закрыть расхождения по стыковым дням и подтвердить, что ни один платёж не посчитан дважды."},
        {"h": "Контроль сметы нового корпуса",
         "p": f"Ремонт финансируется кредитной линией; долг — {rub(man.get('credit_debt'))} ₽. Держать график ввода."},
        {"h": "Депозитная дисциплина",
         "p": f"Overnight-размещение свободных остатков дало {rub(man.get('deposit_interest'))} ₽ за месяц — сохранить режим."},
        {"h": "Планирование следующего месяца",
         "p": "Сверить план по выручке с фактической загрузкой операционных дней."},
    ]
    A("""<section class="dark" id="s10"><div class="wrap"><div class="kick rv">10 · Дальше</div>
<h2 class="rv d1">Точки внимания и фокус следующего месяца</h2><div class="foc">""")
    for i, item in enumerate(focus, start=1):
        A(f'<div class="fcard rv d{i%3}"><div class="fn">{i:02d}</div><h3>{item["h"]}</h3><p>{item["p"]}</p></div>')
    A("</div></div></section>")

    A(f"""<footer><div class="wrap"><span>{escape(meta['company'])} · Отчёт за {escape(period.lower())}</span>
<span>Источники: анализ сч. 51 · отчёт по выручке · обороты сч. 60</span></div></footer>
<script>{JS}</script></body></html>""")
    return "".join(P)
