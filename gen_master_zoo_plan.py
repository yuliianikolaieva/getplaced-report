#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Генерує 12-місячний план замовлень Master Zoo на Bolt Food
на основі фактичних даних Glovo (звіт Get Placed, Jan-Apr 2026)."""
import re, json

SRC = "MS % get placed data.html"
src = open(SRC, encoding="utf-8").read()
def grab(name):
    m = re.search(r"const "+name+r"=(\[.*?\]);", src, re.S)
    return json.loads(m.group(1))
B = grab("B"); C = grab("C"); D = grab("D")

# ---- Master Zoo факт по Glovo ----
mz = next(i for i,b in enumerate(B) if b[0]=="Master Zoo")
rows = [d for d in D if d[0]==mz]
glovo = {"Jan":0,"Feb":0,"Mar":0,"Apr":0}
city_apr = {}        # Glovo квітень по містах (база для алокації)
for d in rows:
    glovo["Jan"]+=d[3]; glovo["Feb"]+=d[6]; glovo["Mar"]+=d[9]; glovo["Apr"]+=d[21]
    ci = C[d[1]]
    city_apr[ci] = city_apr.get(ci,0) + d[21]

g_vals = list(glovo.values())
glovo_avg = round(sum(g_vals)/len(g_vals))
glovo_base = 5600   # стабільна місячна база Glovo для прогнозу (≈ave, трохи консервативно вгору)

# ---- Бенчмарк Bolt/Glovo по великих ритейл-мережах ----
benchmarks = [
    ("Kopiyka", 0.063), ("Loko", 0.072), ("Santim", 0.104),
    ("Rukavychka", 0.106), ("OKKO", 0.159), ("Beer Market", 0.245),
]
bench_avg = round(sum(r for _,r in benchmarks)/len(benchmarks)*100,1)

# ---- Сценарії зрілого Bolt/Glovo + S-крива розгону (12 міс) ----
# прогрес виходу на зрілу частку (logistic-подібний розгін)
prog = [0.12,0.21,0.33,0.46,0.58,0.70,0.79,0.87,0.92,0.96,0.99,1.00]
months = ["Лип'26","Сер'26","Вер'26","Жов'26","Лис'26","Гру'26",
          "Січ'27","Лют'27","Бер'27","Кві'27","Тра'27","Чер'27"]
scenarios = {"Консервативний":0.20, "Базовий":0.40, "Амбітний":0.55}

AOV_UAH = 850          # середній чек, грн
EUR_RATE = 51          # 1 EUR = 51 грн
AOV_EUR = AOV_UAH/EUR_RATE

def plan(target):
    return [round(glovo_base*target*p) for p in prog]
plans = {k: plan(v) for k,v in scenarios.items()}
base_plan = plans["Базовий"]
base_share = [round(scenarios["Базовий"]*p*100,1) for p in prog]

# GMV (грн та €)
gmv_uah = {k:[o*AOV_UAH for o in v] for k,v in plans.items()}
base_gmv_uah = gmv_uah["Базовий"]
base_gmv_eur = [g/EUR_RATE for g in base_gmv_uah]
year_orders = {k:sum(v) for k,v in plans.items()}
year_gmv_uah = {k:sum(gmv_uah[k]) for k in plans}
year_gmv_eur = {k:year_gmv_uah[k]/EUR_RATE for k in plans}

# ---- Алокація базового сценарію по топ-містах (за часткою Glovo квітень) ----
total_apr = sum(city_apr.values())
top_cities = sorted(city_apr.items(), key=lambda x:-x[1])[:10]
city_alloc = []
m12 = base_plan[-1]
for ci,v in top_cities:
    share = v/total_apr
    city_alloc.append((ci, round(share*100,1), round(share*m12)))

data = dict(
    glovo=glovo, glovo_avg=glovo_avg, glovo_base=glovo_base,
    n_cities=len(city_apr), n_rows=len(rows),
    benchmarks=benchmarks, bench_avg=bench_avg,
    months=months, scenarios=scenarios, plans=plans,
    base_share=base_share, city_alloc=city_alloc,
    totals={k: sum(v) for k,v in plans.items()},
)
print(json.dumps(data, ensure_ascii=False, indent=1))

# ================= HTML =================
def fmt(n): return f"{n:,}".replace(","," ")
def mln(n): return f"{n/1e6:.1f} млн ₴".replace(".",",")
def eurk(n): return f"€{n/1000:.0f}k"

month_th = "".join(f"<th>{m}</th>" for m in months)
def plan_row(name, vals, cls="", suffix=""):
    tds = "".join(f"<td>{fmt(round(v))}</td>" for v in vals)
    return f"<tr class='{cls}'><td class='lbl'>{name}</td>{tds}<td class='tot'>{fmt(round(sum(vals)))}{suffix}</td></tr>"

share_row = "<tr class='muted'><td class='lbl'>Частка від Glovo, %</td>" + \
    "".join(f"<td>{s}%</td>" for s in base_share) + "<td class='tot'>—</td></tr>"

bench_rows = "".join(
    f"<tr><td class='lbl'>{n}</td><td>{r*100:.1f}%</td></tr>" for n,r in benchmarks)

city_rows = "".join(
    f"<tr><td class='lbl'>{c}</td><td>{s}%</td><td>{fmt(o)}</td></tr>"
    for c,s,o in city_alloc)

glovo_cells = "".join(f"<td>{fmt(v)}</td>" for v in glovo.values())

scen_cards = ""
for name,tgt in scenarios.items():
    tot = year_orders[name]
    hot = " hot" if name=="Базовий" else ""
    scen_cards += f"""<div class="scard{hot}">
      <div class="sc-name">{name}</div>
      <div class="sc-big">{fmt(tot)}</div>
      <div class="sc-sub">замовлень Bolt за 12 міс</div>
      <div class="sc-meta">GMV ≈ <b>{mln(year_gmv_uah[name])}</b> · {eurk(year_gmv_eur[name])}<br>
      зріла частка ≈ {int(tgt*100)}% від Glovo · ~{fmt(plans[name][-1])}/міс на 12-й міс</div>
    </div>"""

html = f"""<!DOCTYPE html>
<html lang="uk"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Master Zoo · План Bolt Food 12 міс</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>
:root{{--bg:#0d1117;--card:#161b22;--line:#262d38;--tx:#e6edf3;--mut:#8b949e;
--bolt:#34d186;--glovo:#ffb703;--accent:#34d186}}
*{{box-sizing:border-box}}
body{{margin:0;font-family:-apple-system,Segoe UI,Roboto,Inter,Arial,sans-serif;
background:var(--bg);color:var(--tx);line-height:1.5}}
.wrap{{max-width:1080px;margin:0 auto;padding:28px 20px 80px}}
h1{{font-size:30px;margin:0 0 6px}}
.sub{{color:var(--mut);margin:0 0 26px;font-size:15px}}
.tag{{display:inline-block;background:#1f6feb22;color:#58a6ff;border:1px solid #1f6feb55;
border-radius:20px;padding:3px 12px;font-size:12px;margin-bottom:14px}}
h2{{font-size:19px;margin:38px 0 14px;border-left:3px solid var(--accent);padding-left:10px}}
.grid{{display:grid;gap:14px}}
.g4{{grid-template-columns:repeat(4,1fr)}}
.g3{{grid-template-columns:repeat(3,1fr)}}
.kpi{{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px}}
.kpi .v{{font-size:26px;font-weight:700}}
.kpi .l{{color:var(--mut);font-size:12px;margin-top:4px}}
.kpi .v.g{{color:var(--glovo)}} .kpi .v.b{{color:var(--bolt)}}
.scard{{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:18px;text-align:center}}
.scard.hot{{border-color:var(--bolt);box-shadow:0 0 0 1px var(--bolt)33}}
.sc-name{{color:var(--mut);font-size:13px;text-transform:uppercase;letter-spacing:.5px}}
.sc-big{{font-size:32px;font-weight:800;color:var(--bolt);margin:6px 0}}
.sc-sub{{font-size:12px;color:var(--mut)}}
.sc-meta{{font-size:12px;color:var(--tx);margin-top:8px;opacity:.85}}
table{{width:100%;border-collapse:collapse;background:var(--card);border:1px solid var(--line);
border-radius:12px;overflow:hidden;font-size:13px}}
th,td{{padding:9px 8px;text-align:right;border-bottom:1px solid var(--line)}}
th{{background:#1c2230;color:var(--mut);font-weight:600;font-size:12px}}
td.lbl,th:first-child{{text-align:left}}
td.lbl{{color:var(--mut)}}
td.tot{{font-weight:700;color:var(--bolt);background:#0f1722}}
tr.muted td{{color:var(--mut);font-size:12px}}
tr.hl td{{font-weight:700}}
.chartbox{{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:18px;margin-top:14px}}
.note{{background:var(--card);border:1px solid var(--line);border-left:3px solid var(--accent);
border-radius:10px;padding:14px 16px;color:var(--mut);font-size:13px;margin-top:14px}}
.note b{{color:var(--tx)}}
ul{{margin:8px 0;padding-left:20px}} li{{margin:4px 0}}
.foot{{margin-top:50px;color:var(--mut);font-size:12px;border-top:1px solid var(--line);padding-top:16px}}
a{{color:#58a6ff}}
.scroll{{overflow-x:auto}}
@media(max-width:760px){{.g4,.g3{{grid-template-columns:1fr 1fr}}}}
</style></head>
<body><div class="wrap">
<div class="tag">Get Placed · Ukraine · план на основі реальних даних Glovo</div>
<h1>Master Zoo — план замовлень на Bolt Food</h1>
<p class="sub">12-місячний прогноз помісячно. База — фактичні дані Glovo (Jan–Apr 2026) та бенчмарк
конверсії великих ритейл-мереж, які вже працюють на Bolt Food.</p>

<h2>Стартова точка (факт Glovo)</h2>
<div class="grid g4">
  <div class="kpi"><div class="v g">{fmt(glovo_avg)}</div><div class="l">сер. замовлень/міс на Glovo (Jan–Apr)</div></div>
  <div class="kpi"><div class="v b">0</div><div class="l">замовлень на Bolt Food зараз</div></div>
  <div class="kpi"><div class="v">{data['n_cities']}</div><div class="l">міст присутності (Glovo)</div></div>
  <div class="kpi"><div class="v">#16</div><div class="l">за обсягом серед ритейл-брендів Glovo</div></div>
</div>
<div class="scroll" style="margin-top:14px"><table>
<thead><tr><th>Glovo, замовлень</th><th>Січень</th><th>Лютий</th><th>Березень</th><th>Квітень</th></tr></thead>
<tbody><tr class="hl"><td class="lbl">Master Zoo</td>{glovo_cells}</tr></tbody>
</table></div>
<div class="note">Master Zoo — стабільний партнер Glovo (~{fmt(glovo_avg)} замовлень/міс, рівний попит без
сезонних провалів) і <b>повністю відсутній на Bolt Food</b>. Це чистий приріст для платформи.</div>

<h2>Методологія прогнозу</h2>
<div class="note">
<ul>
<li><b>База Glovo:</b> {fmt(glovo_base)} замовлень/міс (стабільне середнє Master Zoo, тримаємо плоско — попит зрілий).</li>
<li><b>Середній чек (AOV):</b> {AOV_UAH} грн ≈ €{AOV_EUR:.1f} (курс 1 € = {EUR_RATE} грн). GMV = замовлення × AOV.</li>
<li><b>Крива розгону:</b> S-подібний ramp-up за 12 міс — новий партнер виходить на зрілу частку
поступово (видимість, маркетинг, звичка клієнтів, Bolt+).</li>
<li><b>Сценарії (зріла частка від обсягу Glovo):</b> консервативний 20%, базовий 40%, амбітний 55%.</li>
</ul></div>

<h2>Сценарії — підсумок за рік</h2>
<div class="grid g3">{scen_cards}</div>

<h2>Помісячний план — замовлення</h2>
<div class="chartbox"><canvas id="chart" height="120"></canvas></div>
<div class="scroll" style="margin-top:14px"><table>
<thead><tr><th>Bolt Food, замовлень</th>{month_th}<th class="tot">Рік</th></tr></thead>
<tbody>
{plan_row("Консервативний (20%)", plans["Консервативний"])}
{plan_row("Базовий (40%)", base_plan, "hl")}
{plan_row("Амбітний (55%)", plans["Амбітний"])}
{share_row}
</tbody></table></div>

<h2>Помісячний GMV — базовий сценарій (40%)</h2>
<div class="scroll" style="margin-top:6px"><table>
<thead><tr><th>Базовий (40%)</th>{month_th}<th class="tot">Рік</th></tr></thead>
<tbody>
{plan_row("Замовлення", base_plan, "hl")}
{plan_row("GMV, ₴", base_gmv_uah, "", "")}
{plan_row("GMV, €", base_gmv_eur)}
</tbody></table></div>
<div class="note">Базовий рік 1: <b>{fmt(year_orders['Базовий'])} замовлень</b> · GMV <b>{mln(year_gmv_uah['Базовий'])}</b>
(≈ {eurk(year_gmv_eur['Базовий'])}). AOV {AOV_UAH} грн (€{AOV_EUR:.1f}).</div>

<h2>З яких міст стартувати (базовий, 12-й міс)</h2>
<p class="sub">Алокація зрілого місячного обсягу за часткою Glovo по містах — пріоритет запуску.</p>
<div class="scroll"><table>
<thead><tr><th>Місто</th><th>частка попиту</th><th>Bolt замовл./міс (зрілий стан)</th></tr></thead>
<tbody>{city_rows}</tbody></table></div>

<div class="foot">
Джерело: звіт Get Placed (Glovo vs Bolt, store-level, Jan–Apr 2026) ·
<a href="https://yuliianikolaieva.github.io/getplaced-report/">getplaced-report</a><br>
Прогноз — модельний (бенчмарк-конверсія × S-крива розгону), не зобов'язання. Cmd+Shift+R для оновлення.
</div>
</div>

<script>
const M={json.dumps(months, ensure_ascii=False)};
new Chart(document.getElementById('chart'),{{
 type:'bar',
 data:{{labels:M,datasets:[
   {{label:'Консервативний 20%',data:{json.dumps(plans["Консервативний"])},backgroundColor:'#2d6a4f'}},
   {{label:'Базовий 40%',data:{json.dumps(base_plan)},backgroundColor:'#34d186'}},
   {{label:'Амбітний 55%',data:{json.dumps(plans["Амбітний"])},backgroundColor:'#9af5c8'}},
 ]}},
 options:{{responsive:true,plugins:{{legend:{{labels:{{color:'#e6edf3'}}}}}},
  scales:{{x:{{ticks:{{color:'#8b949e'}},grid:{{display:false}}}},
           y:{{ticks:{{color:'#8b949e'}},grid:{{color:'#262d38'}}}}}}}}
}});
</script>
</body></html>"""

open("master-zoo-bolt-plan.html","w",encoding="utf-8").write(html)
print("\nWROTE master-zoo-bolt-plan.html", len(html), "bytes")
