import argparse
import json
import sys
from pathlib import Path

import pandas as pd
from scipy import stats


NA_CODE = -99

INCOME_ORDER = [
    "не вистачає навіть на їжу",
    "вистачає на їжу, але купити одяг проблема",
    "вистачає на їжу й одяг, але купити товари тривалого користув",
    "вистачає на все, окрім великих покупок, таких як житло",
    "вистачає на все",
]
INCOME_SHORT = [
    "Не вистачає на їжу",
    "Не вистачає на одяг",
    "Не вистачає на товари тривалого вжитку",
    "Не вистачає на великі покупки",
    "Вистачає на все",
]

SALARY_ORDER = [
    "до 9 000 грн", "9 001 - 15 000 грн", "15 001 - 20 000 грн", "20 001 - 30 000 грн",
    "30 001 - 50 000 грн", "50 001 - 70 000 грн", "70 001 - 100 000 грн", "понад 100 000 грн",
]
SALARY_SHORT = ["<9 тис", "9–15 тис", "15–20 тис", "20–30 тис", "30–50 тис", "50–70 тис", "70–100 тис", ">100 тис"]

SETTLEMENT_MAP = {
    "Село": "Село",
    "Мiсто з населенням до 50 тис": "Місто до 50 тис.",
    "Мiсто з населенням 51-100 тис": "Місто 51–100 тис.",
    "Мiсто з населенням 101-500 тис": "Місто 101–500 тис.",
    "Мiсто з населенням понад 500 тис": "Місто понад 500 тис.",
}
SETTLEMENT_ORDER = ["Село", "Місто до 50 тис.", "Місто 51–100 тис.", "Місто 101–500 тис.", "Місто понад 500 тис."]

AGE_ORDER = ["18-35 років", "36-55 років", "56-65 років"]
REGION_ORDER = ["Захiд", "Центр", "Київ", "Схiд", "Пiвдень", "Пiвнiч"]

DISPLACEMENT_MAP = {
    "проживала завжди": "always",
    "переїхала сюди через війну": "war",
    "переїхала з інших причин": "other",
}
DISPLACEMENT_LABELS = {"always": "Проживала завжди", "war": "Переїхала через війну", "other": "Переїхала з інших причин"}

ORDINAL_3 = {"неважливо": 1, "ні": 1, "бажано": 2, "можливо": 2, "обов'язково": 3}

CONDITIONS = [
    ("v73", "Неповний робочий день"),
    ("v74", "Гнучкий графік"),
    ("v75", "Зручно діставатися (парковка/транспорт)"),
    ("v76", "Дорога до роботи не більше 1 год"),
    ("v77", "Дитсадок/дитяча кімната поруч"),
    ("v78", "Дружній колектив"),
    ("v79", "Можливість працювати віддалено"),
    ("v80", "Навчання та саморозвиток"),
    ("v82", "Кар'єрне зростання"),
]

CHANNELS = [
    ("v98", "Онлайн-агрегатори (work.ua тощо)"),
    ("v99", "Сайти компаній-роботодавців"),
    ("v100", "Родичі"),
    ("v101", "Друзі чи знайомі"),
    ("v102", "Колишні колеги"),
    ("v103", "Державна служба зайнятості (ДСЗ)"),
]

BARRIERS = [
    ("v108", "Недостатній досвід роботи"),
    ("v109", "Немає досвіду роботи взагалі"),
    ("v110", "Застарілі знання чи навички"),
    ("v111", "Бракує потрібних знань чи навичок"),
    ("v112", "Догляд за дітьми чи родичами"),
    ("v113", "Велика конкуренція"),
    ("v114", "Упередженість працедавця"),
    ("v115", "Низький рівень зарплати"),
    ("v116", "Немає віддаленої роботи"),
    ("v117", "Немає гнучкого графіка / неповного дня"),
    ("v118", "Проблеми з транспортом"),
]

DSZ_HELP_MAP = {
    "дуже корисно, ДСЗ допомогла знайти роботу": "very",
    "дещо корисно, запропонували кілька підходящих вакансії": "some",
    "не корисно, не запропонували підходящих вакансій": "none",
    "не шукала роботу через ДСЗ": "not_searched",
}

SECTOR_LABELS = [
    "Сільське господарство, лісівництво, рибальство", "Видобувна промисловість", "Переробна промисловість",
    "Будівництво чи ремонт житла", "Торгівля", "Особисті послуги (перукар тощо)", "Транспорт, перевезення",
    "Готель чи ресторан", "Медіа", "Інформаційні технології", "Фінансовий сектор",
    "Професійні послуги (юрист, бухгалтер...)", "Освіта", "Медицина", "Соціальна допомога",
    "Правоохоронна система", "Армія", "Оборонно-промисловий комплекс", "Держслужба або ОМС", "Мистецтво чи спорт",
]

ATYP_MAP_PREFIX = {"так": "yes", "ні": "no", "важко сказати": "unsure"}



def load_data(path: str) -> pd.DataFrame:
    return pd.read_excel(path, sheet_name=0)


def na_if_skip(value):
    if value == NA_CODE:
        return None
    return value


def map_ordinal(value):
    v = na_if_skip(value)
    if v is None:
        return None
    return ORDINAL_3.get(v)


def map_barrier(value):
    v = na_if_skip(value)
    if v is None or v == "не стосується":
        return None
    if isinstance(v, str):
        return int(v.strip()[0])
    return int(v)


def has_young_children(row) -> bool:
    total = 0
    for col in ("v139", "v140"):
        v = row.get(col)
        if v is not None and v != NA_CODE and v > 0:
            total += v
    return total > 0


# СТАТИСТИЧНА ПЕРЕВІРКА ТРЬОХ ОСНОВНИХ СПОСТЕРЕЖЕНЬ

def run_significance_checks(df: pd.DataFrame) -> dict:
    results = {}

    sub = df[df["v150"].isin(INCOME_ORDER) & df["v147"].isin(["так", "ні"])].copy()
    income_rank = {v: i + 1 for i, v in enumerate(INCOME_ORDER)}
    sub["income_num"] = sub["v150"].map(income_rank)
    a = sub.loc[sub["v147"] == "так", "income_num"]
    b = sub.loc[sub["v147"] == "ні", "income_num"]
    u, p = stats.mannwhitneyu(a, b, alternative="two-sided")
    results["idp_income"] = {
        "n_idp": int(len(a)), "n_non_idp": int(len(b)),
        "mean_idp": round(float(a.mean()), 2), "mean_non_idp": round(float(b.mean()), 2),
        "p_value": float(p),
    }

    work = df[df["v55"] == "так"].copy()
    work["has_young_kids"] = work.apply(has_young_children, axis=1)
    cond_tests = {}
    for code, label in CONDITIONS:
        work[f"{code}_num"] = work[code].apply(map_ordinal)
        a = work.loc[work["has_young_kids"], f"{code}_num"].dropna()
        b = work.loc[~work["has_young_kids"], f"{code}_num"].dropna()
        u, p = stats.mannwhitneyu(a, b, alternative="two-sided")
        cond_tests[code] = {
            "label": label, "mean_kids": round(float(a.mean()), 2),
            "mean_no_kids": round(float(b.mean()), 2), "p_value": float(p),
        }
    results["kids_conditions"] = cond_tests

    for code, _ in CHANNELS:
        work[f"{code}_num"] = work[code].apply(map_ordinal)
    paired = work[["v98_num", "v103_num"]].dropna()
    w, p = stats.wilcoxon(paired["v98_num"], paired["v103_num"])
    help_vals = df["v94"]
    searched = help_vals[~help_vals.isin([NA_CODE, "не шукала роботу через ДСЗ"])]
    results["channels"] = {
        "n": int(len(paired)),
        "mean_aggregators": round(float(paired["v98_num"].mean()), 2),
        "mean_dsz": round(float(paired["v103_num"].mean()), 2),
        "p_value": float(p),
        "n_searched_via_dsz": int(len(searched)),
        "share_not_useful": round(float((searched == "не корисно, не запропонували підходящих вакансій").mean()), 3),
    }

    return results



def build_records(df: pd.DataFrame) -> list:
    records = []
    for _, row in df.iterrows():
        rec = {
            "region": row["v39"],
            "settlement": SETTLEMENT_MAP.get(row["v40"], row["v40"]),
            "age": row["v28"],
            "disp": DISPLACEMENT_MAP.get(row["v145"], None),
            "idp": {"так": 1, "ні": 0}.get(row["v147"]),
            "inc": (INCOME_ORDER.index(row["v150"]) + 1) if row["v150"] in INCOME_ORDER else None,
            "sal": (SALARY_ORDER.index(row["v70"]) + 1) if row["v70"] in SALARY_ORDER else None,
            "pw": 1 if row["v55"] == "так" else 0,
            "kids": 1 if has_young_children(row) else 0,
        }

        cond = {code: map_ordinal(row[code]) for code, _ in CONDITIONS}
        rec["cond"] = cond

        chan = {code: map_ordinal(row[code]) for code, _ in CHANNELS}
        rec["chan"] = chan

        barr = {code: map_barrier(row[code]) for code, _ in BARRIERS}
        rec["barr"] = barr

        dsz_raw = na_if_skip(row["v94"])
        rec["dsz_help"] = DSZ_HELP_MAP.get(dsz_raw) if dsz_raw else None

        atyp_raw = na_if_skip(row["v126"])
        rec["atyp"] = None
        if atyp_raw:
            for prefix, code in ATYP_MAP_PREFIX.items():
                if str(atyp_raw).startswith(prefix):
                    rec["atyp"] = code
                    break

        sectors = []
        for i in range(1, 21):
            col = f"v61_{i:02d}"
            if row.get(col) == 1:
                sectors.append(i)
        rec["sec"] = sectors

        records.append(rec)
    return records


# ГЕНЕРАЦІЯ HTML-ДАШБОРДУ

def build_html(records: list, meta: dict, stats_results: dict) -> str:
    payload = {
        "records": records,
        "conditions": CONDITIONS,
        "channels": CHANNELS,
        "barriers": BARRIERS,
        "sectorLabels": SECTOR_LABELS,
        "incomeLabels": INCOME_SHORT,
        "salaryLabels": SALARY_SHORT,
        "settlementOrder": SETTLEMENT_ORDER,
        "ageOrder": AGE_ORDER,
        "regionOrder": REGION_ORDER,
        "displacementLabels": DISPLACEMENT_LABELS,
        "meta": meta,
        "stats": stats_results,
    }
    data_json = json.dumps(payload, ensure_ascii=False)
    return HTML_TEMPLATE.replace("__DATA_JSON__", data_json)


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="uk">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Жінки на ринку праці — візуалізація опитування</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
  :root{
    --paper:#F1EEE4; --paper-raised:#FAF8F2; --ink:#20261F; --ink-soft:#5C6355;
    --line:#D9D2BE; --green:#2F5233; --green-soft:#7FA37A; --amber:#C89B3C;
    --steel:#3B6E8F; --rust:#B5502E; --sand:#D9CBA3; --max-w:1280px;
  }
  *{box-sizing:border-box;}
  body{ margin:0; background:var(--paper); color:var(--ink); font-family:'Inter',system-ui,-apple-system,sans-serif; line-height:1.45; }
  h1,h2,.serif{ font-family:'Fraunces',Georgia,serif; }
  .wrap{ max-width:var(--max-w); margin:0 auto; }

  header{ background:var(--ink); color:var(--paper); padding:40px 24px 28px; text-align:center; }
  header .kicker{ font-size:13px; color:var(--sand); margin:0 0 10px; }
  header h1{ font-size:clamp(24px,3.4vw,34px); font-weight:500; line-height:1.15; margin:0 auto; max-width:34ch; }
  header p{ font-size:14.5px; color:#C9C4B4; margin:0 auto; max-width:70ch; }

  .filters{ background:var(--paper-raised); border-bottom:1px solid var(--line); position:sticky; top:0; z-index:10; }
  .filters .wrap{ padding:12px 24px; display:flex; flex-wrap:wrap; align-items:flex-end; gap:16px 22px; }
  .field{ display:flex; flex-direction:column; gap:3px; }
  .field label{ font-size:11.5px; color:var(--ink-soft); }
  .field select{
    font-family:inherit; font-size:13.5px; color:var(--ink); background:transparent; border:none;
    border-bottom:1.5px solid var(--line); padding:3px 20px 5px 2px; min-width:140px; cursor:pointer; appearance:none;
    background-image:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='10' height='7'><path d='M0 0L5 6L10 0' fill='none' stroke='%235C6355' stroke-width='1.4'/></svg>");
    background-repeat:no-repeat; background-position:right 4px center;
  }
  .field select:focus{ outline:none; border-bottom-color:var(--green); }
  #resetBtn{
    font-family:inherit; font-size:12.5px; color:var(--ink-soft); background:none; border:1px solid var(--line);
    border-radius:18px; padding:6px 14px; cursor:pointer; margin-left:auto;
  }
  #resetBtn:hover{ border-color:var(--ink-soft); color:var(--ink); }
  #resetBtn:focus-visible, select:focus-visible{ outline:2px solid var(--steel); outline-offset:2px; }
  .n-badge{ font-size:12.5px; color:var(--ink-soft); white-space:nowrap; }
  .n-badge b{ color:var(--ink); }

  main{ padding:22px 24px 8px; }
  .section-label{ max-width:var(--max-w); margin:0 auto; padding:18px 0 4px; font-size:12.5px; color:var(--ink-soft); display:flex; align-items:center; gap:10px; }
  .section-label::after{ content:''; flex:1; height:1px; background:var(--line); }

  .grid{ max-width:var(--max-w); margin:0 auto; display:grid; grid-template-columns:repeat(2,1fr); gap:18px; padding-bottom:18px; }
  .grid.wide{ grid-template-columns:1fr; }
  .card{ background:var(--paper-raised); border:1px solid var(--line); border-radius:6px; padding:18px 20px 14px; display:flex; flex-direction:column; justify-content:center; }
  .card-head{ display:flex; align-items:baseline; justify-content:space-between; gap:10px; margin-bottom:2px; }
  .card h2{ font-size:15.5px; font-weight:600; margin:0; font-family:'Inter',sans-serif; line-height:1.3; }
  .badge{ font-size:10.5px; color:var(--green); background:#E3EBE0; border-radius:10px; padding:2px 8px; white-space:nowrap; flex:none; }
  .chart-box{ margin-top:8px; }
  .chart-box svg{ width:100%; height:auto; display:block; }
  .chart-empty{ font-size:13px; color:var(--ink-soft); padding:30px 8px; text-align:center; }
  .caption{ font-size:13px; color:var(--ink-soft); margin:8px 0 0; }
  .caption b{ color:var(--ink); }
  .legend{ display:flex; gap:14px; flex-wrap:wrap; font-size:11.5px; color:var(--ink-soft); margin-top:6px; }
  .legend .dot{ display:inline-block; width:8px; height:8px; border-radius:50%; margin-right:4px; }

  @media (max-width:820px){ .grid{ grid-template-columns:1fr; } }

  footer{ background:var(--ink); color:#B9B4A4; padding:30px 24px 40px; margin-top:14px; }
  footer .wrap h3{ color:var(--paper); font-weight:500; font-size:15px; margin:0 0 10px; font-family:'Inter',sans-serif; }
  footer ul{ padding-left:16px; margin:0; font-size:12.5px; line-height:1.65; columns:2; column-gap:32px; }
  footer li{ break-inside:avoid; margin-bottom:6px; }
  footer p.small{ font-size:11px; color:#7D7969; margin-top:18px; columns:1; }
</style>
</head>
<body>

<header>
  <div class="wrap">
    <p class="kicker">Мікродані опитування · CAWI, N=671</p>
    <h1>Жінки на ринку праці: візуальний аналіз опитування</h1>
  </div>
</header>

<div class="filters">
  <div class="wrap">
    <div class="field"><label for="fRegion">Макрорегіон</label><select id="fRegion"></select></div>
    <div class="field"><label for="fAge">Вікова група</label><select id="fAge"></select></div>
    <div class="field"><label for="fSettlement">Тип поселення</label><select id="fSettlement"></select></div>
    <div class="field"><label for="fDisp">Статус переміщення</label><select id="fDisp"></select></div>
    <button id="resetBtn" type="button">Скинути фільтри</button>
    <span class="n-badge">Показано: <b id="nShown">671</b> з <b id="metaTotal">671</b></span>
  </div>
</div>

<main>

  <div class="grid">

    <div class="card">
      <div class="card-head"><h2>Статус ВПО й фінансова спроможність</h2><span class="badge" id="badge1">p&nbsp;=&nbsp;…</span></div>
      <div class="chart-box"><div id="chart1"></div></div>
      <p class="caption" id="cap1"></p>
    </div>

    <div class="card">
      <div class="card-head"><h2>Що насправді важливо матерям малих дітей</h2><span class="badge" id="badge2">p&nbsp;=&nbsp;…</span></div>
      <div class="chart-box"><div id="chart2"></div></div>
      <p class="caption" id="cap2"></p>
      <div class="legend">
        <span><span class="dot" style="background:var(--amber)"></span>Є діти до 7 років</span>
        <span><span class="dot" style="background:var(--steel)"></span>Немає дітей до 7 років</span>
      </div>
    </div>

    <div class="card" style="grid-column:1 / -1;">
      <div class="card-head"><h2>Кому насправді довіряють у пошуку роботи</h2><span class="badge" id="badge3">p&nbsp;=&nbsp;…</span></div>
      <div class="chart-box"><div id="chart3"></div></div>
      <p class="caption" id="cap3"></p>
    </div>

  </div>

  <p class="section-label">Ширший портрет вибірки</p>
  <div class="grid">

    <div class="card" style="grid-column:1 / -1;">
      <div class="card-head"><h2>Бар'єри пошуку роботи за віковою групою</h2></div>
      <div class="chart-box"><div id="chart4"></div></div>
      <p class="caption">Середня оцінка серйозності бар'єра (1–5) для кожної вікової групи. Колір масштабується відносно показаних даних.</p>
    </div>

    <div class="card">
      <div class="card-head"><h2>Очікувана зарплата vs достаток домогосподарства</h2></div>
      <div class="chart-box"><div id="chart5"></div></div>
      <p class="caption">Розподіл очікуваної зарплати «на руки» в межах кожної категорії поточного матеріального стану.</p>
    </div>

    <div class="card">
      <div class="card-head"><h2>Готовність опановувати «нетипову» професію</h2></div>
      <div class="chart-box"><div id="chart6"></div></div>
      <p class="caption">Частка відповідей на запитання про освоєння професії, нетипової для жінок, за статусом ВПО.</p>
    </div>

    <div class="card">
      <div class="card-head"><h2>Портрет вибірки</h2></div>
      <div class="chart-box"><div id="chart7"></div></div>
      <p class="caption">Розподіл респонденток за віком, макрорегіоном і типом поселення.</p>
    </div>

    <div class="card">
      <div class="card-head"><h2>Топ-10 бажаних секторів працевлаштування</h2></div>
      <div class="chart-box"><div id="chart8"></div></div>
      <p class="caption">Частка тих, хто планує працювати і відзначила сектор серед до трьох бажаних (можна кілька відповідей).</p>
    </div>

  </div>

</main>

<footer>
  <div class="wrap">
    <h3>Методологія та обмеження</h3>
    <ul>
      <li>N=671, CAWI; усі респондентки — жінки, які на момент опитування не в найманій праці приватного/держ./громадського секторів і не на держслужбі (ці статуси завершували інтерв'ю за логікою анкети).</li>
      <li>Код -99 означає «питання не ставилося через маршрутизацію анкети» — виключено з розрахунків, а не прирівняно до нуля.</li>
      <li>«Діти до 7 років» — сума категорій «до 3 років» і «4–6 років», які живуть з респонденткою.</li>
      <li>Порівняння описові/кореляційні: перетинний дизайн, причинних висновків не дає. Значущість — Mann-Whitney U / Wilcoxon (порядкові шкали).</li>
      <li>У вибірці всі 671 респонденток мають однакове значення сімейного стану — ймовірно, артефакт відбору підмножини для тестового, тому змінна не використана в аналізі.</li>
      <li>Ваг для генсукупності немає — частки й середні не є оцінкою на всю популяцію жінок України.</li>
    </ul>
    <p class="small">Код: analysis.py (pandas + scipy.stats). Графіки — SVG, що генерується на льоту з вбудованого JSON, без зовнішніх бібліотек візуалізації.</p>
  </div>
</footer>

<script>
const DATA = __DATA_JSON__;

const COLORS = {
  income: ['#B5502E', '#D68C4B', '#D9CBA3', '#8FAF86', '#2F5233'],
  salarySeq: ['#EAF0F3','#D3E2EA','#B9D2DF','#98BCCE','#749FBA','#4F7FA0','#335F80','#1F3A52'],
  kids: '#C89B3C', noKids: '#3B6E8F', dsz: '#B5502E', chan: '#3B6E8F',
  grid: '#D9D2BE', ink: '#20261F', inkSoft: '#5C6355',
  yes: '#2F5233', unsure: '#C89B3C', no: '#D9CBA3',
  heatLo: [237,241,235], heatHi: [31,58,82]
};

function el(tag, attrs, parent){
  const e = document.createElementNS('http://www.w3.org/2000/svg', tag);
  for(const k in attrs){ e.setAttribute(k, attrs[k]); }
  if(parent) parent.appendChild(e);
  return e;
}
function svgRoot(container, w, h){
  container.innerHTML = '';
  const svg = el('svg', {viewBox:`0 0 ${w} ${h}`, xmlns:'http://www.w3.org/2000/svg'});
  container.appendChild(svg);
  return svg;
}
function text(svg, x, y, str, opts={}){
  const t = el('text', Object.assign({x, y, fill:COLORS.ink, 'font-family':'Inter,sans-serif', 'font-size':12}, opts));
  t.textContent = str;
  svg.appendChild(t);
  return t;
}
function mean(arr){ return arr.reduce((a,b)=>a+b,0)/arr.length; }
function wrapLabel(str, maxChars){
  const words = str.split(' ');
  const lines = []; let cur = '';
  words.forEach(w=>{
    const test = cur ? cur+' '+w : w;
    if(test.length > maxChars && cur){ lines.push(cur); cur = w; } else { cur = test; }
  });
  if(cur) lines.push(cur);
  return lines;
}
function drawWrappedLabel(svg, x, yCenter, str, fontSize, maxChars, opts={}){
  const lines = wrapLabel(str, maxChars);
  const lineH = fontSize + 3;
  const startY = yCenter - ((lines.length-1)*lineH)/2 + fontSize*0.35;
  lines.forEach((line,i)=> text(svg, x, startY+i*lineH, line, Object.assign({'font-size':fontSize}, opts)));
}
function emptyState(container, msg){ container.innerHTML = `<div class="chart-empty">${msg}</div>`; }
function lerpColor(c1, c2, t){
  return c1.map((v,i)=> Math.round(v + (c2[i]-v)*t));
}

// ---------------- Filter state ----------------
function uniqueSorted(arr){ return [...new Set(arr)].filter(v=>v!==null && v!==undefined).sort((a,b)=>a.localeCompare(b,'uk')); }
const regions = DATA.regionOrder.filter(r=>DATA.records.some(x=>x.region===r));
const ages = DATA.ageOrder.filter(a=>DATA.records.some(x=>x.age===a));
const settlements = DATA.settlementOrder.filter(s=>DATA.records.some(x=>x.settlement===s));
const dispKeys = ['always','war','other'];

function fillSelect(sel, items, allLabel){
  sel.innerHTML = '';
  const o0 = document.createElement('option'); o0.value=''; o0.textContent = allLabel; sel.appendChild(o0);
  items.forEach(v=>{
    const o = document.createElement('option');
    o.value = v; o.textContent = DATA.displacementLabels[v] || v;
    sel.appendChild(o);
  });
}
fillSelect(document.getElementById('fRegion'), regions, 'Усі регіони');
fillSelect(document.getElementById('fAge'), ages, 'Усі вікові групи');
fillSelect(document.getElementById('fSettlement'), settlements, 'Усі типи поселень');
fillSelect(document.getElementById('fDisp'), dispKeys, 'Усі статуси переміщення');

const selRegion = document.getElementById('fRegion');
const selAge = document.getElementById('fAge');
const selSettlement = document.getElementById('fSettlement');
const selDisp = document.getElementById('fDisp');

function getFiltered(){
  return DATA.records.filter(r=>{
    if(selRegion.value && r.region !== selRegion.value) return false;
    if(selAge.value && r.age !== selAge.value) return false;
    if(selSettlement.value && r.settlement !== selSettlement.value) return false;
    if(selDisp.value && r.disp !== selDisp.value) return false;
    return true;
  });
}

// ---------------- Chart 1: 100% stacked bar, income by IDP status ----------------
function renderChart1(rows){
  const box = document.getElementById('chart1');
  const groups = [{key:1, label:'Має статус ВПО'}, {key:0, label:'Немає статусу ВПО'}];
  const rowsByGroup = groups.map(g => rows.filter(r=>r.idp===g.key && r.inc));
  if(!rowsByGroup.some(s=>s.length)){ emptyState(box, 'Недостатньо даних для обраних фільтрів.'); return; }

  const w = 560, barH = 42, gap = 34, top = 26, left = 152, right = 18;
  const plotW = w - left - right;
  const barsBottom = top + (groups.length-1)*(barH+gap) + barH;
  const legendTop = barsBottom + 26, legendLineH = 15;
  const h = legendTop + DATA.incomeLabels.length*legendLineH + 4;
  const svg = svgRoot(box, w, h);

  text(svg, left, top-14, '← гірше', {'font-size':10, fill:COLORS.inkSoft});
  text(svg, w-right, top-14, 'краще →', {'text-anchor':'end','font-size':10, fill:COLORS.inkSoft});

  groups.forEach((g, gi)=>{
    const subset = rowsByGroup[gi];
    const y = top + gi*(barH+gap);
    text(svg, 0, y+barH/2-6, g.label, {'font-weight':600, 'font-size':12.5});
    text(svg, 0, y+barH/2+11, subset.length ? `n = ${subset.length}` : 'немає даних', {fill:COLORS.inkSoft, 'font-size':10.5});
    if(!subset.length) return;
    let x = left;
    for(let inc=1; inc<=5; inc++){
      const share = subset.filter(r=>r.inc===inc).length / subset.length;
      const bw = share*plotW;
      if(bw>0.5){
        el('rect', {x, y, width:bw, height:barH, fill:COLORS.income[inc-1], rx:2}, svg);
        if(share>0.06) text(svg, x+bw/2, y+barH/2+4, Math.round(share*100)+'%', {fill: inc<=1||inc>=5?'#F1EEE4':'#20261F', 'text-anchor':'middle','font-size':11,'font-weight':600});
      }
      x += bw;
    }
  });

  DATA.incomeLabels.forEach((lab,i)=>{
    const ly = legendTop + i*legendLineH;
    el('rect', {x:0, y:ly-8, width:8, height:8, fill:COLORS.income[i], rx:2}, svg);
    text(svg, 13, ly, lab, {'font-size':10.5, fill:COLORS.inkSoft});
  });
}

// ---------------- Chart 2: dumbbell, job conditions by kids ----------------
function renderChart2(rows){
  const box = document.getElementById('chart2');
  const work = rows.filter(r=>r.pw===1);
  if(work.length < 4){ emptyState(box, 'Недостатньо даних для обраних фільтрів.'); return; }
  const items = DATA.conditions.map(([code,label])=>{
    const wk = work.filter(r=>r.kids===1 && r.cond[code]).map(r=>r.cond[code]);
    const nk = work.filter(r=>r.kids===0 && r.cond[code]).map(r=>r.cond[code]);
    if(!wk.length || !nk.length) return null;
    return {label, mk:mean(wk), mn:mean(nk), gap:Math.abs(mean(wk)-mean(nk))};
  }).filter(Boolean);
  if(!items.length){ emptyState(box, 'Недостатньо даних для обраних фільтрів.'); return; }
  items.sort((a,b)=>b.gap-a.gap);

  const w = 560, rowH = 34, top = 14, left = 222, right = 30;
  const plotW = w - left - right;
  const h = top + items.length*rowH + 26;
  const sx = v => left + (v-1)/2*plotW;
  const svg = svgRoot(box, w, h);
  [1,2,3].forEach(v=>{
    const x = sx(v);
    el('line', {x1:x,x2:x,y1:top-4,y2:top+items.length*rowH+2, stroke:COLORS.grid, 'stroke-dasharray':'2,3'}, svg);
    text(svg, x, top+items.length*rowH+16, v===1?'1':(v===2?'2':'3'), {'text-anchor':'middle','font-size':10, fill:COLORS.inkSoft});
  });
  items.forEach((it,i)=>{
    const y = top + i*rowH + rowH/2;
    drawWrappedLabel(svg, left-10, y, it.label, 10.2, 20, {'text-anchor':'end'});
    const x1=sx(it.mn), x2=sx(it.mk);
    el('line', {x1,x2,y1:y,y2:y, stroke:COLORS.grid, 'stroke-width':3}, svg);
    el('circle', {cx:x1, cy:y, r:5.5, fill:COLORS.noKids}, svg);
    el('circle', {cx:x2, cy:y, r:5.5, fill:COLORS.kids}, svg);
  });
}

// ---------------- Chart 3: ranked bars, channel trust ----------------
function renderChart3(rows){
  const box = document.getElementById('chart3');
  const work = rows.filter(r=>r.pw===1);
  if(work.length < 4){ emptyState(box, 'Недостатньо даних для обраних фільтрів.'); return; }
  const items = DATA.channels.map(([code,label])=>{
    const vals = work.filter(r=>r.chan[code]).map(r=>r.chan[code]);
    if(!vals.length) return null;
    return {code, label, m:mean(vals)};
  }).filter(Boolean);
  if(!items.length){ emptyState(box, 'Недостатньо даних для обраних фільтрів.'); return; }
  items.sort((a,b)=>b.m-a.m);

  const w = 1180, rowH = 34, top = 8, left = 220, right = 46;
  const plotW = w - left - right;
  const h = top + items.length*rowH + 6;
  const sx = v => (v-1)/2*plotW;
  const svg = svgRoot(box, w, h);
  items.forEach((it,i)=>{
    const y = top + i*rowH;
    const bw = sx(it.m);
    const isDsz = it.code === 'v103';
    text(svg, left-10, y+rowH/2+4, it.label, {'text-anchor':'end','font-size':12, 'font-weight':isDsz?600:400});
    el('rect', {x:left, y:y+6, width:Math.max(bw,2), height:rowH-14, fill:isDsz?COLORS.dsz:COLORS.chan, rx:3}, svg);
    text(svg, left+bw+8, y+rowH/2+4, it.m.toFixed(2), {'font-size':11.5, fill:COLORS.inkSoft});
  });
  el('line', {x1:left, x2:left, y1:top, y2:h-4, stroke:COLORS.grid}, svg);
}

// ---------------- Chart 4: heatmap, barriers by age ----------------
function renderChart4(rows){
  const box = document.getElementById('chart4');
  const work = rows.filter(r=>r.pw===1);
  const ages = DATA.ageOrder.filter(a=>work.some(r=>r.age===a));
  if(work.length < 4 || !ages.length){ emptyState(box, 'Недостатньо даних для обраних фільтрів.'); return; }

  const matrix = DATA.barriers.map(([code,label])=>{
    return ages.map(age=>{
      const vals = work.filter(r=>r.age===age && r.barr[code]).map(r=>r.barr[code]);
      return vals.length ? mean(vals) : null;
    });
  });
  const flat = matrix.flat().filter(v=>v!==null);
  if(!flat.length){ emptyState(box, 'Недостатньо даних для обраних фільтрів.'); return; }
  const vmin = Math.min(...flat), vmax = Math.max(...flat);

  // Фіксована цільова ширина ~ ширині картки (full-width), щоб SVG не
  // масштабувався вгору і шрифт не "роздувало" (ports at ~1:1 scale).
  const w = 1180, left = 292, right = 24, headH = 28, cellH = 30, rowGap = 3;
  const plotW = w - left - right;
  const cellW = plotW / ages.length;
  const top = headH;
  const h = top + DATA.barriers.length*cellH + 46;
  const svg = svgRoot(box, w, h);

  ages.forEach((age,ci)=>{
    text(svg, left+ci*cellW+cellW/2, headH-10, age, {'text-anchor':'middle','font-size':12.5,'font-weight':600});
  });
  DATA.barriers.forEach(([code,label], ri)=>{
    const y = top + ri*cellH;
    text(svg, left-10, y+cellH/2+4, label, {'text-anchor':'end','font-size':11});
    ages.forEach((age, ci)=>{
      const v = matrix[ri][ci];
      const x = left + ci*cellW;
      if(v===null){
        el('rect', {x, y, width:cellW-rowGap, height:cellH-rowGap, fill:'#EDEAE0', rx:3}, svg);
        text(svg, x+(cellW-rowGap)/2, y+(cellH-rowGap)/2+4, '—', {'text-anchor':'middle','font-size':11, fill:COLORS.inkSoft});
        return;
      }
      const t = vmax>vmin ? (v-vmin)/(vmax-vmin) : 0.5;
      const [r,g,b] = lerpColor(COLORS.heatLo, COLORS.heatHi, t);
      el('rect', {x, y, width:cellW-rowGap, height:cellH-rowGap, fill:`rgb(${r},${g},${b})`, rx:3}, svg);
      text(svg, x+(cellW-rowGap)/2, y+(cellH-rowGap)/2+4, v.toFixed(2), {'text-anchor':'middle','font-size':12, fill: t>0.55?'#F1EEE4':'#20261F','font-weight':600});
    });
  });
  // legend scale — підписи по боках градієнта, щоб не накладались одне на одне
  const legY = top + DATA.barriers.length*cellH + 22;
  const legW = 180, legX = left + 118;
  const steps = 24;
  for(let i=0;i<steps;i++){
    const t = i/(steps-1);
    const [r,g,b] = lerpColor(COLORS.heatLo, COLORS.heatHi, t);
    el('rect', {x:legX+i*(legW/steps), y:legY, width:legW/steps+0.5, height:9, fill:`rgb(${r},${g},${b})`}, svg);
  }
  text(svg, legX-8, legY+8, vmin.toFixed(2)+' менш серйозно', {'font-size':10.5, fill:COLORS.inkSoft, 'text-anchor':'end'});
  text(svg, legX+legW+8, legY+8, vmax.toFixed(2)+' серйозніше', {'font-size':10.5, fill:COLORS.inkSoft});
}

// ---------------- Chart 5: salary expectation by income adequacy (100% stacked) ----------------
function renderChart5(rows){
  const box = document.getElementById('chart5');
  const withData = rows.filter(r=>r.inc && r.sal);
  if(withData.length < 4){ emptyState(box, 'Недостатньо даних для обраних фільтрів.'); return; }

  const w = 560, barH = 38, gap = 16, top = 8, left = 142, right = 40;
  const plotW = w - left - right;
  const incLevels = [5,4,3,2,1];
  const rowsH = incLevels.length*(barH+gap);
  const legendH = 40;
  const h = top + rowsH + legendH;
  const svg = svgRoot(box, w, h);

  incLevels.forEach((inc, gi)=>{
    const subset = withData.filter(r=>r.inc===inc);
    const y = top + gi*(barH+gap);
    drawWrappedLabel(svg, 0, y+barH/2, DATA.incomeLabels[inc-1], 10.5, 22, {fill:COLORS.ink});
    if(!subset.length){ text(svg, left, y+barH/2+4, 'немає даних', {fill:COLORS.inkSoft,'font-size':10.5}); return; }
    let x = left;
    for(let sal=1; sal<=8; sal++){
      const share = subset.filter(r=>r.sal===sal).length / subset.length;
      const bw = share*plotW;
      if(bw>0.4) el('rect', {x, y, width:bw, height:barH, fill:COLORS.salarySeq[sal-1], rx:2}, svg);
      x += bw;
    }
    text(svg, left+plotW+6, y+barH/2+4, `n=${subset.length}`, {'font-size':9.5, fill:COLORS.inkSoft});
  });

  // легенда: колірні плашки + один короткий підпис напряму шкали
  const legY = top + rowsH + 12;
  let lx = 0;
  DATA.salaryLabels.forEach((lab,i)=>{
    el('rect', {x:lx, y:legY, width:9, height:9, fill:COLORS.salarySeq[i], rx:2}, svg);
    lx += 23;
  });
  text(svg, 0, legY+24, 'Очікувана зарплата: від <9 тис до >100 тис грн (зліва направо)', {'font-size':10, fill:COLORS.inkSoft});
}

// ---------------- Chart 6: atypical profession readiness by IDP ----------------
function renderChart6(rows){
  const box = document.getElementById('chart6');
  const groups = [{key:1, label:'ВПО'}, {key:0, label:'Місцеві (не ВПО)'}];
  const cats = [['yes','Готова'], ['unsure','Важко сказати'], ['no','Не готова']];
  const colors = {yes:COLORS.yes, unsure:COLORS.unsure, no:COLORS.no};

  const w = 380, h = 230, top = 14, bottom = 34, left = 30, right = 10;
  const plotH = h - top - bottom;
  const groupW = (w-left-right)/groups.length;
  const barW = 30, barGap = 6;

  const data = groups.map(g=>{
    const subset = rows.filter(r=>r.idp===g.key && r.atyp);
    return cats.map(([key])=> subset.length ? subset.filter(r=>r.atyp===key).length/subset.length : null).concat([subset.length]);
  });
  if(data.every(d=>d[3]===0)){ emptyState(box, 'Недостатньо даних для обраних фільтрів.'); return; }

  const svg = svgRoot(box, w, h);
  el('line', {x1:left, x2:w-right, y1:top+plotH, y2:top+plotH, stroke:COLORS.grid}, svg);
  groups.forEach((g,gi)=>{
    const n = data[gi][3];
    const gx = left + gi*groupW + groupW/2;
    text(svg, gx, top+plotH+16, g.label + (n?` (n=${n})`:''), {'text-anchor':'middle','font-size':10.5});
    cats.forEach(([key,label], ci)=>{
      const share = data[gi][ci];
      const bx = gx - (cats.length*(barW+barGap))/2 + ci*(barW+barGap);
      const bh = (share||0)*plotH;
      el('rect', {x:bx, y:top+plotH-bh, width:barW, height:bh, fill:colors[key], rx:2}, svg);
      if(share) text(svg, bx+barW/2, top+plotH-bh-5, Math.round(share*100)+'%', {'text-anchor':'middle','font-size':10,'font-weight':600});
    });
  });
  const legY = 8;
  let lx = left;
  cats.forEach(([key,label])=>{
    el('rect', {x:lx, y:legY-8, width:8, height:8, fill:colors[key], rx:2}, svg);
    text(svg, lx+12, legY, label, {'font-size':9.5, fill:COLORS.inkSoft});
    lx += label.length*5.3 + 26;
  });
}

// ---------------- Chart 7: sample portrait (age / region / settlement) ----------------
function renderChart7(rows){
  const box = document.getElementById('chart7');
  if(!rows.length){ emptyState(box, 'Недостатньо даних.'); return; }
  const panels = [
    {title:'Вік', items: DATA.ageOrder.map(a=>[a, rows.filter(r=>r.age===a).length])},
    {title:'Регіон', items: DATA.regionOrder.map(a=>[a, rows.filter(r=>r.region===a).length])},
    {title:'Поселення', items: DATA.settlementOrder.map(a=>[a, rows.filter(r=>r.settlement===a).length])},
  ];
  const colW = 190, rowH = 34, top = 26, gapRight = 14;
  const maxRows = Math.max(...panels.map(p=>p.items.length));
  const h = top + maxRows*rowH + 4;
  const w = colW*panels.length;
  const svg = svgRoot(box, w, h);
  const barMaxW = colW - gapRight - 40;
  panels.forEach((p, pi)=>{
    const x0 = pi*colW;
    text(svg, x0, 12, p.title, {'font-weight':600, 'font-size':12});
    const maxV = Math.max(...p.items.map(i=>i[1]), 1);
    p.items.forEach(([label,count], ri)=>{
      const y = top + ri*rowH;
      text(svg, x0, y+9, label, {'font-size':10.5, fill:COLORS.inkSoft});
      const bw = Math.max((count/maxV)*barMaxW, 3);
      el('rect', {x:x0, y:y+16, width:bw, height:6, fill:COLORS.steel, rx:3}, svg);
      text(svg, x0+bw+7, y+21.5, count, {'font-size':10, fill:COLORS.ink, 'font-weight':600});
    });
  });
}

// ---------------- Chart 8: top sectors ----------------
function renderChart8(rows){
  const box = document.getElementById('chart8');
  const work = rows.filter(r=>r.pw===1);
  if(work.length < 4){ emptyState(box, 'Недостатньо даних для обраних фільтрів.'); return; }
  const counts = new Array(20).fill(0);
  work.forEach(r=> r.sec.forEach(i=> counts[i-1]++));
  const items = counts.map((c,i)=>({label:DATA.sectorLabels[i], share:c/work.length}))
    .sort((a,b)=>b.share-a.share).slice(0,10);

  const w = 500, rowH = 27, top = 6, left = 230, right = 40;
  const plotW = w-left-right;
  const h = top + items.length*rowH + 4;
  const maxShare = Math.max(...items.map(i=>i.share), 0.01);
  const svg = svgRoot(box, w, h);
  items.forEach((it,i)=>{
    const y = top+i*rowH;
    const bw = (it.share/maxShare)*plotW;
    text(svg, left-8, y+rowH/2+4, it.label, {'text-anchor':'end','font-size':11});
    el('rect', {x:left, y:y+5, width:Math.max(bw,2), height:rowH-12, fill:COLORS.green, rx:3}, svg);
    text(svg, left+bw+6, y+rowH/2+4, Math.round(it.share*100)+'%', {'font-size':10.5, fill:COLORS.inkSoft});
  });
}

function updateBadgesAndCaptions(){
  const s = DATA.stats;
  const p = v => v < 0.001 ? '< 0.001' : v.toFixed(3);
  document.getElementById('badge1').textContent = `p ${p(s.idp_income.p_value)}`;
  document.getElementById('cap1').innerHTML = `Достаток нижчий серед жінок зі статусом ВПО: середній бал <b>${s.idp_income.mean_idp}</b> проти <b>${s.idp_income.mean_non_idp}</b> без статусу (шкала 1–5).`;

  const care = s.kids_conditions['v77'];
  document.getElementById('badge2').textContent = `p ${p(care.p_value)}`;
  document.getElementById('cap2').innerHTML = `«Дитсадок поруч» — найбільший розрив (<b>${Math.abs(care.mean_kids-care.mean_no_kids).toFixed(2)}</b> бала) серед 9 умов; гнучкий графік і віддалена робота різниці не показують.`;

  document.getElementById('badge3').textContent = `p ${p(s.channels.p_value)}`;
  document.getElementById('cap3').innerHTML = `ДСЗ — останнє місце з 6 каналів (<b>${s.channels.mean_dsz}</b> проти <b>${s.channels.mean_aggregators}</b> у агрегаторів). Серед тих, хто реально звертався по допомогу (n=${s.channels.n_searched_via_dsz}), <b>${Math.round(s.channels.share_not_useful*100)}%</b> оцінили це як «не корисно».`;
}

function renderAll(){
  const rows = getFiltered();
  document.getElementById('nShown').textContent = rows.length;
  renderChart1(rows); renderChart2(rows); renderChart3(rows);
  renderChart4(rows); renderChart5(rows); renderChart6(rows);
  renderChart7(rows); renderChart8(rows);
}

[selRegion, selAge, selSettlement, selDisp].forEach(s=>s.addEventListener('change', renderAll));
document.getElementById('resetBtn').addEventListener('click', ()=>{
  selRegion.value=''; selAge.value=''; selSettlement.value=''; selDisp.value='';
  renderAll();
});

document.getElementById('metaTotal').textContent = DATA.records.length;
updateBadgesAndCaptions();
renderAll();
</script>

</body>
</html>
"""


# MAIN

def main():
    parser = argparse.ArgumentParser(description="Аналіз опитування та генерація HTML-дашборду")
    parser.add_argument("--input", required=True, help="Шлях до .xlsx з мікроданими")
    parser.add_argument("--output", required=True, help="Шлях для збереження .html звіту")
    args = parser.parse_args()

    print(f"[1/4] Завантаження даних з {args.input} ...")
    df = load_data(args.input)
    print(f"      Завантажено {len(df)} рядків, {len(df.columns)} стовпців.")

    print("[2/4] Статистична перевірка трьох основних спостережень ...")
    stats_results = run_significance_checks(df)
    print(json.dumps(stats_results, ensure_ascii=False, indent=2))

    print("[3/4] Побудова рядкового датасету для клієнтських фільтрів ...")
    records = build_records(df)

    print("[4/4] Генерація HTML-дашборду ...")
    html = build_html(records, {"n_total": len(df)}, stats_results)
    out_path = Path(args.output)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(html, encoding="utf-8")
    print(f"      Готово: {out_path} ({out_path.stat().st_size/1024:.0f} KB)")


if __name__ == "__main__":
    sys.exit(main())
