#!/usr/bin/env python3
"""
payments_yearly.py — scan a folder of monthly payment Excel files
and build a yearly summary HTML.

Usage:
    python3 payments_yearly.py ./folder/                 # one-shot scan
    python3 payments_yearly.py ./folder/ --watch         # poll every 5s
    python3 payments_yearly.py ./folder/ --year 2026     # only files whose payments land in 2026
    python3 payments_yearly.py ./folder/ --open          # open the HTML when done
    python3 payments_yearly.py ./folder/ --interval 10   # custom poll interval for --watch

The script scans every ``*.xlsx`` in the folder, parses each with
``payments_ui.parse_payments`` and groups the transactions by year/month.
Per year it writes ``yearly_summary_<YEAR>.html`` into the folder.
"""
from __future__ import annotations

import json
import re
import sys
import time
import webbrowser
from collections import defaultdict
from pathlib import Path

from payments_ui import (
    HIGH_THRESHOLD,
    _normalize_merchant,
    build_insights,
    parse_payments,
)


def _month_label(ym: str) -> str:
    months = [
        "ינואר", "פברואר", "מרץ", "אפריל", "מאי", "יוני",
        "יולי", "אוגוסט", "ספטמבר", "אוקטובר", "נובמבר", "דצמבר",
    ]
    y, m = ym.split("-")
    return f"{months[int(m) - 1]} {y}"


def scan_folder(folder: Path, year_filter: int | None = None) -> dict[int, dict]:
    """Parse every xlsx in ``folder`` and group payments by year."""
    per_year: dict[int, list] = defaultdict(list)
    sources_per_year: dict[int, list] = defaultdict(list)

    for xlsx in sorted(folder.glob("*.xlsx")):
        # Skip Excel lock files
        if xlsx.name.startswith("~$"):
            continue
        try:
            data = parse_payments(xlsx)
        except Exception as e:
            print(f"  skip {xlsx.name}: {e}", file=sys.stderr)
            continue

        seen_years = set()
        for p in data["payments"]:
            if not p["date"]:
                continue
            try:
                year = int(p["date"][:4])
            except ValueError:
                continue
            if year_filter is not None and year != year_filter:
                continue
            per_year[year].append(p)
            seen_years.add(year)

        for y in seen_years:
            sources_per_year[y].append({
                "name": xlsx.name,
                "issuer": data.get("issuer", ""),
                "title": data.get("title", xlsx.stem),
            })

    return {
        year: {
            "year": year,
            "payments": payments,
            "sources": sources_per_year[year],
        }
        for year, payments in per_year.items()
    }


def build_yearly_summary(year_data: dict) -> dict:
    payments = year_data["payments"]

    # Per-month breakdown
    monthly = defaultdict(lambda: {"count": 0, "total": 0.0})
    for p in payments:
        ym = p["date"][:7] if p["date"] else "unknown"
        monthly[ym]["count"] += 1
        monthly[ym]["total"] += p["charge"]
    months = sorted(
        [
            {
                "ym": ym,
                "label": _month_label(ym) if ym != "unknown" else "—",
                "count": d["count"],
                "total": round(d["total"], 2),
            }
            for ym, d in monthly.items() if ym != "unknown"
        ],
        key=lambda x: x["ym"],
    )

    # Per-category for the year
    cats = defaultdict(lambda: {"count": 0, "total": 0.0})
    for p in payments:
        c = p["category"] or "ללא קטגוריה"
        cats[c]["count"] += 1
        cats[c]["total"] += p["charge"]
    categories = sorted(
        [{"name": k, "count": v["count"], "total": round(v["total"], 2)} for k, v in cats.items()],
        key=lambda x: x["total"],
        reverse=True,
    )

    # Top merchants for the year (normalized)
    mer = defaultdict(lambda: {"count": 0, "total": 0.0, "aliases": set()})
    for p in payments:
        name = p.get("canonical") or _normalize_merchant(p["merchant"])
        mer[name]["count"] += 1
        mer[name]["total"] += p["charge"]
        mer[name]["aliases"].add(p["merchant"])
    top_merchants = sorted(
        [
            {
                "name": k,
                "count": v["count"],
                "total": round(v["total"], 2),
                "aliases": sorted(v["aliases"]),
            }
            for k, v in mer.items()
        ],
        key=lambda x: x["total"],
        reverse=True,
    )[:25]

    # Reuse the single-file insights for the whole year (duplicates/subs/installments)
    insights = build_insights(payments)

    # Month-over-month category breakdown (for stacked bar chart)
    all_yms = sorted({p["date"][:7] for p in payments if p["date"] and len(p["date"]) >= 7})
    cat_by_month: dict = defaultdict(lambda: defaultdict(float))
    for p in payments:
        ym = p["date"][:7] if p["date"] and len(p["date"]) >= 7 else None
        if not ym:
            continue
        cat_by_month[p["category"] or "ללא קטגוריה"][ym] += p["charge"]
    top_cat_names = [c["name"] for c in categories[:8]]
    monthly_by_category = {
        "months": all_yms,
        "month_labels": [_month_label(ym) for ym in all_yms],
        "categories": [
            {
                "name": name,
                "data": [round(cat_by_month[name].get(ym, 0), 2) for ym in all_yms],
            }
            for name in top_cat_names
        ],
    }

    # Per-merchant per-month totals (for trend arrows)
    mer_by_month: dict = defaultdict(lambda: defaultdict(float))
    for p in payments:
        ym = p["date"][:7] if p["date"] and len(p["date"]) >= 7 else None
        if not ym:
            continue
        name = p.get("canonical") or _normalize_merchant(p["merchant"])
        mer_by_month[name][ym] += p["charge"]
    merchant_monthly = {
        name: {ym: round(total, 2) for ym, total in d.items()}
        for name, d in mer_by_month.items()
    }

    return {
        "year": year_data["year"],
        "sources": year_data["sources"],
        "total_count": len(payments),
        "total_amount": round(sum(p["charge"] for p in payments), 2),
        "months": months,
        "categories": categories,
        "top_merchants": top_merchants,
        "monthly_by_category": monthly_by_category,
        "merchant_monthly": merchant_monthly,
        "duplicates": insights["duplicates"][:50],
        "subscriptions": insights["subscriptions"],
        "installments": insights["installments"],
        "total_installment_remaining": insights["total_installment_remaining"],
        "high_threshold": HIGH_THRESHOLD,
    }


YEARLY_HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="he" dir="rtl" data-theme="light">
<head>
<meta charset="UTF-8">
<title>סיכום שנתי __YEAR__</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
  :root {
    --bg: #f5f5f7; --card: #fff; --text: #222; --muted: #666; --soft: #888;
    --border: #eee; --border-strong: #ddd; --hover: #f8f8fb; --th-bg: #fafafa;
    --shadow: 0 1px 3px rgba(0,0,0,0.08);
    --primary: #2196f3; --high: #c62828; --refund: #2e7d32;
  }
  [data-theme="dark"] {
    --bg: #111418; --card: #1c2128; --text: #e6edf3; --muted: #9aa4af; --soft: #7a8591;
    --border: #2a323c; --border-strong: #394350; --hover: #232b36; --th-bg: #1a2028;
    --shadow: 0 1px 3px rgba(0,0,0,0.4);
    --primary: #64b5f6; --high: #ef5350; --refund: #66bb6a;
  }
  * { box-sizing: border-box; }
  body { font-family: -apple-system, "Segoe UI", Arial, sans-serif; margin: 0; padding: 24px;
         background: var(--bg); color: var(--text); transition: background 0.2s, color 0.2s; }
  header { display: flex; justify-content: space-between; align-items: flex-start; gap: 16px; margin-bottom: 20px; }
  header h1 { font-size: 22px; margin: 0 0 4px; }
  .src { color: var(--soft); font-size: 12px; line-height: 1.6; }
  .theme-toggle { background: var(--card); color: var(--text); border: 1px solid var(--border-strong);
                  width: 40px; height: 40px; border-radius: 10px; cursor: pointer; font-size: 18px; box-shadow: var(--shadow); }
  .cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 12px; margin-bottom: 20px; }
  .card { background: var(--card); padding: 16px; border-radius: 10px; box-shadow: var(--shadow); }
  .card .label { font-size: 11px; color: var(--muted); letter-spacing: 0.5px; }
  .card .value { font-size: 22px; font-weight: 600; margin-top: 4px; }
  .card .sub { font-size: 12px; color: var(--soft); margin-top: 2px; }
  .charts-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(340px, 1fr)); gap: 16px; margin-bottom: 20px; }
  .chart-card { background: var(--card); padding: 16px 20px; border-radius: 10px; box-shadow: var(--shadow); }
  .chart-card h2 { margin: 0 0 12px; font-size: 15px; }
  .chart-wrap { position: relative; height: 280px; }
  .section { background: var(--card); padding: 16px 20px; border-radius: 10px; box-shadow: var(--shadow); margin-bottom: 20px; }
  .section > h2 { margin: 0 0 12px; font-size: 16px; cursor: pointer; user-select: none; }
  .section.collapsed > :not(h2) { display: none; }
  .section > h2::before { content: "▾"; display: inline-block; width: 1em; font-size: 11px; color: var(--muted); transform: scaleX(-1); }
  .section.collapsed > h2::before { content: "▸"; transform: scaleX(-1); }
  table { width: 100%; border-collapse: collapse; font-size: 14px; }
  th, td { padding: 8px 10px; border-bottom: 1px solid var(--border); text-align: right; }
  th { background: var(--th-bg); font-weight: 600; font-size: 13px; }
  tr:hover td { background: var(--hover); }
  .num { font-variant-numeric: tabular-nums; white-space: nowrap; }
  .amount-high { color: var(--high); font-weight: 600; }
  .count { color: var(--soft); font-size: 12px; font-weight: normal; margin-right: 6px; }
  .aliases { color: var(--soft); font-size: 11px; }
  .empty { color: var(--soft); font-style: italic; padding: 8px 0; }
  .sum-row td { border-top: 2px solid var(--border-strong); border-bottom: none; font-weight: 700; background: var(--th-bg); }
  .src ul { margin: 4px 0 0; padding-right: 16px; }
  .trend-up   { color: #c62828; font-size: 11px; font-weight: 700; white-space: nowrap; margin-right: 4px; }
  .trend-down { color: #2e7d32; font-size: 11px; font-weight: 700; white-space: nowrap; margin-right: 4px; }
  .trend-flat { color: var(--soft); font-size: 11px; margin-right: 4px; }
  .insights-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 10px; }
  .insight-card { display: flex; align-items: flex-start; gap: 10px; padding: 10px 14px;
    border-radius: 8px; border-right: 4px solid transparent; background: var(--bg); font-size: 14px; line-height: 1.5; }
  .insight-card.ok    { border-color: #43a047; }
  .insight-card.warn  { border-color: #fb8c00; }
  .insight-card.alert { border-color: var(--high); }
  .insight-card .ic   { font-size: 20px; line-height: 1; flex-shrink: 0; margin-top: 1px; }
  .insight-card .body { flex: 1; color: var(--text); }
  .insight-card .body strong { color: var(--primary); font-weight: 700; }
  .chart-card.full { grid-column: 1 / -1; }
  .chart-wrap.tall { height: 340px; }
</style>
</head>
<body>
<header>
  <div>
    <h1>סיכום שנתי __YEAR__</h1>
    <div class="src" id="src"></div>
  </div>
  <button class="theme-toggle" id="theme-toggle" title="מצב כהה/בהיר">🌙</button>
</header>

<div class="cards" id="cards"></div>

<div class="charts-grid">
  <div class="chart-card"><h2>הוצאה חודשית</h2><div class="chart-wrap"><canvas id="chart-monthly"></canvas></div></div>
  <div class="chart-card"><h2>חלוקה לפי ענף</h2><div class="chart-wrap"><canvas id="chart-category"></canvas></div></div>
  <div class="chart-card"><h2>Top בתי עסק</h2><div class="chart-wrap"><canvas id="chart-merchants"></canvas></div></div>
  <div class="chart-card full"><h2>קטגוריות לפי חודש</h2><div class="chart-wrap tall"><canvas id="chart-cat-monthly"></canvas></div></div>
</div>

<div class="section" id="sec-insights">
  <h2>תובנות שנתיות 🧠</h2>
  <div class="insights-grid" id="insights-grid"></div>
</div>

<div class="section" id="sec-months">
  <h2>פירוט חודשי</h2>
  <table><thead><tr><th>חודש</th><th>עסקאות</th><th>סה"כ (₪)</th><th>שינוי</th></tr></thead><tbody id="months-tbody"></tbody></table>
</div>

<div class="section" id="sec-merchants">
  <h2>Top בתי עסק <span class="count" id="merchants-count"></span></h2>
  <table><thead><tr><th>בית עסק</th><th>עסקאות</th><th>סה"כ (₪)</th><th>ממוצע (₪)</th><th>מגמה</th></tr></thead><tbody id="merchants-tbody"></tbody></table>
</div>

<div class="section" id="sec-subscriptions">
  <h2>הוראות קבע <span class="count" id="subs-count"></span></h2>
  <table><thead><tr><th>בית עסק</th><th>חיובים</th><th>סה"כ (₪)</th></tr></thead><tbody id="subs-tbody"></tbody></table>
</div>

<div class="section" id="sec-installments">
  <h2>תשלומים פתוחים <span class="count" id="inst-count"></span></h2>
  <div class="src" id="inst-remaining"></div>
  <table><thead><tr><th>תאריך</th><th>בית עסק</th><th>חודשי (₪)</th><th>נותרו</th><th>יתרה (₪)</th></tr></thead><tbody id="inst-tbody"></tbody></table>
</div>

<div class="section" id="sec-duplicates">
  <h2>חיובים כפולים אפשריים <span class="count" id="dup-count"></span></h2>
  <table><thead><tr><th>תאריך</th><th>בית עסק</th><th>סכום בודד</th><th>חזרות</th><th>סה"כ</th></tr></thead><tbody id="dup-tbody"></tbody></table>
</div>

<script>
const DATA = __DATA__;

const fmt = n => new Intl.NumberFormat('he-IL', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(n);
const fmt0 = n => new Intl.NumberFormat('he-IL', { maximumFractionDigits: 0 }).format(n);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[c]));

// ── Trend arrow ──────────────────────────────────────────────────────────────
function trendArrow(prev, curr) {
  if (prev == null || prev === 0) return '<span class="trend-flat">—</span>';
  const pct = (curr - prev) / Math.abs(prev) * 100;
  if (pct > 5)  return `<span class="trend-up">↑ ${Math.round(pct)}%</span>`;
  if (pct < -5) return `<span class="trend-down">↓ ${Math.abs(Math.round(pct))}%</span>`;
  return '<span class="trend-flat">→</span>';
}

function applyTheme(t) {
  document.documentElement.dataset.theme = t;
  document.getElementById('theme-toggle').textContent = t === 'dark' ? '☀️' : '🌙';
  renderCharts();
}
function initTheme() {
  const saved = localStorage.getItem('payments-theme') ||
    (window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
  applyTheme(saved);
  document.getElementById('theme-toggle').addEventListener('click', () => {
    const next = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark';
    localStorage.setItem('payments-theme', next);
    applyTheme(next);
  });
}

function renderSrc() {
  const items = DATA.sources.map(s => `<li>${esc(s.name)}${s.issuer ? ` · <span style="opacity:0.7">${esc(s.issuer)}</span>` : ''}</li>`).join('');
  document.getElementById('src').innerHTML = `מקורות (${DATA.sources.length}):<ul>${items}</ul>`;
}

function renderCards() {
  const avg = DATA.months.length ? DATA.total_amount / DATA.months.length : 0;
  const biggestMonth = [...DATA.months].sort((a, b) => b.total - a.total)[0];
  const cards = [
    ['סה"כ שנתי', '₪' + fmt(DATA.total_amount), DATA.total_count + ' עסקאות'],
    ['ממוצע חודשי', '₪' + fmt(avg), DATA.months.length + ' חודשים'],
    ['חודש יקר ביותר', biggestMonth ? ('₪' + fmt(biggestMonth.total)) : '—', biggestMonth ? biggestMonth.label : ''],
    ['יתרת תשלומים עתידית', '₪' + fmt(DATA.total_installment_remaining), DATA.installments.length + ' תוכניות'],
  ];
  document.getElementById('cards').innerHTML = cards
    .map(([l, v, sub]) => `<div class="card"><div class="label">${esc(l)}</div><div class="value">${esc(v)}</div><div class="sub">${esc(sub)}</div></div>`)
    .join('');
}

let charts = {};
function destroyCharts() { Object.values(charts).forEach(c => c && c.destroy()); charts = {}; }

function chartColors(n) {
  const base = ['#2196f3','#ef6c00','#43a047','#8e24aa','#e53935','#00acc1','#fb8c00','#6d4c41','#546e7a','#d81b60','#7cb342','#3949ab'];
  return Array.from({ length: n }, (_, i) => base[i % base.length]);
}

function renderCharts() {
  if (typeof Chart === 'undefined') return;
  destroyCharts();
  const text = getComputedStyle(document.documentElement).getPropertyValue('--text').trim() || '#222';
  const grid = getComputedStyle(document.documentElement).getPropertyValue('--border').trim() || '#eee';
  Chart.defaults.color = text;
  Chart.defaults.borderColor = grid;

  // Monthly total bar chart
  charts.monthly = new Chart(document.getElementById('chart-monthly'), {
    type: 'bar',
    data: {
      labels: DATA.months.map(m => m.label),
      datasets: [{ data: DATA.months.map(m => m.total), backgroundColor: '#2196f3' }],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: { y: { ticks: { callback: v => '₪' + fmt0(v), font: { size: 10 } } }, x: { ticks: { font: { size: 10 } } } },
    },
  });

  // Category doughnut
  const cats = DATA.categories.slice(0, 10);
  charts.cat = new Chart(document.getElementById('chart-category'), {
    type: 'doughnut',
    data: { labels: cats.map(c => c.name), datasets: [{ data: cats.map(c => c.total), backgroundColor: chartColors(cats.length), borderWidth: 0 }] },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom', labels: { boxWidth: 12, font: { size: 11 } } } } },
  });

  // Top merchants bar chart
  const tm = DATA.top_merchants.slice(0, 10);
  charts.merchants = new Chart(document.getElementById('chart-merchants'), {
    type: 'bar',
    data: { labels: tm.map(m => m.name), datasets: [{ data: tm.map(m => m.total), backgroundColor: '#43a047' }] },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: { x: { ticks: { callback: v => '₪' + fmt0(v), font: { size: 10 } } }, y: { ticks: { font: { size: 11 } } } },
    },
  });

  // ── Category-by-month stacked bar chart ────────────────────────────────
  const mbc = DATA.monthly_by_category;
  if (mbc && mbc.months.length && mbc.categories.length) {
    const colors = chartColors(mbc.categories.length);
    charts.catMonthly = new Chart(document.getElementById('chart-cat-monthly'), {
      type: 'bar',
      data: {
        labels: mbc.month_labels,
        datasets: mbc.categories.map((cat, i) => ({
          label: cat.name,
          data: cat.data,
          backgroundColor: colors[i],
          stack: 'stack',
        })),
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { position: 'bottom', labels: { boxWidth: 12, font: { size: 11 } } },
          tooltip: {
            callbacks: {
              label: ctx => ` ${ctx.dataset.label}: ₪${fmt0(ctx.parsed.y)}`,
            },
          },
        },
        scales: {
          x: { stacked: true, ticks: { font: { size: 11 } } },
          y: { stacked: true, ticks: { callback: v => '₪' + fmt0(v), font: { size: 10 } } },
        },
      },
    });
  }
}

// ── Monthly table with trend arrows ─────────────────────────────────────────
function renderMonths() {
  const rows = DATA.months;
  if (!rows.length) { document.getElementById('months-tbody').innerHTML = `<tr><td colspan="4" class="empty">—</td></tr>`; return; }
  const totalCount = rows.reduce((s, m) => s + m.count, 0);
  const totalAmount = rows.reduce((s, m) => s + m.total, 0);
  document.getElementById('months-tbody').innerHTML = rows.map((m, i) => {
    const prev = i > 0 ? rows[i - 1].total : null;
    return `<tr>
      <td>${esc(m.label)}</td>
      <td class="num">${m.count}</td>
      <td class="num">${fmt(m.total)}</td>
      <td class="num">${trendArrow(prev, m.total)}</td>
    </tr>`;
  }).join('') + `<tr class="sum-row"><td>סה"כ</td><td class="num">${totalCount}</td><td class="num">${fmt(totalAmount)}</td><td></td></tr>`;
}

// ── Merchants table with trend arrows ────────────────────────────────────────
function renderMerchants() {
  const rows = DATA.top_merchants;
  document.getElementById('merchants-count').textContent = `(${rows.length})`;
  if (!rows.length) { document.getElementById('merchants-tbody').innerHTML = `<tr><td colspan="5" class="empty">—</td></tr>`; return; }
  const totalCount = rows.reduce((s, m) => s + m.count, 0);
  const totalAmount = rows.reduce((s, m) => s + m.total, 0);

  // Get last two available months across the year
  const allYms = DATA.monthly_by_category?.months ?? [];
  const lastYm   = allYms[allYms.length - 1] ?? null;
  const beforeYm = allYms[allYms.length - 2] ?? null;

  document.getElementById('merchants-tbody').innerHTML = rows.map(m => {
    const mm = DATA.merchant_monthly?.[m.name] ?? {};
    const curr = lastYm   ? (mm[lastYm]   ?? 0) : null;
    const prev = beforeYm ? (mm[beforeYm] ?? 0) : null;
    const arrow = (curr !== null && prev !== null) ? trendArrow(prev, curr) : '<span class="trend-flat">—</span>';
    return `<tr>
      <td>${esc(m.name)}${m.aliases.length > 1 ? `<div class="aliases">${esc(m.aliases.join(' · '))}</div>` : ''}</td>
      <td class="num">${m.count}</td>
      <td class="num">${fmt(m.total)}</td>
      <td class="num">${fmt(m.total / m.count)}</td>
      <td class="num">${arrow}</td>
    </tr>`;
  }).join('') + `<tr class="sum-row"><td>סה"כ</td><td class="num">${totalCount}</td><td class="num">${fmt(totalAmount)}</td><td></td><td></td></tr>`;
}

function renderSubs() {
  const rows = DATA.subscriptions;
  document.getElementById('subs-count').textContent = `(${rows.length})`;
  if (!rows.length) { document.getElementById('subs-tbody').innerHTML = `<tr><td colspan="3" class="empty">—</td></tr>`; return; }
  const totalCount = rows.reduce((s, r) => s + r.count, 0);
  const totalAmount = rows.reduce((s, r) => s + r.total, 0);
  document.getElementById('subs-tbody').innerHTML = rows.map(s => `
    <tr><td>${esc(s.merchant)}</td><td class="num">${s.count}</td><td class="num">${fmt(s.total)}</td></tr>
  `).join('') + `<tr class="sum-row"><td>סה"כ</td><td class="num">${totalCount}</td><td class="num">${fmt(totalAmount)}</td></tr>`;
}

function renderInstallments() {
  const rows = DATA.installments;
  document.getElementById('inst-count').textContent = `(${rows.length})`;
  document.getElementById('inst-remaining').textContent = `יתרה עתידית כוללת: ₪${fmt(DATA.total_installment_remaining)}`;
  if (!rows.length) { document.getElementById('inst-tbody').innerHTML = `<tr><td colspan="5" class="empty">אין תשלומים פתוחים</td></tr>`; return; }
  const totalCharge = rows.reduce((s, i) => s + i.charge, 0);
  const totalRemaining = rows.reduce((s, i) => s + i.remaining_amount, 0);
  document.getElementById('inst-tbody').innerHTML = rows.map(i => `
    <tr>
      <td class="num">${esc(i.date)}</td>
      <td>${esc(i.merchant)}</td>
      <td class="num">${fmt(i.charge)}</td>
      <td class="num">${i.remaining_count || 0}</td>
      <td class="num ${i.remaining_amount > 0 ? 'amount-high' : ''}">${fmt(i.remaining_amount)}</td>
    </tr>
  `).join('') + `<tr class="sum-row"><td colspan="2">סה"כ</td><td class="num">${fmt(totalCharge)}</td><td></td><td class="num">${fmt(totalRemaining)}</td></tr>`;
}

function renderDuplicates() {
  const rows = DATA.duplicates;
  document.getElementById('dup-count').textContent = `(${rows.length})`;
  if (!rows.length) { document.getElementById('dup-tbody').innerHTML = `<tr><td colspan="5" class="empty">—</td></tr>`; return; }
  const total = rows.reduce((s, d) => s + d.total, 0);
  document.getElementById('dup-tbody').innerHTML = rows.map(d => `
    <tr>
      <td class="num">${esc(d.date)}</td>
      <td>${esc(d.merchant)}</td>
      <td class="num">${fmt(d.amount)}</td>
      <td class="num">${d.count}</td>
      <td class="num amount-high">${fmt(d.total)}</td>
    </tr>
  `).join('') + `<tr class="sum-row"><td colspan="4">סה"כ</td><td class="num">${fmt(total)}</td></tr>`;
}

// ── Yearly insights ──────────────────────────────────────────────────────────
function renderYearlyInsights() {
  const months = DATA.months;
  const cats   = DATA.categories;
  const top    = DATA.top_merchants;
  const subs   = DATA.subscriptions;
  const items  = [];

  // Total overview
  if (months.length) {
    const avg = DATA.total_amount / months.length;
    items.push({ ic: '📅', level: 'ok',
      html: `<strong>${months.length} חודשים</strong> של נתונים · ממוצע חודשי <strong>₪${fmt(avg)}</strong>` });
  }

  // Most & least expensive month
  if (months.length >= 2) {
    const sorted = [...months].sort((a, b) => b.total - a.total);
    const peak = sorted[0], low = sorted[sorted.length - 1];
    items.push({ ic: '📈', level: peak.total > DATA.total_amount / months.length * 1.5 ? 'warn' : 'ok',
      html: `חודש הוצאה גבוה ביותר: <strong>${esc(peak.label)}</strong> — ₪${fmt(peak.total)}<br>חודש הוצאה נמוך ביותר: <strong>${esc(low.label)}</strong> — ₪${fmt(low.total)}` });
  }

  // Trend: last month vs month before
  if (months.length >= 2) {
    const last = months[months.length - 1], prev = months[months.length - 2];
    const delta = last.total - prev.total;
    const pct   = Math.abs(Math.round(delta / prev.total * 100));
    items.push({ ic: delta > 0 ? '⬆️' : '⬇️', level: delta > prev.total * 0.2 ? 'warn' : 'ok',
      html: `${esc(last.label)} לעומת ${esc(prev.label)}: <strong>${delta >= 0 ? '+' : ''}₪${fmt(delta)}</strong> (${delta >= 0 ? '+' : ''}${pct}%)` });
  }

  // Top category
  if (cats.length) {
    const c = cats[0];
    const pct = Math.round(c.total / DATA.total_amount * 100);
    items.push({ ic: '🏆', level: pct > 50 ? 'warn' : 'ok',
      html: `הקטגוריה הדומיננטית: <strong>${esc(c.name)}</strong> — ₪${fmt(c.total)} (${pct}% מסה"כ)` });
  }

  // Subscription burden
  if (subs.length) {
    const subTotal = subs.reduce((s, r) => s + r.total, 0);
    const pct = Math.round(subTotal / DATA.total_amount * 100);
    items.push({ ic: '🔄', level: pct > 25 ? 'warn' : 'ok',
      html: `${subs.length} הוראות קבע — ₪${fmt(subTotal)} סה"כ (${pct}% מההוצאה השנתית)` });
  }

  // Top merchant
  if (top.length) {
    const m = top[0];
    const pct = Math.round(m.total / DATA.total_amount * 100);
    items.push({ ic: '🏪', level: 'ok',
      html: `בית עסק מוביל: <strong>${esc(m.name)}</strong> — ₪${fmt(m.total)} (${pct}%, ${m.count} עסקאות)` });
  }

  // Most frequent merchant
  const byCount = [...top].sort((a, b) => b.count - a.count);
  if (byCount.length && byCount[0].name !== top[0].name) {
    const m = byCount[0];
    items.push({ ic: '🔁', level: 'ok',
      html: `הכי הרבה ביקורים: <strong>${esc(m.name)}</strong> — ${m.count} עסקאות בשנה (ממוצע ₪${fmt(m.total / m.count)} לביקור)` });
  }

  // Installment future debt
  if (DATA.total_installment_remaining > 0) {
    items.push({ ic: '📋', level: DATA.total_installment_remaining > 5000 ? 'warn' : 'ok',
      html: `יתרת תשלומים עתידיים: <strong>₪${fmt(DATA.total_installment_remaining)}</strong> ב-${DATA.installments.length} תוכניות פתוחות` });
  }

  // Duplicates
  if (DATA.duplicates.length) {
    const dupTotal = DATA.duplicates.reduce((s, d) => s + d.total, 0);
    items.push({ ic: '⚠️', level: 'alert',
      html: `<strong>${DATA.duplicates.length} חיובים כפולים חשודים</strong> בסה"כ ₪${fmt(dupTotal)} — כדאי לבדוק` });
  }

  // Category volatility (highest month-to-month swing in any category)
  const mbc = DATA.monthly_by_category;
  if (mbc && mbc.months.length >= 2 && mbc.categories.length) {
    let maxSwing = 0, swingCat = '', swingFrom = 0, swingTo = 0;
    for (const cat of mbc.categories) {
      for (let i = 1; i < cat.data.length; i++) {
        const swing = Math.abs(cat.data[i] - cat.data[i - 1]);
        if (swing > maxSwing && (cat.data[i - 1] > 0 || cat.data[i] > 0)) {
          maxSwing = swing; swingCat = cat.name;
          swingFrom = cat.data[i - 1]; swingTo = cat.data[i];
        }
      }
    }
    if (swingCat) {
      const dir = swingTo > swingFrom ? 'עלייה' : 'ירידה';
      items.push({ ic: '📊', level: 'ok',
        html: `תנודתיות גבוהה ביותר: קטגוריה <strong>${esc(swingCat)}</strong> — ${dir} של ₪${fmt(maxSwing)} בחודש בודד` });
    }
  }

  const grid = document.getElementById('insights-grid');
  grid.innerHTML = items.length
    ? items.map(it => `<div class="insight-card ${it.level}"><span class="ic">${it.ic}</span><span class="body">${it.html}</span></div>`).join('')
    : `<div class="empty">אין תובנות זמינות.</div>`;
}

document.querySelectorAll('.section > h2').forEach(h => {
  h.addEventListener('click', () => h.parentElement.classList.toggle('collapsed'));
});

initTheme();
renderSrc();
renderCards();
renderYearlyInsights();
renderMonths();
renderMerchants();
renderSubs();
renderInstallments();
renderDuplicates();
renderCharts();
</script>
</body>
</html>
"""


def generate_yearly_html(summary: dict) -> str:
    payload = json.dumps(summary, ensure_ascii=False)
    return (
        YEARLY_HTML_TEMPLATE
        .replace("__YEAR__", str(summary["year"]))
        .replace("__DATA__", payload)
    )


def _file_fingerprint(folder: Path) -> tuple:
    """Cheap hash of xlsx files in folder so we can detect changes in --watch mode."""
    items = []
    for p in sorted(folder.glob("*.xlsx")):
        if p.name.startswith("~$"):
            continue
        try:
            st = p.stat()
            items.append((p.name, st.st_size, int(st.st_mtime)))
        except OSError:
            continue
    return tuple(items)


def process_folder(folder: Path, year_filter: int | None = None) -> list[Path]:
    """Scan the folder and write one yearly_summary_YYYY.html per year."""
    years = scan_folder(folder, year_filter=year_filter)
    if not years:
        print(f"  no recognized transactions found in {folder}")
        return []

    written = []
    for year in sorted(years):
        summary = build_yearly_summary(years[year])
        out = folder / f"yearly_summary_{year}.html"
        out.write_text(generate_yearly_html(summary), encoding="utf-8")
        written.append(out)
        months = len(summary["months"])
        print(
            f"  {year}: {summary['total_count']:4d} payments, "
            f"₪{summary['total_amount']:,.2f} across {months} month(s) → {out.name}"
        )
    return written


def main() -> int:
    args = sys.argv[1:]
    if not args or "-h" in args or "--help" in args:
        print(__doc__)
        return 1

    positional = [a for a in args if not a.startswith("-")]
    if not positional:
        print("error: folder path required", file=sys.stderr)
        return 1

    folder = Path(positional[0])
    if not folder.is_dir():
        print(f"error: not a directory: {folder}", file=sys.stderr)
        return 1

    year_filter = None
    if "--year" in args:
        idx = args.index("--year")
        if idx + 1 >= len(args):
            print("error: --year requires a value", file=sys.stderr)
            return 1
        year_filter = int(args[idx + 1])

    interval = 5.0
    if "--interval" in args:
        idx = args.index("--interval")
        if idx + 1 >= len(args):
            print("error: --interval requires a value", file=sys.stderr)
            return 1
        interval = float(args[idx + 1])

    print(f"scanning: {folder.resolve()}")
    written = process_folder(folder, year_filter=year_filter)

    if "--open" in args and written:
        webbrowser.open(written[-1].resolve().as_uri())

    if "--watch" in args:
        print(f"watching for changes (interval {interval}s, Ctrl-C to stop)...")
        last = _file_fingerprint(folder)
        try:
            while True:
                time.sleep(interval)
                current = _file_fingerprint(folder)
                if current != last:
                    last = current
                    print(f"change detected — re-scanning")
                    process_folder(folder, year_filter=year_filter)
        except KeyboardInterrupt:
            print("\nstopped.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
