#!/usr/bin/env python3
"""
bank_ui.py — parse a bank account statement (First International / Fibi)
from XLS or PDF and generate a rich income-vs-expense HTML dashboard.

Supported sources:
  * FibiSave*.xls   — Excel export from online banking
  * FibiSave*.pdf   — PDF export from online banking
"""
from __future__ import annotations

import json
import re
from collections import defaultdict
from datetime import datetime
from pathlib import Path

# ---------------------------------------------------------------------------
# Category classification
# ---------------------------------------------------------------------------

# Each entry: (pattern, category_he, direction)
# direction: 'income' | 'expense' | 'internal' | 'savings'
_CAT_RULES: list[tuple[re.Pattern, str, str]] = [
    (re.compile(r'משכנתא',        re.I), 'משכנתא',          'expense'),
    (re.compile(r'בינלאומי.משכנת', re.I), 'משכנתא',          'expense'),
    (re.compile(r'כרטיסי אשראי',  re.I), 'כרטיסי אשראי',   'expense'),
    (re.compile(r'עפ.י הרשאה כאל', re.I), 'כרטיסי אשראי',  'expense'),
    (re.compile(r'מקס איט',       re.I), 'כרטיסי אשראי',   'expense'),
    (re.compile(r'הלוואה.תשלום',  re.I), 'הלוואות',         'expense'),
    (re.compile(r'הלוואה',        re.I), 'הלוואות',         'expense'),
    (re.compile(r'ריבית על הלוואה',re.I),'הלוואות',         'expense'),
    (re.compile(r'פקדון|פק קרן|פקמש|חידוש פיקדון|פירעון ריבית פיקדון', re.I), 'פיקדונות וחסכונות', 'savings'),
    (re.compile(r'הפניקס ביטוח',  re.I), 'ביטוח',           'expense'),
    (re.compile(r'מנורה',         re.I), 'ביטוח',           'expense'),
    (re.compile(r'שלמה פסגה',     re.I), 'ביטוח רכב',      'expense'),
    (re.compile(r'ביטוח לאומי',   re.I), 'ביטוח לאומי',    'income'),
    (re.compile(r'קצבת ילדים',    re.I), 'קצבאות',          'income'),
    (re.compile(r'מס.הכנסה החזר', re.I), 'החזר מס',         'income'),
    (re.compile(r'ביוטיק|biotek|biotic', re.I), 'משכורת',   'income'),
    (re.compile(r'אלו.ט.אגודה',   re.I), 'הכנסה מנכס',     'income'),
    (re.compile(r'זיכוי מביט|זיכויממביט', re.I), 'העברה נכנסת', 'income'),
    (re.compile(r'זיכוי מפייבוקס|זיכויפייבוקס|PAYBOX', re.I), 'העברה נכנסת', 'income'),
    (re.compile(r'זיכוי מיידי',   re.I), 'העברה נכנסת',    'income'),
    (re.compile(r'זיכוי מב\.',    re.I), 'העברה נכנסת',    'income'),
    (re.compile(r'זיכוי ממרכנת',  re.I), 'העברה נכנסת',    'income'),
    (re.compile(r'זיכוי מדיסקונט', re.I),'העברה נכנסת',    'income'),
    (re.compile(r'זיכוי מבנק',    re.I), 'העברה נכנסת',    'income'),
    (re.compile(r'זיכוי מבל',     re.I), 'העברה נכנסת',    'income'),
    (re.compile(r'^זיכוי$',       re.I), 'העברה נכנסת',    'income'),
    (re.compile(r'אלקטרה פאוור',  re.I), 'שכ"ד מנכס',      'income'),
    (re.compile(r'העברה מהחשבון', re.I), 'העברה יוצאת',    'expense'),
    (re.compile(r'משיכת שיק',     re.I), 'צ\'קים',          'expense'),
    (re.compile(r'הפקדת שיק',     re.I), 'צ\'קים',          'income'),
    (re.compile(r'הפקדת מזומן|סניפומט.*הפקד', re.I), 'הפקדת מזומן', 'income'),
    (re.compile(r'סניפומט|כספומט', re.I), 'משיכת מזומן',   'expense'),
    (re.compile(r'מע.זה.ב|מועזה', re.I), 'מס על מקרקעין', 'expense'),
    (re.compile(r'המועצה להסדר',  re.I), 'מיסים ואגרות',   'expense'),
]

# סוג פעולה codes that are internal bank movements (deposits, PK, etc.)
_SAVINGS_CODES = {'341', '343', '391', '393', '240', '295', '495'}


def _classify(desc: str, sof: str = '') -> tuple[str, str]:
    """Return (category_he, direction) for a transaction."""
    for pattern, cat, direction in _CAT_RULES:
        if pattern.search(desc):
            return cat, direction
    # fallback by transaction type code
    if sof in _SAVINGS_CODES:
        return 'פיקדונות וחסכונות', 'savings'
    return 'אחר', 'expense'


# ---------------------------------------------------------------------------
# XLS parser
# ---------------------------------------------------------------------------

def _parse_xls(path: Path) -> dict:
    import xlrd
    wb = xlrd.open_workbook(str(path))
    ws = wb.sheet_by_index(0)

    def xl_date(val) -> str | None:
        if not val or str(val).strip() in ('', ' '):
            return None
        try:
            return datetime(*xlrd.xldate_as_tuple(float(val), wb.datemode)).strftime('%Y-%m-%d')
        except Exception:
            return None

    def parse_num(s) -> float:
        s = str(s).replace(',', '').strip()
        if not s or s == ' ':
            return 0.0
        try:
            return float(s)
        except Exception:
            return 0.0

    # Find opening balance from row 5 (יתרת פתיחה)
    opening_balance = None
    for r in range(3, 8):
        row = [ws.cell_value(r, c) for c in range(ws.ncols)]
        if 'פתיחה' in str(row[5]) or 'פתיחה' in str(row[4]):
            try:
                opening_balance = float(str(row[1]).replace(',', '').strip())
            except Exception:
                pass
            break

    transactions = []
    for r in range(5, ws.nrows):
        row = [ws.cell_value(r, c) for c in range(ws.ncols)]
        date = xl_date(row[8])
        if not date:
            continue
        desc    = str(row[5]).strip()
        credit  = parse_num(row[3])
        debit   = parse_num(row[4])
        balance = parse_num(row[1]) if str(row[1]).strip() not in ('', ' ') else None
        sof     = str(int(float(row[7]))) if str(row[7]).strip() not in ('', ' ') else ''

        cat, direction = _classify(desc, sof)
        transactions.append({
            'date':      date,
            'desc':      desc,
            'credit':    credit,
            'debit':     debit,
            'balance':   balance,
            'category':  cat,
            'direction': direction,
        })

    closing_balance = None
    for t in reversed(transactions):
        if t['balance'] is not None:
            closing_balance = t['balance']
            break

    return {
        'source': path.name,
        'title': path.stem,
        'opening_balance': opening_balance,
        'closing_balance': closing_balance,
        'transactions': transactions,
    }


# ---------------------------------------------------------------------------
# PDF parser
# ---------------------------------------------------------------------------

def _fix_heb(s: str) -> str:
    """Reverse reversed-Hebrew text from PDF extraction."""
    if not s or not re.search(r'[\u05d0-\u05ea]', s):
        return s
    words = s.split()
    fixed = []
    for w in words:
        if re.search(r'[\u05d0-\u05ea]', w):
            fixed.append(w[::-1])
        else:
            fixed.append(w)
    return ' '.join(reversed(fixed))


def _parse_pdf(path: Path) -> dict:
    import pdfplumber
    DATE_RE = re.compile(r'^\d{2}/\d{2}/\d{4}$')
    NUM_RE  = re.compile(r'^\.?\d[\d,]*\.?\d*$')

    opening_balance = None
    transactions = []

    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if not row or not row[0]:
                        continue
                    cell0 = row[0].strip()

                    # Opening balance row
                    if len(row) > 3 and row[3] and not DATE_RE.match(cell0):
                        try:
                            val = float(str(row[3]).replace(',', ''))
                            # Only pick up if description mentions opening balance
                            desc_raw = row[6] if len(row) > 6 else ''
                            if 'החיתפ' in str(desc_raw) or 'פתיחה' in str(desc_raw):
                                opening_balance = val
                        except Exception:
                            pass
                        continue

                    if not DATE_RE.match(cell0):
                        continue

                    try:
                        d, m, y = cell0.split('/')
                        date_iso = f'{y}-{m}-{d}'
                        sof      = row[1].strip() if len(row) > 1 else ''
                        balance_raw = row[3].strip() if len(row) > 3 else ''
                        debit_raw   = row[4].strip() if len(row) > 4 else ''
                        credit_raw  = row[5].strip() if len(row) > 5 else ''
                        desc_raw    = row[6].strip() if len(row) > 6 else ''

                        def pn(s):
                            s = s.replace(',', '').strip()
                            if not s:
                                return 0.0
                            try:
                                return float(s)
                            except Exception:
                                return 0.0

                        credit  = pn(credit_raw)
                        debit   = pn(debit_raw)
                        balance = pn(balance_raw) if balance_raw and NUM_RE.match(balance_raw.replace(',', '')) else None
                        desc    = _fix_heb(desc_raw)

                        cat, direction = _classify(desc, sof)
                        transactions.append({
                            'date':      date_iso,
                            'desc':      desc,
                            'credit':    credit,
                            'debit':     debit,
                            'balance':   balance,
                            'category':  cat,
                            'direction': direction,
                        })
                    except Exception:
                        continue

    closing_balance = None
    for t in reversed(transactions):
        if t['balance'] is not None:
            closing_balance = t['balance']
            break

    return {
        'source': path.name,
        'title': path.stem,
        'opening_balance': opening_balance,
        'closing_balance': closing_balance,
        'transactions': transactions,
    }


# ---------------------------------------------------------------------------
# Public parse entry point
# ---------------------------------------------------------------------------

def parse_bank_statement(path: Path) -> dict:
    """Parse a bank statement file (XLS or PDF) and return structured data."""
    ext = path.suffix.lower()
    if ext in ('.xls', '.xlsx'):
        raw = _parse_xls(path)
    elif ext == '.pdf':
        raw = _parse_pdf(path)
    else:
        raise ValueError(f"Unsupported file type: {ext}")

    transactions = raw['transactions']

    # Mark credit-side savings (deposit redemptions like פק קרן) as 'internal'
    # so they don't inflate or deflate any totals — they're just your own money returning.
    # Only debit-side savings (new money going into deposits) count as savings outflow.
    for t in transactions:
        if t['direction'] == 'savings' and t['credit'] > 0 and t['debit'] == 0:
            t['direction'] = 'internal'

    # Monthly summaries
    monthly: dict = defaultdict(lambda: {'income': 0.0, 'expense': 0.0, 'savings': 0.0, 'count': 0})
    for t in transactions:
        ym = t['date'][:7]
        if t['direction'] == 'income':
            monthly[ym]['income'] += t['credit']
        elif t['direction'] == 'expense':
            monthly[ym]['expense'] += t['debit']
        elif t['direction'] == 'savings':
            monthly[ym]['savings'] += t['debit']
        monthly[ym]['count'] += 1

    months = sorted([
        {
            'ym':      ym,
            'label':   _month_label(ym),
            'income':  round(d['income'], 2),
            'expense': round(d['expense'], 2),
            'savings': round(max(d['savings'], 0), 2),
            'net':     round(d['income'] - d['expense'], 2),
            'count':   d['count'],
        }
        for ym, d in monthly.items()
    ], key=lambda x: x['ym'])

    # Category breakdown (income)
    income_cats: dict = defaultdict(float)
    expense_cats: dict = defaultdict(float)
    for t in transactions:
        if t['direction'] == 'income':
            income_cats[t['category']] += t['credit']
        elif t['direction'] == 'expense':
            expense_cats[t['category']] += t['debit']

    income_by_cat  = sorted([{'name': k, 'total': round(v, 2)} for k, v in income_cats.items()],  key=lambda x: -x['total'])
    expense_by_cat = sorted([{'name': k, 'total': round(v, 2)} for k, v in expense_cats.items()], key=lambda x: -x['total'])

    total_income  = round(sum(t['credit'] for t in transactions if t['direction'] == 'income'),  2)
    total_expense = round(sum(t['debit']  for t in transactions if t['direction'] == 'expense'), 2)
    total_savings = round(sum(t['debit']  for t in transactions if t['direction'] == 'savings'), 2)

    # Balance trend (only rows with a balance value)
    balance_trend = [
        {'date': t['date'], 'balance': t['balance']}
        for t in transactions if t['balance'] is not None
    ]

    # Top income sources
    income_sources: dict = defaultdict(float)
    for t in transactions:
        if t['direction'] == 'income':
            income_sources[t['desc'][:40]] += t['credit']
    top_income = sorted([{'name': k, 'total': round(v, 2)} for k, v in income_sources.items()], key=lambda x: -x['total'])[:15]

    # Top expense recipients
    expense_recipients: dict = defaultdict(float)
    for t in transactions:
        if t['direction'] == 'expense':
            expense_recipients[t['desc'][:40]] += t['debit']
    top_expense = sorted([{'name': k, 'total': round(v, 2)} for k, v in expense_recipients.items()], key=lambda x: -x['total'])[:15]

    return {
        'source':           raw['source'],
        'title':            raw['title'],
        'opening_balance':  raw['opening_balance'],
        'closing_balance':  raw['closing_balance'],
        'total_income':     total_income,
        'total_expense':    total_expense,
        'total_savings':    total_savings,
        'net':              round(total_income - total_expense, 2),
        'months':           months,
        'income_by_cat':    income_by_cat,
        'expense_by_cat':   expense_by_cat,
        'balance_trend':    balance_trend,
        'top_income':       top_income,
        'top_expense':      top_expense,
        'transactions':     transactions,
    }


def _month_label(ym: str) -> str:
    months = ['ינואר','פברואר','מרץ','אפריל','מאי','יוני',
              'יולי','אוגוסט','ספטמבר','אוקטובר','נובמבר','דצמבר']
    y, m = ym.split('-')
    return f"{months[int(m)-1]} {y}"


# ---------------------------------------------------------------------------
# HTML generation
# ---------------------------------------------------------------------------

BANK_HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="he" dir="rtl" data-theme="light">
<head>
<meta charset="UTF-8">
<title>תנועות בחשבון</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
  :root{
    --bg:#f5f5f7;--card:#fff;--text:#222;--muted:#666;--soft:#888;
    --border:#eee;--border-strong:#ddd;--hover:#f8f8fb;--th-bg:#fafafa;
    --shadow:0 1px 3px rgba(0,0,0,.08);
    --primary:#2196f3;--income:#2e7d32;--expense:#c62828;--savings:#6a1b9a;
    --income-bg:#e8f5e9;--expense-bg:#ffebee;--savings-bg:#f3e5f5;
  }
  [data-theme="dark"]{
    --bg:#111418;--card:#1c2128;--text:#e6edf3;--muted:#9aa4af;--soft:#7a8591;
    --border:#2a323c;--border-strong:#394350;--hover:#232b36;--th-bg:#1a2028;
    --shadow:0 1px 3px rgba(0,0,0,.4);
    --primary:#64b5f6;--income:#66bb6a;--expense:#ef5350;--savings:#ce93d8;
    --income-bg:#1b4d20;--expense-bg:#4a1212;--savings-bg:#3a1c4a;
  }
  *{box-sizing:border-box;}
  body{font-family:-apple-system,"Segoe UI",Arial,sans-serif;margin:0;padding:24px;
       background:var(--bg);color:var(--text);transition:background .2s,color .2s;}
  header{display:flex;justify-content:space-between;align-items:center;gap:16px;margin-bottom:20px;}
  header h1{font-size:20px;margin:0 0 3px;}
  .sub{font-size:12px;color:var(--soft);}
  .btn{background:var(--card);color:var(--text);border:1px solid var(--border-strong);
       height:36px;border-radius:8px;cursor:pointer;font-size:13px;font-weight:600;
       padding:0 14px;box-shadow:var(--shadow);text-decoration:none;display:inline-flex;align-items:center;gap:6px;}
  .btn:hover{background:var(--hover);}
  .cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px;margin-bottom:20px;}
  .card{background:var(--card);padding:16px;border-radius:10px;box-shadow:var(--shadow);border-right:4px solid transparent;}
  .card.income-card{border-color:var(--income);}
  .card.expense-card{border-color:var(--expense);}
  .card.savings-card{border-color:var(--savings);}
  .card.net-card{border-color:var(--primary);}
  .card .lbl{font-size:11px;color:var(--muted);letter-spacing:.5px;}
  .card .val{font-size:22px;font-weight:700;margin-top:4px;}
  .card .val.income{color:var(--income);}
  .card .val.expense{color:var(--expense);}
  .card .val.savings{color:var(--savings);}
  .card .val.primary{color:var(--primary);}
  .card .sub2{font-size:11px;color:var(--soft);margin-top:3px;}
  .section{background:var(--card);padding:16px 20px;border-radius:10px;box-shadow:var(--shadow);margin-bottom:20px;}
  .section>h2{margin:0 0 14px;font-size:16px;cursor:pointer;user-select:none;}
  .section.collapsed>:not(h2){display:none;}
  .section>h2::before{content:"▾";display:inline-block;width:1em;font-size:11px;color:var(--muted);transform:scaleX(-1);}
  .section.collapsed>h2::before{content:"▸";}
  .charts-row{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px;}
  @media(max-width:700px){.charts-row{grid-template-columns:1fr;}}
  .chart-card{background:var(--card);padding:16px 20px;border-radius:10px;box-shadow:var(--shadow);}
  .chart-card h2{margin:0 0 12px;font-size:15px;}
  .chart-wrap{position:relative;height:280px;}
  .chart-wrap.tall{height:320px;}
  table{width:100%;border-collapse:collapse;font-size:13px;}
  th,td{padding:7px 10px;border-bottom:1px solid var(--border);text-align:right;white-space:nowrap;}
  th{background:var(--th-bg);font-weight:600;font-size:12px;position:sticky;top:0;z-index:2;}
  tr:hover td{background:var(--hover);}
  .num{font-variant-numeric:tabular-nums;}
  .credit{color:var(--income);font-weight:600;}
  .debit{color:var(--expense);font-weight:600;}
  .tbl-wrap{overflow-x:auto;}
  .filter{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px;}
  .filter input,.filter select{padding:7px 12px;border:1px solid var(--border-strong);border-radius:6px;
    font-size:13px;font-family:inherit;background:var(--card);color:var(--text);}
  .filter input{min-width:200px;}
  .badge{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600;}
  .badge-income{background:var(--income-bg);color:var(--income);}
  .badge-expense{background:var(--expense-bg);color:var(--expense);}
  .badge-savings{background:var(--savings-bg);color:var(--savings);}
  .badge-internal{background:var(--th-bg);color:var(--muted);}
  .sum-row td{border-top:2px solid var(--border-strong);border-bottom:none;font-weight:700;background:var(--th-bg);}
  .empty{color:var(--soft);font-style:italic;padding:8px 0;}
  .insights-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(270px,1fr));gap:10px;}
  .ic{display:flex;align-items:flex-start;gap:10px;padding:10px 14px;border-radius:8px;
      border-right:4px solid transparent;background:var(--bg);font-size:14px;line-height:1.5;}
  .ic.ok{border-color:#43a047;}.ic.warn{border-color:#fb8c00;}.ic.alert{border-color:var(--expense);}
  .ic .emoji{font-size:20px;flex-shrink:0;margin-top:1px;}
  .ic .body strong{color:var(--primary);font-weight:700;}
</style>
</head>
<body>
<header>
  <div>
    <h1 id="title">תנועות בחשבון</h1>
    <div class="sub" id="sub"></div>
  </div>
  <div style="display:flex;gap:8px;">
    <a href="/" class="btn">← קובץ חדש</a>
    <a href="/bank" class="btn">📂 דף חשבון חדש</a>
    <button class="btn" id="btn-save">💾 שמור HTML</button>
    <button class="btn" id="theme-toggle">🌙</button>
  </div>
</header>

<div class="cards" id="cards"></div>

<div class="charts-row">
  <div class="chart-card"><h2>הכנסות vs הוצאות לפי חודש</h2><div class="chart-wrap tall"><canvas id="chart-monthly"></canvas></div></div>
  <div class="chart-card"><h2>יתרת חשבון לאורך זמן</h2><div class="chart-wrap tall"><canvas id="chart-balance"></canvas></div></div>
</div>

<div class="charts-row">
  <div class="chart-card"><h2>פילוח הכנסות לפי קטגוריה</h2><div class="chart-wrap"><canvas id="chart-income-cat"></canvas></div></div>
  <div class="chart-card"><h2>פילוח הוצאות לפי קטגוריה</h2><div class="chart-wrap"><canvas id="chart-expense-cat"></canvas></div></div>
</div>

<div class="section">
  <h2>תובנות 🧠</h2>
  <div class="insights-grid" id="insights-grid"></div>
</div>

<div class="section">
  <h2>סיכום חודשי</h2>
  <table id="monthly-table">
    <thead><tr>
      <th>חודש</th><th>הכנסות (₪)</th><th>הוצאות (₪)</th><th>חיסכון/פיקדון (₪)</th><th>נטו (₪)</th><th>עסקאות</th>
    </tr></thead>
    <tbody id="monthly-tbody"></tbody>
  </table>
</div>

<div class="section">
  <h2>Top מקורות הכנסה</h2>
  <table id="income-table">
    <thead><tr><th>תיאור</th><th>סה"כ (₪)</th></tr></thead>
    <tbody id="income-tbody"></tbody>
  </table>
</div>

<div class="section">
  <h2>Top יעדי הוצאה</h2>
  <table id="expense-table">
    <thead><tr><th>תיאור</th><th>סה"כ (₪)</th></tr></thead>
    <tbody id="expense-tbody"></tbody>
  </table>
</div>

<div class="section" id="sec-all">
  <h2>כל התנועות</h2>
  <div class="filter">
    <input id="search" type="text" placeholder="חיפוש תיאור...">
    <select id="dir-filter">
      <option value="">כל הכיוונים</option>
      <option value="income">הכנסות</option>
      <option value="expense">הוצאות</option>
      <option value="savings">פיקדונות</option>
      <option value="internal">פנימי (פק קרן)</option>
    </select>
    <select id="cat-filter"><option value="">כל הקטגוריות</option></select>
  </div>
  <div class="tbl-wrap">
    <table id="all-table">
      <thead><tr>
        <th>תאריך</th><th>תיאור</th><th>קטגוריה</th><th>זכות (₪)</th><th>חובה (₪)</th><th>יתרה (₪)</th>
      </tr></thead>
      <tbody id="all-tbody"></tbody>
    </table>
  </div>
</div>

<script>
const DATA = __DATA__;
const fmt  = n => new Intl.NumberFormat('he-IL',{minimumFractionDigits:2,maximumFractionDigits:2}).format(n);
const fmt0 = n => new Intl.NumberFormat('he-IL',{maximumFractionDigits:0}).format(n);
const esc  = s => String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const COLORS=['#2196f3','#ef6c00','#43a047','#8e24aa','#e53935','#00acc1','#fb8c00','#6d4c41','#546e7a','#d81b60','#7cb342','#3949ab'];

let charts={};
function destroyCharts(){Object.values(charts).forEach(c=>c&&c.destroy());charts={};}

function applyTheme(t){
  document.documentElement.dataset.theme=t;
  document.getElementById('theme-toggle').textContent=t==='dark'?'☀️':'🌙';
  renderCharts();
}
function initTheme(){
  const saved=localStorage.getItem('payments-theme')||(window.matchMedia('(prefers-color-scheme:dark)').matches?'dark':'light');
  applyTheme(saved);
  document.getElementById('theme-toggle').addEventListener('click',()=>{
    const next=document.documentElement.dataset.theme==='dark'?'light':'dark';
    localStorage.setItem('payments-theme',next);applyTheme(next);
  });
}

// ── Header ────────────────────────────────────────────────────────────────
document.getElementById('title').textContent='תנועות בחשבון';
document.getElementById('sub').textContent=`מקור: ${esc(DATA.source)}`;

// ── Cards ─────────────────────────────────────────────────────────────────
function renderCards(){
  const netClass=DATA.net>=0?'income':'expense';
  const items=[
    {l:'סה"כ הכנסות',v:'₪'+fmt(DATA.total_income),cls:'income-card',vc:'income',sub:`${DATA.months.length} חודשים`},
    {l:'סה"כ הוצאות',v:'₪'+fmt(DATA.total_expense),cls:'expense-card',vc:'expense',sub:''},
    {l:'פיקדונות וחיסכון',v:'₪'+fmt(DATA.total_savings),cls:'savings-card',vc:'savings',sub:''},
    {l:'נטו (הכנסות − הוצאות)',v:(DATA.net>=0?'+':'')+'₪'+fmt(DATA.net),cls:'net-card',vc:netClass,sub:''},
    {l:'יתרת פתיחה',v:'₪'+fmt(DATA.opening_balance||0),cls:'',vc:'primary',sub:''},
    {l:'יתרת סגירה',v:'₪'+fmt(DATA.closing_balance||0),cls:'',vc:'primary',sub:''},
  ];
  document.getElementById('cards').innerHTML=items.map(it=>
    `<div class="card ${it.cls}"><div class="lbl">${esc(it.l)}</div><div class="val ${it.vc}">${it.v}</div><div class="sub2">${it.sub}</div></div>`
  ).join('');
}

// ── Charts ─────────────────────────────────────────────────────────────────
function renderCharts(){
  if(typeof Chart==='undefined') return;
  destroyCharts();
  const text=getComputedStyle(document.documentElement).getPropertyValue('--text').trim()||'#222';
  const grid=getComputedStyle(document.documentElement).getPropertyValue('--border').trim()||'#eee';
  Chart.defaults.color=text; Chart.defaults.borderColor=grid;

  // Monthly income vs expense grouped bar
  const ml=DATA.months.map(m=>m.label);
  charts.monthly=new Chart(document.getElementById('chart-monthly'),{
    type:'bar',
    data:{labels:ml,datasets:[
      {label:'הכנסות',data:DATA.months.map(m=>m.income),backgroundColor:'rgba(46,125,50,.8)'},
      {label:'הוצאות',data:DATA.months.map(m=>m.expense),backgroundColor:'rgba(198,40,40,.8)'},
    ]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'top'},tooltip:{callbacks:{label:ctx=>` ${ctx.dataset.label}: ₪${fmt0(ctx.parsed.y)}`}}},
      scales:{x:{ticks:{font:{size:10}}},y:{ticks:{callback:v=>'₪'+fmt0(v),font:{size:10}}}},
    },
  });

  // Balance trend line
  const bt=DATA.balance_trend;
  charts.balance=new Chart(document.getElementById('chart-balance'),{
    type:'line',
    data:{labels:bt.map(b=>b.date.slice(5)),datasets:[{
      data:bt.map(b=>b.balance),borderColor:'#2196f3',backgroundColor:'rgba(33,150,243,.1)',
      tension:.3,fill:true,pointRadius:0,
    }]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>' ₪'+fmt(ctx.parsed.y)}}},
      scales:{x:{ticks:{font:{size:9},maxTicksLimit:12}},y:{ticks:{callback:v=>'₪'+fmt0(v),font:{size:10}}}},
    },
  });

  // Income doughnut
  const ic=DATA.income_by_cat.slice(0,10);
  charts.incomeCat=new Chart(document.getElementById('chart-income-cat'),{
    type:'doughnut',
    data:{labels:ic.map(c=>c.name),datasets:[{data:ic.map(c=>c.total),backgroundColor:COLORS,borderWidth:0}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'bottom',labels:{boxWidth:12,font:{size:11}}},
        tooltip:{callbacks:{label:ctx=>` ${ctx.label}: ₪${fmt(ctx.parsed)}`}}},
    },
  });

  // Expense doughnut
  const ec=DATA.expense_by_cat.slice(0,10);
  charts.expenseCat=new Chart(document.getElementById('chart-expense-cat'),{
    type:'doughnut',
    data:{labels:ec.map(c=>c.name),datasets:[{data:ec.map(c=>c.total),backgroundColor:COLORS,borderWidth:0}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'bottom',labels:{boxWidth:12,font:{size:11}}},
        tooltip:{callbacks:{label:ctx=>` ${ctx.label}: ₪${fmt(ctx.parsed)}`}}},
    },
  });
}

// ── Monthly table ──────────────────────────────────────────────────────────
function renderMonthly(){
  const ti=DATA.months.reduce((s,m)=>s+m.income,0);
  const te=DATA.months.reduce((s,m)=>s+m.expense,0);
  const ts=DATA.months.reduce((s,m)=>s+m.savings,0);
  const tn=DATA.months.reduce((s,m)=>s+m.net,0);
  const tc=DATA.months.reduce((s,m)=>s+m.count,0);
  document.getElementById('monthly-tbody').innerHTML=
    DATA.months.map(m=>`<tr>
      <td>${esc(m.label)}</td>
      <td class="num credit">${m.income>0?fmt(m.income):'—'}</td>
      <td class="num debit">${m.expense>0?fmt(m.expense):'—'}</td>
      <td class="num" style="color:var(--savings)">${m.savings>0?fmt(m.savings):'—'}</td>
      <td class="num ${m.net>=0?'credit':'debit'}">${(m.net>=0?'+':'')+'₪'+fmt(Math.abs(m.net))}</td>
      <td class="num">${m.count}</td>
    </tr>`).join('')+
    `<tr class="sum-row"><td>סה"כ</td><td class="num credit">${fmt(ti)}</td><td class="num debit">${fmt(te)}</td>
     <td class="num">${fmt(ts)}</td><td class="num ${tn>=0?'credit':'debit'}">${(tn>=0?'+':'')+'₪'+fmt(Math.abs(tn))}</td>
     <td class="num">${tc}</td></tr>`;
}

// ── Top income / expense tables ────────────────────────────────────────────
function renderTopTables(){
  document.getElementById('income-tbody').innerHTML=
    DATA.top_income.map(t=>`<tr><td>${esc(t.name)}</td><td class="num credit">${fmt(t.total)}</td></tr>`).join('');
  document.getElementById('expense-tbody').innerHTML=
    DATA.top_expense.map(t=>`<tr><td>${esc(t.name)}</td><td class="num debit">${fmt(t.total)}</td></tr>`).join('');
}

// ── All transactions table ─────────────────────────────────────────────────
function initFilters(){
  const cats=[...new Set(DATA.transactions.map(t=>t.category).filter(Boolean))].sort();
  document.getElementById('cat-filter').insertAdjacentHTML('beforeend',
    cats.map(c=>`<option>${esc(c)}</option>`).join(''));
}

function renderAll(){
  const q=document.getElementById('search').value.trim().toLowerCase();
  const df=document.getElementById('dir-filter').value;
  const cf=document.getElementById('cat-filter').value;

  let rows=DATA.transactions.filter(t=>{
    if(q && !t.desc.toLowerCase().includes(q)) return false;
    if(df && t.direction!==df) return false;
    if(cf && t.category!==cf) return false;
    return true;
  });

  const dirLabel={income:'הכנסה',expense:'הוצאה',savings:'פיקדון'};
  document.getElementById('all-tbody').innerHTML=rows.map(t=>`<tr>
    <td class="num">${esc(t.date)}</td>
    <td>${esc(t.desc)}</td>
    <td><span class="badge badge-${t.direction}">${esc(t.category)}</span></td>
    <td class="num ${t.credit>0?'credit':''}">${t.credit>0?fmt(t.credit):'—'}</td>
    <td class="num ${t.debit>0?'debit':''}">${t.debit>0?fmt(t.debit):'—'}</td>
    <td class="num">${t.balance!=null?fmt(t.balance):'—'}</td>
  </tr>`).join('');
}

// ── Insights ──────────────────────────────────────────────────────────────
function renderInsights(){
  const items=[];
  const months=DATA.months;
  const n=months.length;

  // Savings rate
  const sr=DATA.total_income>0?Math.round(DATA.total_savings/DATA.total_income*100):0;
  items.push({l:sr<10?'warn':'ok',e:'🏦',
    h:`שיעור חיסכון: <strong>${sr}%</strong> מההכנסות הופנו לפיקדונות (₪${fmt(DATA.total_savings)})`});

  // Net per month
  const avgNet=n?DATA.net/n:0;
  items.push({l:avgNet<0?'alert':'ok',e:avgNet>=0?'✅':'⚠️',
    h:`נטו ממוצע לחודש: <strong>${avgNet>=0?'+':''}₪${fmt(avgNet)}</strong> (הכנסות פחות הוצאות בלי פיקדונות)`});

  // Biggest income month
  const peakIncome=[...months].sort((a,b)=>b.income-a.income)[0];
  if(peakIncome) items.push({l:'ok',e:'📈',
    h:`חודש הכנסה גבוה ביותר: <strong>${esc(peakIncome.label)}</strong> — ₪${fmt(peakIncome.income)}`});

  // Biggest expense month
  const peakExpense=[...months].sort((a,b)=>b.expense-a.expense)[0];
  if(peakExpense) items.push({l:'warn',e:'📉',
    h:`חודש הוצאה גבוה ביותר: <strong>${esc(peakExpense.label)}</strong> — ₪${fmt(peakExpense.expense)}`});

  // Top income category
  if(DATA.income_by_cat.length){
    const top=DATA.income_by_cat[0];
    const pct=Math.round(top.total/DATA.total_income*100);
    items.push({l:'ok',e:'🏆',
      h:`מקור הכנסה מוביל: <strong>${esc(top.name)}</strong> — ₪${fmt(top.total)} (${pct}% מסה"כ הכנסות)`});
  }

  // Top expense category
  if(DATA.expense_by_cat.length){
    const top=DATA.expense_by_cat[0];
    const pct=Math.round(top.total/DATA.total_expense*100);
    items.push({l:pct>40?'warn':'ok',e:'💸',
      h:`הוצאה מובילה: <strong>${esc(top.name)}</strong> — ₪${fmt(top.total)} (${pct}% מסה"כ הוצאות)`});
  }

  // Mortgage burden
  const mortgage=DATA.expense_by_cat.find(c=>c.name==='משכנתא');
  if(mortgage){
    const pct=Math.round(mortgage.total/DATA.total_income*100);
    items.push({l:pct>35?'warn':'ok',e:'🏠',
      h:`עלות משכנתא: <strong>₪${fmt(mortgage.total)}</strong> — ${pct}% מסה"כ ההכנסות`});
  }

  // Credit card burden
  const cc=DATA.expense_by_cat.find(c=>c.name==='כרטיסי אשראי');
  if(cc){
    const pct=Math.round(cc.total/DATA.total_expense*100);
    items.push({l:pct>50?'warn':'ok',e:'💳',
      h:`כרטיסי אשראי: <strong>₪${fmt(cc.total)}</strong> — ${pct}% מסה"כ ההוצאות`});
  }

  // Balance change
  if(DATA.opening_balance!=null && DATA.closing_balance!=null){
    const delta=DATA.closing_balance-DATA.opening_balance;
    items.push({l:delta<0?'warn':'ok',e:delta>=0?'📊':'📊',
      h:`יתרה <strong>${delta>=0?'עלתה':'ירדה'}</strong> ב-₪${fmt(Math.abs(delta))} מתחילת התקופה (${fmt(DATA.opening_balance)} ← ${fmt(DATA.closing_balance)})`});
  }

  // Insurance total
  const ins=(DATA.expense_by_cat.find(c=>c.name==='ביטוח')||{total:0}).total +
            (DATA.expense_by_cat.find(c=>c.name==='ביטוח רכב')||{total:0}).total;
  if(ins>0){
    items.push({l:'ok',e:'🛡️',
      h:`סה"כ ביטוחים: <strong>₪${fmt(ins)}</strong> לתקופה`});
  }

  document.getElementById('insights-grid').innerHTML=items.map(it=>
    `<div class="ic ${it.l}"><span class="emoji">${it.e}</span><span class="body">${it.h}</span></div>`
  ).join('');
}

// ── Save HTML ─────────────────────────────────────────────────────────────
document.getElementById('btn-save').addEventListener('click',()=>{
  const blob=new Blob(['<!DOCTYPE html>'+document.documentElement.outerHTML],{type:'text/html;charset=utf-8'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob);
  a.download=`bank_${DATA.source.replace(/\.[^.]+$/,'')}.html`;
  a.click(); URL.revokeObjectURL(a.href);
});

// ── Collapsible sections ──────────────────────────────────────────────────
document.querySelectorAll('.section>h2').forEach(h=>{
  h.addEventListener('click',()=>h.parentElement.classList.toggle('collapsed'));
});

['search','dir-filter','cat-filter'].forEach(id=>
  document.getElementById(id).addEventListener('input',renderAll));

initTheme();
renderCards();
renderMonthly();
renderTopTables();
initFilters();
renderAll();
renderInsights();
</script>
</body>
</html>
"""


def generate_bank_html(data: dict) -> str:
    return BANK_HTML_TEMPLATE.replace("__DATA__", json.dumps(data, ensure_ascii=False))
