#!/usr/bin/env python3
"""
payments_ui.py — parse a credit-card Excel file and generate a web UI.

Usage:
    python3 payments_ui.py "april payments.xlsx" [output.html]

Reads the transactions sheet, writes a self-contained HTML file with:
  * summary cards (total / regular / non-regular / high)
  * charts: category breakdown, daily trend, top merchants
  * flagged table of non-regular or high payments
  * possible duplicate charges
  * recurring subscriptions / standing orders
  * open installment plans with remaining balance
  * the full transactions table with search, filters, sorting
  * dark mode toggle
"""
from __future__ import annotations

import datetime
import json
import re
import sys
import webbrowser
from collections import defaultdict
from pathlib import Path

import openpyxl

HIGH_THRESHOLD = 500  # ILS — charges at/above this are flagged as "high"
REGULAR_TYPE = "רגילה"

# ---------------------------------------------------------------------------
# Merchant normalization
# ---------------------------------------------------------------------------

# Phrases removed verbatim from merchant names before canonicalization.
_SUFFIX_PHRASES = [
    "- עסקאות בהרשאה", "עסקאות בהרשאה",
    "- הו\"ק", "- הו״ק", "הו\"ק", "הו״ק", "הוראת קבע",
    "בע\"מ", "בע״מ",
    "- אונליין- צמרת", "-אונליין",
    "חשבונית חודשית",
    "און ליין", "אונליין",
    "-גמא", "-מזרחי",
]

# Location tokens that get stripped to collapse store branches together.
_LOCATION_WORDS = [
    "אם המושבות", "פתח תקווה", "פ\"ת", "פ״ת", "גיסין", "בשוק",
]

# Explicit merchant aliases — highest priority, applied before regex work.
_CANONICAL_MAP = {
    "BIT": "BIT",
    "PAYBOX": "PAYBOX",
    "WOLT": "Wolt",
    "aliexpress": "AliExpress",
    "APPLE.COM/BILL": "Apple",
    "SpotifyIL": "Spotify",
    "ביי-מי שוברי מתנה": "ביי-מי שוברי מתנה",
}

_BRANCH_RE = re.compile(r"[-\s]+\d+\s*$")
_WHITESPACE_RE = re.compile(r"\s+")
_INSTALLMENT_RE = re.compile(r"תשלום\s*(\d+)\s*(?:מתוך|/)\s*(\d+)")


def _normalize_merchant(name: str) -> str:
    """Return a canonical version of a merchant name, stripping branches and locations."""
    if not name:
        return name
    s = name.strip()

    # Exact alias lookup first.
    if s in _CANONICAL_MAP:
        return _CANONICAL_MAP[s]

    for phrase in _SUFFIX_PHRASES:
        s = s.replace(phrase, " ")
    for loc in _LOCATION_WORDS:
        s = s.replace(loc, " ")

    s = _BRANCH_RE.sub("", s)
    s = _WHITESPACE_RE.sub(" ", s).strip(" -—·")
    return s or name.strip()


# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------


def parse_payments(xlsx_path: Path) -> dict:
    """Parse a payments file (Excel .xlsx or Cal PDF), auto-detecting format.

    Supported layouts:
      * Cal / בינלאומי הראשון — header row starts with "תאריך עסקה"  (.xlsx)
      * Isracard (ישראכרט)    — header row starts with "תאריך רכישה" (.xlsx)
      * Cal digital PDF       — "דף פירוט דיגיטלי" PDF statement     (.pdf)
    """
    if xlsx_path.suffix.lower() == ".pdf":
        return _parse_cal_pdf(xlsx_path)

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    fmt, header_idx = _detect_format(rows)
    if fmt is None:
        raise ValueError(
            "Unrecognized payments file format "
            "(expected Cal / בינלאומי הראשון or Isracard)"
        )

    title = _extract_title(rows, xlsx_path)
    data_rows = rows[header_idx + 1:]

    if fmt == "cal":
        payments = _parse_cal_rows(data_rows)
    else:  # isracard
        payments = _parse_isracard_rows(data_rows)

    for p in payments:
        p["canonical"] = _normalize_merchant(p["merchant"])

    return {
        "title": title,
        "source": xlsx_path.name,
        "issuer": fmt,
        "payments": payments,
    }


# ---------------------------------------------------------------------------
# Cal PDF parser
# ---------------------------------------------------------------------------

_PDF_DATE_EMBEDDED = re.compile(r'^(.*?)(\d{2}/\d{2}/\d{4})$')
_PDF_DATE_ONLY     = re.compile(r'^\d{2}/\d{2}/\d{4}$')
_PDF_AMOUNT        = re.compile(r'^-?[\d,]+\.\d{2}$')
_PDF_HEB           = re.compile(r'[\u05d0-\u05ea]')

# In RTL-sorted word array: for phrase "word1 word2" (word1 on right, word2 on left)
# → lookup key is rev(word1) + ' ' + rev(word2)
_PDF_CATS = [
    ('ירצומ למשח', 'מוצרי חשמל'),
    ('יאנפ יוליב', 'פנאי בילוי'),
    ('חוטיב ניפו', 'ביטוח ופינ'),
    ('בכר רובחתו', 'רכב ותחבור'),
    ('ןוזמ אקשמו', 'מזון ומשקא'),
    ('ןוזמ ריהמ',  'מזון מהיר'),
    ('טוהיר תיבו', 'ריהוט ובית'),
    ('האופר ירבו', 'רפואה וברי'),
    ('תרושקת חמו', 'תקשורת ומח'),
    ('תרושקת',    'תקשורת ומח'),
    ('יתב ובלכ',  'בתי כלבו'),
    ('ובלכ יתב',  'בתי כלבו'),
    ('תודעסמ',    'מסעדות'),
    ('תודסומ',    'מוסדות'),
    ('תונוש',     'שונות'),
    ('היגרנא',    'אנרגיה'),
    ('הנפוא',     'אופנה'),
    ('תוריית',    'תיירות'),
    ('סינניפ',    'פיננסים'),
    ('ירצומ',     'מוצרי חשמל'),
    ('זג',        'גז'),
]

_PDF_STRIP = {
    'ההזמ', 'סיטרכ', 'Pay', 'Apple', 'אל', 'לא', 'תארוה', 'עבק', 'חמו',
    'Å', '<', '>', 'ב', 'ן', 'מ',
}
_PDF_AD_FRAGS = {
    'תגצהב', 'הנתומ', 'הצעה', 'הלוואה', 'תיבירב', 'הכומנ', 'רתוי',
    'קנבהמ', 'הנתומ', 'םיטרפ', 'היצקילפאבו', 'רתאב', 'ףקתו', 'ןיקת',
}


def _pdf_rev(word: str) -> str:
    """Reverse a reversed-Hebrew word back to readable form; keep non-Hebrew as-is."""
    return word[::-1] if _PDF_HEB.search(word) else word


def _parse_cal_pdf(pdf_path: Path) -> dict:
    """Parse a Cal credit-card digital PDF statement."""
    try:
        import pdfplumber
    except ImportError:
        raise ImportError("pdfplumber is required for PDF parsing: pip install pdfplumber")

    all_words: list[dict] = []
    with pdfplumber.open(pdf_path) as pdf:
        for pg, page in enumerate(pdf.pages):
            for w in page.extract_words(x_tolerance=4, y_tolerance=3):
                all_words.append({**w, "abs_top": pg * 10000 + w["top"]})

    # Build row groups anchored on date words
    anchor_ys: dict[int, list] = {}
    for w in all_words:
        m = _PDF_DATE_EMBEDDED.match(w["text"])
        if m or _PDF_DATE_ONLY.match(w["text"]):
            anchor_ys.setdefault(round(w["abs_top"]), []).append(w)

    payments: list[dict] = []
    for anchor_y in sorted(anchor_ys):
        row_words = [w for w in all_words if abs(w["abs_top"] - anchor_y) <= 5]
        row_words.sort(key=lambda w: -w["x0"])   # descending x → RTL order
        texts = [w["text"] for w in row_words]

        # Split fused merchant+date word
        date_str = None
        for i, t in enumerate(texts):
            m = _PDF_DATE_EMBEDDED.match(t)
            if m:
                frag, date_str = m.group(1).strip(), m.group(2)
                texts = texts[:i] + ([frag] if frag else []) + texts[i + 1:]
                break
            if _PDF_DATE_ONLY.match(t):
                date_str = t
                texts = texts[:i] + texts[i + 1:]
                break
        if not date_str:
            continue
        d, mo, y = date_str.split("/")
        if y not in ("2023", "2024", "2025", "2026"):
            continue
        date_iso = f"{y}-{mo}-{d}"

        # Amounts: in RTL array the NUMBER is to the LEFT of ₪ (= lower x = higher index)
        # but appears at index si-1 (lower index in descending sort = higher x = more right)
        # Actually: ₪ is at x, number is at x+dx (right of ₪ in PDF = higher x = lower index)
        shek_idxs = [i for i, t in enumerate(texts) if t == "₪"]
        amounts, amount_idxs = [], set()
        for si in shek_idxs:
            if si > 0 and _PDF_AMOUNT.match(texts[si - 1]):
                amounts.append(float(texts[si - 1].replace(",", "")))
                amount_idxs.update([si - 1, si])
        if not amounts:
            continue

        charge = amounts[-1]   # leftmost in PDF = last found = charge column
        amount = amounts[0] if len(amounts) > 1 else charge

        # Middle tokens = everything before the first amount/₪
        first_amt = min(amount_idxs)
        middle = [t for i, t in enumerate(texts) if i < first_amt]
        mid_str = " ".join(middle)

        # Skip cashback/discount lines
        if "החנה עצבמ" in mid_str or "CalExtra" in mid_str:
            continue

        # Detect type (reversed Hebrew patterns)
        if charge < 0:
            ttype = "זיכוי"
        elif "תארוה עבק" in mid_str:
            ttype = "הוראת קבע"
        elif "םולשת" in mid_str or "םימולשת" in mid_str:
            ttype = "תשלומים"
        else:
            ttype = REGULAR_TYPE

        # Detect category
        category = ""
        for key, val in _PDF_CATS:
            if key in mid_str:
                category = val
                mid_str = mid_str.replace(key, " ").strip()
                break

        # Clean merchant
        mer_tokens = []
        for t in mid_str.split():
            if t in _PDF_STRIP or t in _PDF_AD_FRAGS:
                continue
            if re.match(r"^\d{3,4}$", t):   # card last 4 digits
                continue
            if re.match(r"^\d+$", t):        # installment numbers
                continue
            mer_tokens.append(_pdf_rev(t))

        merchant = " ".join(mer_tokens).strip(" -—·.,")
        merchant = re.sub(r"\s*-?\d[\d,]*\.\d{2}\s*", " ", merchant).strip()
        merchant = re.sub(r"\b(תשלום|מ)\b", "", merchant).strip()
        merchant = re.sub(r"\s+", " ", merchant).strip()
        if not merchant:
            merchant = "—"

        payments.append({
            "date": date_iso,
            "merchant": merchant,
            "amount": abs(amount),
            "charge": charge,
            "type": ttype,
            "category": category,
            "notes": "",
        })

    # Normalize canonical names
    for p in payments:
        p["canonical"] = _normalize_merchant(p["merchant"])

    # Build a title from the PDF filename
    title = pdf_path.stem.replace("_", " ")

    return {
        "title": title,
        "source": pdf_path.name,
        "issuer": "cal_pdf",
        "payments": payments,
    }


def _detect_format(rows):
    """Return ('cal'|'isracard', header_row_index) or (None, -1)."""
    for i, row in enumerate(rows):
        if not row or not isinstance(row[0], str):
            continue
        h = row[0]
        if "תאריך\nעסקה" in h or h.strip() == "תאריך עסקה":
            return "cal", i
        if "תאריך רכישה" in h:
            return "isracard", i
    return None, -1


def _extract_title(rows, xlsx_path: Path) -> str:
    """Use the first non-empty row as the document title."""
    for row in rows:
        if not row:
            continue
        parts = [c.strip() for c in row if isinstance(c, str) and c.strip()]
        if parts:
            return " · ".join(parts) if len(parts) > 1 else parts[0]
    return xlsx_path.stem


def _fmt_date(value) -> str:
    if isinstance(value, datetime.datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, str):
        for pattern in ("%d.%m.%y", "%d/%m/%Y", "%d/%m/%y"):
            try:
                return datetime.datetime.strptime(value.strip(), pattern).strftime("%Y-%m-%d")
            except ValueError:
                continue
        return value.strip()
    return ""


def _parse_cal_rows(rows) -> list[dict]:
    payments = []
    for row in rows:
        row = list(row) + [None] * max(0, 7 - len(row))
        date, merchant, amount, charge, ttype, category, notes = row[:7]
        if not isinstance(merchant, str) or not isinstance(charge, (int, float)):
            continue
        payments.append({
            "date": _fmt_date(date),
            "merchant": merchant.strip(),
            "amount": float(amount) if isinstance(amount, (int, float)) else 0.0,
            "charge": float(charge),
            "type": (ttype or "").strip(),
            "category": (category or "").strip(),
            "notes": (notes or "").strip() if isinstance(notes, str) else "",
        })
    return payments


def _parse_isracard_rows(rows) -> list[dict]:
    """Isracard columns: date, merchant, amount, amt_curr, charge, chg_curr, voucher, extras."""
    payments = []
    for row in rows:
        row = list(row) + [None] * max(0, 8 - len(row))
        date, merchant, amount, amt_curr, charge, chg_curr, _voucher, extras = row[:8]
        if date is None or not isinstance(merchant, str) or not isinstance(charge, (int, float)):
            continue

        extras_text = extras.strip() if isinstance(extras, str) else ""

        if "תשלום" in extras_text:
            ttype = "תשלומים"
        elif "הוראת קבע" in extras_text:
            ttype = "הוראת קבע"
        elif "זיכוי" in extras_text or charge < 0:
            ttype = "זיכוי"
        else:
            ttype = REGULAR_TYPE

        category = ""
        if isinstance(amt_curr, str) and amt_curr.strip() and amt_curr.strip() != "₪":
            category = f"חו\"ל ({amt_curr.strip()})"

        payments.append({
            "date": _fmt_date(date),
            "merchant": merchant.strip(),
            "amount": float(amount) if isinstance(amount, (int, float)) else 0.0,
            "charge": float(charge),
            "type": ttype,
            "category": category,
            "notes": extras_text,
        })
    return payments


# ---------------------------------------------------------------------------
# Summary & insights
# ---------------------------------------------------------------------------


def build_summary(payments: list[dict]) -> dict:
    regular = [p for p in payments if p["type"] == REGULAR_TYPE]
    non_regular = [p for p in payments if p["type"] and p["type"] != REGULAR_TYPE]
    high = [p for p in payments if p["charge"] >= HIGH_THRESHOLD]
    return {
        "total_count": len(payments),
        "total_amount": sum(p["charge"] for p in payments),
        "regular_count": len(regular),
        "regular_amount": sum(p["charge"] for p in regular),
        "non_regular_count": len(non_regular),
        "non_regular_amount": sum(p["charge"] for p in non_regular),
        "high_count": len(high),
        "high_amount": sum(p["charge"] for p in high),
    }


def _parse_installment(notes: str):
    m = _INSTALLMENT_RE.search(notes or "")
    if m:
        return int(m.group(1)), int(m.group(2))
    return None


def build_insights(payments: list[dict]) -> dict:
    """Compute analytics: categories, top merchants, daily trend, duplicates, subs, installments."""
    # Category breakdown
    cats = defaultdict(lambda: {"count": 0, "total": 0.0})
    for p in payments:
        c = p["category"] or "ללא קטגוריה"
        cats[c]["count"] += 1
        cats[c]["total"] += p["charge"]
    categories = sorted(
        [{"name": k, **v} for k, v in cats.items()],
        key=lambda x: x["total"],
        reverse=True,
    )

    # Top merchants (by canonical name)
    merchants = defaultdict(lambda: {"count": 0, "total": 0.0, "aliases": set()})
    for p in payments:
        m = p.get("canonical") or p["merchant"]
        merchants[m]["count"] += 1
        merchants[m]["total"] += p["charge"]
        merchants[m]["aliases"].add(p["merchant"])
    top_merchants = sorted(
        [
            {
                "name": k,
                "count": v["count"],
                "total": v["total"],
                "aliases": sorted(v["aliases"]),
            }
            for k, v in merchants.items()
        ],
        key=lambda x: x["total"],
        reverse=True,
    )[:20]

    # Daily spending trend
    days = defaultdict(float)
    for p in payments:
        if p["date"]:
            days[p["date"]] += p["charge"]
    daily_trend = sorted(
        [{"date": d, "total": round(t, 2)} for d, t in days.items()],
        key=lambda x: x["date"],
    )

    # Duplicate detection: same canonical merchant + same amount + same date
    dup_groups = defaultdict(list)
    for p in payments:
        key = (p.get("canonical") or p["merchant"], round(p["charge"], 2), p["date"])
        dup_groups[key].append(p)
    duplicates = []
    for (merchant, amount, date), items in dup_groups.items():
        if len(items) >= 2 and amount > 0:
            duplicates.append({
                "merchant": merchant,
                "amount": amount,
                "date": date,
                "count": len(items),
                "total": round(amount * len(items), 2),
            })
    duplicates.sort(key=lambda x: (-x["total"], x["date"]))

    # Subscriptions / recurring charges (standing orders)
    sub_groups = defaultdict(lambda: {"count": 0, "total": 0.0, "amounts": set(), "dates": []})
    for p in payments:
        if p["type"] == "הוראת קבע":
            key = p.get("canonical") or p["merchant"]
            sub_groups[key]["count"] += 1
            sub_groups[key]["total"] += p["charge"]
            sub_groups[key]["amounts"].add(round(p["charge"], 2))
            sub_groups[key]["dates"].append(p["date"])
    subscriptions = sorted(
        [
            {
                "merchant": k,
                "count": v["count"],
                "total": round(v["total"], 2),
                "amounts": sorted(v["amounts"]),
                "dates": sorted(v["dates"]),
            }
            for k, v in sub_groups.items()
        ],
        key=lambda x: x["total"],
        reverse=True,
    )

    # Open installment plans with remaining balance
    installments = []
    for p in payments:
        is_installment = p["type"] == "תשלומים" or "תשלום" in (p["notes"] or "")
        if not is_installment:
            continue
        parsed = _parse_installment(p["notes"] or "")
        if parsed:
            current, total_count = parsed
            remaining_count = max(total_count - current, 0)
            remaining_amount = round(remaining_count * p["charge"], 2)
        else:
            current, total_count, remaining_count, remaining_amount = 0, 0, 0, 0.0
        installments.append({
            "date": p["date"],
            "merchant": p.get("canonical") or p["merchant"],
            "charge": round(p["charge"], 2),
            "current": current,
            "total": total_count,
            "remaining_count": remaining_count,
            "remaining_amount": remaining_amount,
            "notes": p["notes"] or "",
        })
    installments.sort(key=lambda x: -x["remaining_amount"])
    total_remaining = round(sum(i["remaining_amount"] for i in installments), 2)

    return {
        "categories": categories,
        "top_merchants": top_merchants,
        "daily_trend": daily_trend,
        "duplicates": duplicates,
        "subscriptions": subscriptions,
        "installments": installments,
        "total_installment_remaining": total_remaining,
    }


# ---------------------------------------------------------------------------
# HTML rendering
# ---------------------------------------------------------------------------

HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="he" dir="rtl" data-theme="light">
<head>
<meta charset="UTF-8">
<title>__TITLE__</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
  :root {
    --bg: #f5f5f7; --card: #ffffff; --text: #222; --muted: #666; --soft: #888;
    --border: #eee; --border-strong: #ddd; --hover: #f8f8fb; --th-bg: #fafafa;
    --th-hover: #eef; --shadow: 0 1px 3px rgba(0,0,0,0.08);
    --primary: #2196f3; --high: #c62828; --refund: #2e7d32;
    --chip-bg: #fafafa;
    --b-reg-bg: #e3f2fd; --b-reg-fg: #1565c0;
    --b-stand-bg: #fff3e0; --b-stand-fg: #e65100;
    --b-inst-bg: #f3e5f5; --b-inst-fg: #6a1b9a;
    --b-ref-bg: #e8f5e9; --b-ref-fg: #2e7d32;
    --b-rep-bg: #fff8e1; --b-rep-fg: #8d6e63;
    --b-other-bg: #eee; --b-other-fg: #555;
  }
  [data-theme="dark"] {
    --bg: #111418; --card: #1c2128; --text: #e6edf3; --muted: #9aa4af; --soft: #7a8591;
    --border: #2a323c; --border-strong: #394350; --hover: #232b36; --th-bg: #1a2028;
    --th-hover: #222d3d; --shadow: 0 1px 3px rgba(0,0,0,0.4);
    --primary: #64b5f6; --high: #ef5350; --refund: #66bb6a;
    --chip-bg: #1a2028;
    --b-reg-bg: #0d3b66; --b-reg-fg: #9fd3ff;
    --b-stand-bg: #4a2900; --b-stand-fg: #ffcc80;
    --b-inst-bg: #3a1c4a; --b-inst-fg: #ce93d8;
    --b-ref-bg: #1b4d20; --b-ref-fg: #a5d6a7;
    --b-rep-bg: #3f2f00; --b-rep-fg: #d7ccc8;
    --b-other-bg: #2a323c; --b-other-fg: #aaa;
  }
  * { box-sizing: border-box; }
  body { font-family: -apple-system, "Segoe UI", Arial, sans-serif; margin: 0; padding: 24px;
         background: var(--bg); color: var(--text); transition: background 0.2s, color 0.2s; }
  header { display: flex; justify-content: space-between; align-items: flex-start; gap: 16px; margin-bottom: 16px; }
  header h1 { font-size: 20px; margin: 0 0 4px; }
  .src { color: var(--soft); font-size: 12px; }
  .theme-toggle { background: var(--card); color: var(--text); border: 1px solid var(--border-strong);
                  width: 40px; height: 40px; border-radius: 10px; cursor: pointer; font-size: 18px;
                  box-shadow: var(--shadow); }
  .theme-toggle:hover { background: var(--hover); }
  .cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 12px; margin-bottom: 20px; }
  .card { background: var(--card); padding: 16px; border-radius: 10px; box-shadow: var(--shadow); }
  .card .label { font-size: 11px; color: var(--muted); letter-spacing: 0.5px; }
  .card .value { font-size: 22px; font-weight: 600; margin-top: 4px; }
  .card .sub { font-size: 12px; color: var(--soft); margin-top: 2px; }
  .charts-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap: 16px; margin-bottom: 20px; }
  .chart-card { background: var(--card); padding: 16px 20px; border-radius: 10px; box-shadow: var(--shadow); min-height: 320px; }
  .chart-card h2 { margin: 0 0 12px; font-size: 15px; }
  .chart-wrap { position: relative; height: 260px; }
  .section { background: var(--card); padding: 16px 20px; border-radius: 10px; box-shadow: var(--shadow); margin-bottom: 20px; }
  .section > h2 { margin: 0 0 12px; font-size: 16px; cursor: pointer; user-select: none; }
  .section.collapsed > :not(h2) { display: none; }
  .section > h2::before { content: "▾"; display: inline-block; width: 1em; font-size: 11px; color: var(--muted); transform: scaleX(-1); }
  .section.collapsed > h2::before { content: "▸"; transform: scaleX(-1); }
  table { width: 100%; border-collapse: collapse; font-size: 14px; }
  th, td { padding: 8px 10px; border-bottom: 1px solid var(--border); text-align: right; vertical-align: top; }
  th { background: var(--th-bg); cursor: pointer; user-select: none; font-weight: 600; font-size: 13px; color: var(--text); }
  th:hover { background: var(--th-hover); }
  tr:hover td { background: var(--hover); }
  .filter { display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 12px; }
  .filter input, .filter select { padding: 8px 12px; border: 1px solid var(--border-strong); border-radius: 6px;
                                   font-size: 14px; font-family: inherit; background: var(--card); color: var(--text); }
  .filter input { min-width: 220px; }
  .badge { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: 600; white-space: nowrap; }
  .badge-regular     { background: var(--b-reg-bg);   color: var(--b-reg-fg); }
  .badge-standing    { background: var(--b-stand-bg); color: var(--b-stand-fg); }
  .badge-installment { background: var(--b-inst-bg);  color: var(--b-inst-fg); }
  .badge-refund      { background: var(--b-ref-bg);   color: var(--b-ref-fg); }
  .badge-repay       { background: var(--b-rep-bg);   color: var(--b-rep-fg); }
  .badge-other       { background: var(--b-other-bg); color: var(--b-other-fg); }
  .amount-high { color: var(--high); font-weight: 600; }
  .amount-refund { color: var(--refund); }
  .num { font-variant-numeric: tabular-nums; white-space: nowrap; }
  .reason { color: var(--muted); font-size: 12px; }
  .count { color: var(--soft); font-size: 12px; font-weight: normal; margin-right: 6px; }
  .aliases { color: var(--soft); font-size: 11px; }
  .empty { color: var(--soft); font-style: italic; padding: 8px 0; }
  .sum-row td { border-top: 2px solid var(--border-strong); border-bottom: none; font-weight: 700; background: var(--th-bg); }
  .chk-col { width: 32px; text-align: center !important; padding: 6px 4px !important; }
  .chk-col input { width: 16px; height: 16px; cursor: pointer; accent-color: var(--primary); }
  .insights-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 10px; }
  .insight-card { display: flex; align-items: flex-start; gap: 10px; padding: 10px 14px;
    border-radius: 8px; border-right: 4px solid transparent; background: var(--bg); font-size: 14px; line-height: 1.5; }
  .insight-card.ok   { border-color: #43a047; }
  .insight-card.warn { border-color: #fb8c00; }
  .insight-card.alert{ border-color: var(--high); }
  .insight-card .ic  { font-size: 20px; line-height: 1; flex-shrink: 0; margin-top: 1px; }
  .insight-card .body { flex: 1; color: var(--text); }
  .insight-card .body strong { color: var(--primary); font-weight: 700; }
  .floating-bar { position: fixed; bottom: -400px; left: 50%; transform: translateX(-50%);
    background: var(--card); border: 1px solid var(--border-strong); box-shadow: 0 -4px 24px rgba(0,0,0,0.18);
    border-radius: 14px; padding: 12px 20px;
    display: flex; flex-direction: column; gap: 0;
    font-size: 15px; font-weight: 600; z-index: 100; transition: bottom 0.3s ease;
    min-width: 340px; max-width: 540px; width: 90%; }
  .floating-bar.visible { bottom: 24px; }
  .floating-bar .bar-top { display: flex; align-items: center; gap: 16px; white-space: nowrap; flex-wrap: wrap; }
  .floating-bar .sel-total { color: var(--primary); }
  .floating-bar button { background: none; border: 1px solid var(--border-strong); color: var(--muted);
    padding: 4px 14px; border-radius: 6px; font-size: 13px; cursor: pointer; font-family: inherit; }
  .floating-bar button:hover { background: var(--hover); color: var(--text); }
  .floating-bar .sel-items { margin-top: 8px; border-top: 1px solid var(--border); padding-top: 6px;
    max-height: 180px; overflow-y: auto; display: flex; flex-direction: column; gap: 1px; }
  .floating-bar .sel-item { display: flex; justify-content: space-between; align-items: baseline;
    padding: 3px 2px; font-size: 13px; font-weight: 400; border-bottom: 1px solid var(--border); }
  .floating-bar .sel-item:last-child { border-bottom: none; }
  .floating-bar .sel-item .item-label { color: var(--muted); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 320px; }
  .floating-bar .sel-item .item-amount { font-weight: 600; white-space: nowrap; padding-right: 12px; color: var(--text); }
</style>
</head>
<body>

<header>
  <div>
    <h1 id="title"></h1>
    <div class="src" id="src"></div>
  </div>
  <div style="display:flex;gap:8px;align-items:flex-start;">
    <a href="/" class="theme-toggle" title="העלאת קובץ חדש" style="width:auto;padding:0 14px;font-size:13px;font-weight:600;text-decoration:none;display:flex;align-items:center;height:40px;">← קובץ חדש</a>
    <button class="theme-toggle" id="btn-save-html" title="שמור כ-HTML לצפייה ללא שרת" style="width:auto;padding:0 14px;font-size:13px;font-weight:600;">💾 שמור HTML</button>
    <button class="theme-toggle" id="btn-save" title="שמור כ-JSON לטעינה מחדש ללא Excel" style="width:auto;padding:0 14px;font-size:13px;font-weight:600;">💾 שמור JSON</button>
    <button class="theme-toggle" id="theme-toggle" title="החלף מצב כהה/בהיר">🌙</button>
  </div>
</header>

<div class="cards" id="cards"></div>

<div class="charts-grid">
  <div class="chart-card"><h2>חלוקה לפי ענף</h2><div class="chart-wrap"><canvas id="chart-category"></canvas></div></div>
  <div class="chart-card"><h2>מגמה יומית</h2><div class="chart-wrap"><canvas id="chart-daily"></canvas></div></div>
  <div class="chart-card"><h2>Top בתי עסק</h2><div class="chart-wrap"><canvas id="chart-merchants"></canvas></div></div>
</div>

<div class="section" id="sec-insights">
  <h2>תובנות חכמות 🧠</h2>
  <div class="insights-grid" id="insights-grid"></div>
</div>

<div class="section" id="sec-flagged">
  <h2>סיכום חריגים — לא רגילות או סכום גבוה <span class="count" id="flagged-count"></span></h2>
  <table id="flagged-table">
    <thead><tr>
      <th class="chk-col"><input type="checkbox" class="select-all"></th><th>תאריך</th><th>בית עסק</th><th>סוג</th><th>ענף</th><th>סכום (₪)</th><th>סיבה</th>
    </tr></thead>
    <tbody></tbody>
  </table>
</div>

<div class="section" id="sec-duplicates">
  <h2>תשלומים כפולים אפשריים <span class="count" id="duplicates-count"></span></h2>
  <table id="duplicates-table">
    <thead><tr>
      <th class="chk-col"><input type="checkbox" class="select-all"></th><th>תאריך</th><th>בית עסק</th><th>סכום בודד (₪)</th><th>חזרות</th><th>סה"כ (₪)</th>
    </tr></thead>
    <tbody></tbody>
  </table>
</div>

<div class="section" id="sec-subscriptions">
  <h2>מנויים והוראות קבע <span class="count" id="subscriptions-count"></span></h2>
  <table id="subscriptions-table">
    <thead><tr>
      <th class="chk-col"><input type="checkbox" class="select-all"></th><th>בית עסק</th><th>חיובים בחודש</th><th>סכומים (₪)</th><th>סה"כ (₪)</th>
    </tr></thead>
    <tbody></tbody>
  </table>
</div>

<div class="section" id="sec-installments">
  <h2>תשלומים פתוחים <span class="count" id="installments-count"></span></h2>
  <div class="filter"><div class="src" id="installments-remaining"></div></div>
  <table id="installments-table">
    <thead><tr>
      <th class="chk-col"><input type="checkbox" class="select-all"></th><th>תאריך</th><th>בית עסק</th><th>תשלום חודשי (₪)</th><th>תשלום נוכחי</th><th>נותרו</th><th>יתרה (₪)</th>
    </tr></thead>
    <tbody></tbody>
  </table>
</div>

<div class="section" id="sec-merchants">
  <h2>Top בתי עסק (לפי סכום) <span class="count" id="merchants-count"></span></h2>
  <table id="merchants-table">
    <thead><tr>
      <th class="chk-col"><input type="checkbox" class="select-all"></th><th>בית עסק</th><th>עסקאות</th><th>סה"כ (₪)</th><th>ממוצע (₪)</th>
    </tr></thead>
    <tbody></tbody>
  </table>
</div>

<div class="section" id="sec-all">
  <h2>כל העסקאות <span class="count" id="all-count"></span></h2>
  <div class="filter">
    <input id="search" type="text" placeholder="חיפוש בית עסק / ענף...">
    <select id="type-filter"><option value="">כל הסוגים</option></select>
    <select id="category-filter"><option value="">כל הענפים</option></select>
  </div>
  <table id="all-table">
    <thead><tr>
      <th class="chk-col"><input type="checkbox" class="select-all"></th>
      <th data-k="date">תאריך</th>
      <th data-k="merchant">בית עסק</th>
      <th data-k="type">סוג</th>
      <th data-k="category">ענף</th>
      <th data-k="charge">סכום (₪)</th>
      <th data-k="notes">הערות</th>
    </tr></thead>
    <tbody></tbody>
  </table>
</div>

<div class="floating-bar" id="floating-bar">
  <div class="bar-top">
    <span id="sel-summary"></span>
    <span class="sel-total" id="sel-total"></span>
    <button id="sel-clear">נקה בחירה</button>
  </div>
  <div class="sel-items" id="sel-items"></div>
</div>

<script>
const DATA = __DATA__;

const fmt = n => new Intl.NumberFormat('he-IL', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(n);
const fmt0 = n => new Intl.NumberFormat('he-IL', { maximumFractionDigits: 0 }).format(n);

function esc(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[c]));
}

function typeBadge(t) {
  const map = { 'רגילה':'regular','הוראת קבע':'standing','תשלומים':'installment','זיכוי':'refund','פרעון':'repay' };
  const cls = map[t] || 'other';
  return `<span class="badge badge-${cls}">${esc(t || '—')}</span>`;
}

function amountClass(p) {
  if (p.charge < 0) return 'amount-refund';
  if (p.charge >= DATA.high_threshold) return 'amount-high';
  return '';
}

// -------------------- Dark mode --------------------

function applyTheme(t) {
  document.documentElement.dataset.theme = t;
  document.getElementById('theme-toggle').textContent = t === 'dark' ? '☀️' : '🌙';
  renderCharts(); // rebuild with new theme colors
}
function initTheme() {
  const saved = localStorage.getItem('payments-theme') ||
    (window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
  applyTheme(saved);
  document.getElementById('theme-toggle').addEventListener('click', () => {
    const cur = document.documentElement.dataset.theme;
    const next = cur === 'dark' ? 'light' : 'dark';
    localStorage.setItem('payments-theme', next);
    applyTheme(next);
  });
}

// -------------------- Cards --------------------

function renderCards() {
  const s = DATA.summary;
  document.getElementById('title').textContent = DATA.title;
  document.getElementById('src').textContent = 'מקור: ' + DATA.source;
  const cards = [
    ['סה"כ עסקאות', s.total_count, 'חיוב כולל: ₪' + fmt(s.total_amount)],
    ['רגילות', s.regular_count, '₪' + fmt(s.regular_amount)],
    ['לא רגילות', s.non_regular_count, '₪' + fmt(s.non_regular_amount)],
    ['סכום גבוה (≥ ₪' + DATA.high_threshold + ')', s.high_count, '₪' + fmt(s.high_amount)],
  ];
  document.getElementById('cards').innerHTML = cards
    .map(([l, v, sub]) => `<div class="card"><div class="label">${esc(l)}</div><div class="value">${esc(v)}</div><div class="sub">${esc(sub)}</div></div>`)
    .join('');
}

// -------------------- Charts --------------------

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

  const ins = DATA.insights;

  // Category doughnut
  const cats = ins.categories.slice(0, 10);
  charts.cat = new Chart(document.getElementById('chart-category'), {
    type: 'doughnut',
    data: {
      labels: cats.map(c => c.name),
      datasets: [{ data: cats.map(c => c.total), backgroundColor: chartColors(cats.length), borderWidth: 0 }],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: 'bottom', labels: { boxWidth: 12, font: { size: 11 } } } },
    },
  });

  // Daily trend line
  const dt = ins.daily_trend;
  charts.daily = new Chart(document.getElementById('chart-daily'), {
    type: 'line',
    data: {
      labels: dt.map(d => d.date.slice(5)),
      datasets: [{ data: dt.map(d => d.total), borderColor: '#2196f3', backgroundColor: 'rgba(33,150,243,0.15)', tension: 0.3, fill: true, pointRadius: 2 }],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { ticks: { font: { size: 10 } } },
        y: { ticks: { callback: v => '₪' + fmt0(v), font: { size: 10 } } },
      },
    },
  });

  // Top merchants horizontal bar
  const tm = ins.top_merchants.slice(0, 10);
  charts.merchants = new Chart(document.getElementById('chart-merchants'), {
    type: 'bar',
    data: {
      labels: tm.map(m => m.name),
      datasets: [{ data: tm.map(m => m.total), backgroundColor: '#43a047' }],
    },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { ticks: { callback: v => '₪' + fmt0(v), font: { size: 10 } } },
        y: { ticks: { font: { size: 11 } } },
      },
    },
  });
}

// -------------------- Tables --------------------

function renderFlagged() {
  const rows = DATA.payments.map(p => {
    const reasons = [];
    if (p.type && p.type !== 'רגילה') reasons.push('סוג: ' + p.type);
    if (p.charge >= DATA.high_threshold) reasons.push('סכום גבוה');
    return { p, reasons };
  }).filter(x => x.reasons.length > 0);
  rows.sort((a, b) => b.p.charge - a.p.charge);

  document.getElementById('flagged-count').textContent = `(${rows.length})`;
  const tbody = document.querySelector('#flagged-table tbody');
  if (!rows.length) { tbody.innerHTML = `<tr><td colspan="7" class="empty">אין עסקאות חריגות</td></tr>`; return; }
  const total = rows.reduce((s, x) => s + x.p.charge, 0);
  tbody.innerHTML = rows.map(({ p, reasons }) => `
    <tr>
      <td class="chk-col"><input type="checkbox" class="row-chk" data-amount="${p.charge}" data-label="${esc(p.date + ' · ' + p.merchant)}"></td>
      <td class="num">${esc(p.date)}</td>
      <td>${esc(p.merchant)}</td>
      <td>${typeBadge(p.type)}</td>
      <td>${esc(p.category || '—')}</td>
      <td class="num ${amountClass(p)}">${fmt(p.charge)}</td>
      <td class="reason">${esc(reasons.join(' · '))}</td>
    </tr>
  `).join('') + `<tr class="sum-row"><td></td><td colspan="4">סה"כ</td><td class="num">${fmt(total)}</td><td></td></tr>`;
}

function renderDuplicates() {
  const rows = DATA.insights.duplicates;
  document.getElementById('duplicates-count').textContent = `(${rows.length})`;
  const tbody = document.querySelector('#duplicates-table tbody');
  if (!rows.length) { tbody.innerHTML = `<tr><td colspan="6" class="empty">לא נמצאו חיובים כפולים חשודים</td></tr>`; return; }
  const total = rows.reduce((s, d) => s + d.total, 0);
  tbody.innerHTML = rows.map(d => `
    <tr>
      <td class="chk-col"><input type="checkbox" class="row-chk" data-amount="${d.total}" data-label="${esc(d.date + ' · ' + d.merchant + ' (×' + d.count + ')')}"></td>
      <td class="num">${esc(d.date)}</td>
      <td>${esc(d.merchant)}</td>
      <td class="num">${fmt(d.amount)}</td>
      <td class="num">${d.count}</td>
      <td class="num amount-high">${fmt(d.total)}</td>
    </tr>
  `).join('') + `<tr class="sum-row"><td></td><td colspan="4">סה"כ</td><td class="num">${fmt(total)}</td></tr>`;
}

function renderSubscriptions() {
  const rows = DATA.insights.subscriptions;
  document.getElementById('subscriptions-count').textContent = `(${rows.length})`;
  const tbody = document.querySelector('#subscriptions-table tbody');
  if (!rows.length) { tbody.innerHTML = `<tr><td colspan="5" class="empty">לא זוהו הוראות קבע</td></tr>`; return; }
  const total = rows.reduce((s, r) => s + r.total, 0);
  tbody.innerHTML = rows.map(s => `
    <tr>
      <td class="chk-col"><input type="checkbox" class="row-chk" data-amount="${s.total}" data-label="${esc('הו&quot;ק · ' + s.merchant)}"></td>
      <td>${esc(s.merchant)}</td>
      <td class="num">${s.count}</td>
      <td class="num">${s.amounts.map(a => fmt(a)).join(', ')}</td>
      <td class="num">${fmt(s.total)}</td>
    </tr>
  `).join('') + `<tr class="sum-row"><td></td><td>סה"כ</td><td class="num">${rows.reduce((s,r) => s+r.count, 0)}</td><td></td><td class="num">${fmt(total)}</td></tr>`;
}

function renderInstallments() {
  const rows = DATA.insights.installments;
  const total = DATA.insights.total_installment_remaining;
  document.getElementById('installments-count').textContent = `(${rows.length})`;
  document.getElementById('installments-remaining').textContent = `יתרה עתידית כוללת: ₪${fmt(total)}`;
  const tbody = document.querySelector('#installments-table tbody');
  if (!rows.length) { tbody.innerHTML = `<tr><td colspan="7" class="empty">אין תשלומים פתוחים</td></tr>`; return; }
  const totalCharge = rows.reduce((s, i) => s + i.charge, 0);
  const totalRemaining = rows.reduce((s, i) => s + i.remaining_amount, 0);
  tbody.innerHTML = rows.map(i => `
    <tr>
      <td class="chk-col"><input type="checkbox" class="row-chk" data-amount="${i.charge}" data-label="${esc(i.date + ' · ' + i.merchant + (i.total ? ' (' + i.current + '/' + i.total + ')' : ''))}"></td>
      <td class="num">${esc(i.date)}</td>
      <td>${esc(i.merchant)}</td>
      <td class="num">${fmt(i.charge)}</td>
      <td class="num">${i.current || '—'}${i.total ? ' / ' + i.total : ''}</td>
      <td class="num">${i.remaining_count || 0}</td>
      <td class="num ${i.remaining_amount > 0 ? 'amount-high' : ''}">${fmt(i.remaining_amount)}</td>
    </tr>
  `).join('') + `<tr class="sum-row"><td></td><td colspan="2">סה"כ</td><td class="num">${fmt(totalCharge)}</td><td></td><td></td><td class="num">${fmt(totalRemaining)}</td></tr>`;
}

function renderMerchants() {
  const rows = DATA.insights.top_merchants;
  document.getElementById('merchants-count').textContent = `(${rows.length})`;
  const tbody = document.querySelector('#merchants-table tbody');
  if (!rows.length) { tbody.innerHTML = `<tr><td colspan="5" class="empty">—</td></tr>`; return; }
  const totalCount = rows.reduce((s, m) => s + m.count, 0);
  const totalAmount = rows.reduce((s, m) => s + m.total, 0);
  tbody.innerHTML = rows.map(m => `
    <tr>
      <td class="chk-col"><input type="checkbox" class="row-chk" data-amount="${m.total}" data-label="${esc(m.name + ' (' + m.count + ' עסקאות)')}"></td>
      <td>${esc(m.name)}${m.aliases.length > 1 ? `<div class="aliases">${esc(m.aliases.join(' · '))}</div>` : ''}</td>
      <td class="num">${m.count}</td>
      <td class="num">${fmt(m.total)}</td>
      <td class="num">${fmt(m.total / m.count)}</td>
    </tr>
  `).join('') + `<tr class="sum-row"><td></td><td>סה"כ</td><td class="num">${totalCount}</td><td class="num">${fmt(totalAmount)}</td><td></td></tr>`;
}

let sortKey = 'date';
let sortDir = -1;

function renderAll() {
  const q = document.getElementById('search').value.trim().toLowerCase();
  const tf = document.getElementById('type-filter').value;
  const cf = document.getElementById('category-filter').value;

  let rows = DATA.payments.filter(p => {
    if (q && !(p.merchant.toLowerCase().includes(q) || (p.category || '').toLowerCase().includes(q))) return false;
    if (tf && p.type !== tf) return false;
    if (cf && (p.category || '') !== cf) return false;
    return true;
  });

  rows.sort((a, b) => {
    const av = a[sortKey], bv = b[sortKey];
    if (av === '' || av == null) return 1;
    if (bv === '' || bv == null) return -1;
    if (typeof av === 'number' && typeof bv === 'number') return (av - bv) * sortDir;
    return String(av).localeCompare(String(bv), 'he') * sortDir;
  });

  const total = rows.reduce((s, p) => s + p.charge, 0);
  document.getElementById('all-count').textContent = `(${rows.length})`;
  document.querySelector('#all-table tbody').innerHTML = rows.map(p => `
    <tr>
      <td class="chk-col"><input type="checkbox" class="row-chk" data-amount="${p.charge}" data-label="${esc(p.date + ' · ' + p.merchant)}"></td>
      <td class="num">${esc(p.date)}</td>
      <td>${esc(p.merchant)}</td>
      <td>${typeBadge(p.type)}</td>
      <td>${esc(p.category || '—')}</td>
      <td class="num ${amountClass(p)}">${fmt(p.charge)}</td>
      <td>${esc(p.notes || '')}</td>
    </tr>
  `).join('') + (rows.length ? `<tr class="sum-row"><td></td><td colspan="4">סה"כ</td><td class="num">${fmt(total)}</td><td></td></tr>` : '');
  const sa = document.querySelector('#all-table .select-all');
  if (sa) sa.checked = false;
  updateFloatingBar();
}

// -------------------- Smart Insights --------------------

function renderInsights() {
  const payments = DATA.payments;
  const ins = DATA.insights;
  const sum = DATA.summary;
  const items = [];

  // ── Category dominance ──────────────────────────────────────────────────
  if (ins.categories.length) {
    const top = ins.categories[0];
    const pct = Math.round(top.total / sum.total_amount * 100);
    items.push({
      ic: '🏆', level: pct > 50 ? 'warn' : 'ok',
      html: `הקטגוריה הגדולה ביותר היא <strong>${esc(top.name)}</strong> — ₪${fmt(top.total)} (${pct}% מסה"כ)`,
    });
    if (ins.categories.length >= 2) {
      const second = ins.categories[1];
      const pct2 = Math.round(second.total / sum.total_amount * 100);
      items.push({
        ic: '🥈', level: 'ok',
        html: `קטגוריה שנייה: <strong>${esc(second.name)}</strong> — ₪${fmt(second.total)} (${pct2}%)`,
      });
    }
  }

  // ── Subscription burden ─────────────────────────────────────────────────
  if (ins.subscriptions.length) {
    const subTotal = ins.subscriptions.reduce((s, r) => s + r.total, 0);
    const subPct = Math.round(subTotal / sum.total_amount * 100);
    items.push({
      ic: '🔄', level: subPct > 25 ? 'warn' : 'ok',
      html: `${ins.subscriptions.length} הוראות קבע בסה"כ <strong>₪${fmt(subTotal)}</strong> — ${subPct}% מסך ההוצאות`,
    });
    // Most expensive subscription
    const topSub = ins.subscriptions[0];
    items.push({
      ic: '💳', level: 'ok',
      html: `הוראת הקבע הגדולה ביותר: <strong>${esc(topSub.merchant)}</strong> — ₪${fmt(topSub.total)}`,
    });
  }

  // ── Future installment debt ─────────────────────────────────────────────
  if (ins.total_installment_remaining > 0) {
    const biggest = [...ins.installments].sort((a, b) => b.remaining_amount - a.remaining_amount)[0];
    items.push({
      ic: '📅', level: ins.total_installment_remaining > 3000 ? 'warn' : 'ok',
      html: `יתרת תשלומים עתידיים: <strong>₪${fmt(ins.total_installment_remaining)}</strong> (${ins.installments.length} תוכניות) | הגדולה: <strong>${esc(biggest.merchant)}</strong>`,
    });
  }

  // ── Duplicate charges alert ─────────────────────────────────────────────
  if (ins.duplicates.length) {
    const dupTotal = ins.duplicates.reduce((s, d) => s + d.total, 0);
    items.push({
      ic: '⚠️', level: 'alert',
      html: `<strong>${ins.duplicates.length} חיובים כפולים חשודים</strong> בסה"כ ₪${fmt(dupTotal)} — מומלץ לבדוק`,
    });
  }

  // ── Top merchant by spend ───────────────────────────────────────────────
  if (ins.top_merchants.length) {
    const top = ins.top_merchants[0];
    const pct = Math.round(top.total / sum.total_amount * 100);
    items.push({
      ic: '🏪', level: 'ok',
      html: `בית העסק עם ההוצאה הגבוהה ביותר: <strong>${esc(top.name)}</strong> — ₪${fmt(top.total)} (${pct}%, ${top.count} עסקאות)`,
    });
  }

  // ── Most frequent merchant ─────────────────────────────────────────────
  const byCount = [...ins.top_merchants].sort((a, b) => b.count - a.count);
  if (byCount.length && byCount[0] !== ins.top_merchants[0] && byCount[0].count >= 3) {
    const top = byCount[0];
    items.push({
      ic: '🔁', level: 'ok',
      html: `הכי הרבה ביקורים: <strong>${esc(top.name)}</strong> — ${top.count} עסקאות (ממוצע ₪${fmt(top.total / top.count)} לביקור)`,
    });
  }

  // ── Average transaction size ────────────────────────────────────────────
  if (sum.total_count > 0) {
    const avg = sum.total_amount / sum.total_count;
    items.push({
      ic: '📊', level: 'ok',
      html: `<strong>${sum.total_count}</strong> עסקאות בסה"כ · ממוצע לעסקה: <strong>₪${fmt(avg)}</strong>`,
    });
  }

  // ── Daily spend + peak day ──────────────────────────────────────────────
  if (ins.daily_trend.length) {
    const avgDaily = sum.total_amount / ins.daily_trend.length;
    const peak = ins.daily_trend.reduce((a, b) => b.total > a.total ? b : a);
    items.push({
      ic: '📈', level: 'ok',
      html: `ממוצע יומי: <strong>₪${fmt(avgDaily)}</strong> · יום שיא: <strong>${peak.date}</strong> (₪${fmt(peak.total)})`,
    });
  }

  // ── Day-of-week pattern ─────────────────────────────────────────────────
  const dowTotals = [0,0,0,0,0,0,0];
  const dowCounts = [0,0,0,0,0,0,0];
  const dowNames  = ['ראשון','שני','שלישי','רביעי','חמישי','שישי','שבת'];
  for (const p of payments) {
    if (!p.date) continue;
    const d = new Date(p.date).getDay();
    dowTotals[d] += p.charge;
    dowCounts[d]++;
  }
  const topDow = dowTotals.indexOf(Math.max(...dowTotals));
  if (dowCounts[topDow] > 0) {
    items.push({
      ic: '🗓️', level: 'ok',
      html: `יום ההוצאה המרוכז ביותר: <strong>יום ${dowNames[topDow]}</strong> — ₪${fmt(dowTotals[topDow])} ב-${dowCounts[topDow]} עסקאות`,
    });
  }

  // ── Largest single transaction ──────────────────────────────────────────
  const positivePayments = payments.filter(p => p.charge > 0);
  if (positivePayments.length) {
    const biggest = positivePayments.reduce((a, b) => b.charge > a.charge ? b : a);
    items.push({
      ic: '💸', level: biggest.charge >= DATA.high_threshold * 3 ? 'warn' : 'ok',
      html: `עסקה גדולה ביותר: <strong>${esc(biggest.merchant)}</strong> — ₪${fmt(biggest.charge)} (${biggest.date})`,
    });
  }

  // ── High-value transaction count ────────────────────────────────────────
  const highCount = payments.filter(p => p.charge >= DATA.high_threshold).length;
  if (highCount > 0) {
    const highTotal = payments.filter(p => p.charge >= DATA.high_threshold).reduce((s, p) => s + p.charge, 0);
    const highPct = Math.round(highTotal / sum.total_amount * 100);
    items.push({
      ic: '🔺', level: highPct > 40 ? 'warn' : 'ok',
      html: `<strong>${highCount}</strong> עסקאות מעל ₪${DATA.high_threshold} — סה"כ ₪${fmt(highTotal)} (${highPct}% מהסכום הכולל)`,
    });
  }

  // ── Foreign currency ───────────────────────────────────────────────────
  const foreign = payments.filter(p => p.category && p.category.includes('חו"ל'));
  if (foreign.length) {
    const foreignTotal = foreign.reduce((s, p) => s + p.charge, 0);
    const foreignPct = Math.round(foreignTotal / sum.total_amount * 100);
    items.push({
      ic: '🌍', level: foreignPct > 20 ? 'warn' : 'ok',
      html: `<strong>${foreign.length}</strong> עסקאות בחו"ל — ₪${fmt(foreignTotal)} (${foreignPct}% מסה"כ)`,
    });
  }

  // ── Refunds ────────────────────────────────────────────────────────────
  const refunds = payments.filter(p => p.charge < 0);
  if (refunds.length) {
    const refundTotal = Math.abs(refunds.reduce((s, p) => s + p.charge, 0));
    items.push({
      ic: '↩️', level: 'ok',
      html: `<strong>${refunds.length}</strong> זיכויים בסה"כ <strong>₪${fmt(refundTotal)}</strong> הוחזרו לחשבון`,
    });
  }

  // ── Render ─────────────────────────────────────────────────────────────
  const grid = document.getElementById('insights-grid');
  if (!items.length) {
    grid.innerHTML = `<div class="empty">אין תובנות זמינות.</div>`;
    return;
  }
  grid.innerHTML = items.map(item => `
    <div class="insight-card ${item.level}">
      <span class="ic">${item.ic}</span>
      <span class="body">${item.html}</span>
    </div>
  `).join('');
}

function initFilters() {
  const types = [...new Set(DATA.payments.map(p => p.type).filter(Boolean))].sort();
  const cats = [...new Set(DATA.payments.map(p => p.category).filter(Boolean))].sort();
  document.getElementById('type-filter').insertAdjacentHTML('beforeend', types.map(t => `<option>${esc(t)}</option>`).join(''));
  document.getElementById('category-filter').insertAdjacentHTML('beforeend', cats.map(c => `<option>${esc(c)}</option>`).join(''));
}

// Collapsible sections
document.querySelectorAll('.section > h2').forEach(h => {
  h.addEventListener('click', () => h.parentElement.classList.toggle('collapsed'));
});

document.querySelectorAll('#all-table th[data-k]').forEach(th => {
  th.addEventListener('click', () => {
    const k = th.dataset.k;
    if (sortKey === k) { sortDir = -sortDir; } else { sortKey = k; sortDir = 1; }
    renderAll();
  });
});

['search', 'type-filter', 'category-filter'].forEach(id =>
  document.getElementById(id).addEventListener('input', renderAll));

// -------------------- Selection / floating bar --------------------

function updateFloatingBar() {
  const checked = [...document.querySelectorAll('.row-chk:checked')];
  const bar = document.getElementById('floating-bar');
  if (!checked.length) { bar.classList.remove('visible'); return; }
  const total = checked.reduce((s, cb) => s + parseFloat(cb.dataset.amount || 0), 0);
  document.getElementById('sel-summary').textContent = `נבחרו ${checked.length} פריטים`;
  document.getElementById('sel-total').textContent = `סה"כ ₪${fmt(total)}`;
  document.getElementById('sel-items').innerHTML = checked.map(cb => {
    const amount = parseFloat(cb.dataset.amount || 0);
    const label = cb.dataset.label || '—';
    const cls = amount < 0 ? 'amount-refund' : amount >= DATA.high_threshold ? 'amount-high' : '';
    return `<div class="sel-item">
      <span class="item-label">${label}</span>
      <span class="item-amount ${cls}">₪${fmt(amount)}</span>
    </div>`;
  }).join('');
  bar.classList.add('visible');
}

// -------------------- Save HTML --------------------

function saveHTML() {
  const html = '<!DOCTYPE html>' + document.documentElement.outerHTML;
  const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = (DATA.source || DATA.title || 'payments').replace(/\.[^.]+$/, '') + '.html';
  a.click();
  URL.revokeObjectURL(a.href);
}

document.getElementById('btn-save-html').addEventListener('click', saveHTML);

// -------------------- Save JSON --------------------

function saveJSON() {
  const payload = {
    title: DATA.title,
    source: DATA.source,
    issuer: DATA.issuer,
    payments: DATA.payments,
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = (DATA.source || DATA.title || 'payments').replace(/\.[^.]+$/, '') + '.json';
  a.click();
  URL.revokeObjectURL(a.href);
}

document.getElementById('btn-save').addEventListener('click', saveJSON);

document.addEventListener('change', e => {
  if (e.target.classList.contains('row-chk')) {
    const tbody = e.target.closest('tbody');
    const table = tbody.closest('table');
    const selectAll = table.querySelector('.select-all');
    if (selectAll) {
      const all = tbody.querySelectorAll('.row-chk');
      selectAll.checked = [...all].every(cb => cb.checked);
    }
    updateFloatingBar();
  }
  if (e.target.classList.contains('select-all')) {
    const table = e.target.closest('table');
    table.querySelectorAll('tbody .row-chk').forEach(cb => { cb.checked = e.target.checked; });
    updateFloatingBar();
  }
});

document.getElementById('sel-clear').addEventListener('click', () => {
  document.querySelectorAll('.row-chk:checked').forEach(cb => { cb.checked = false; });
  document.querySelectorAll('.select-all').forEach(cb => { cb.checked = false; });
  updateFloatingBar();
});

initTheme();
initFilters();
renderCards();
renderInsights();
renderFlagged();
renderDuplicates();
renderSubscriptions();
renderInstallments();
renderMerchants();
renderAll();
renderCharts();
</script>
</body>
</html>
"""


def generate_html(data: dict) -> str:
    payload = {
        **data,
        "summary": build_summary(data["payments"]),
        "insights": build_insights(data["payments"]),
        "high_threshold": HIGH_THRESHOLD,
    }
    return HTML_TEMPLATE \
        .replace("__TITLE__", payload["title"]) \
        .replace("__DATA__", json.dumps(payload, ensure_ascii=False))


# ---------------------------------------------------------------------------
# Two-month comparison
# ---------------------------------------------------------------------------


def build_comparison(data_a: dict, data_b: dict) -> dict:
    """Build a diff between two parsed payments files."""
    def by_category(payments):
        d = defaultdict(float)
        for p in payments:
            d[p["category"] or "ללא קטגוריה"] += p["charge"]
        return d

    def by_merchant(payments):
        d = defaultdict(float)
        for p in payments:
            d[p.get("canonical") or p["merchant"]] += p["charge"]
        return d

    cat_a, cat_b = by_category(data_a["payments"]), by_category(data_b["payments"])
    mer_a, mer_b = by_merchant(data_a["payments"]), by_merchant(data_b["payments"])

    categories = sorted(
        [
            {
                "name": c,
                "a": round(cat_a.get(c, 0), 2),
                "b": round(cat_b.get(c, 0), 2),
                "delta": round(cat_b.get(c, 0) - cat_a.get(c, 0), 2),
            }
            for c in set(cat_a) | set(cat_b)
        ],
        key=lambda x: abs(x["delta"]),
        reverse=True,
    )

    merchants = sorted(
        [
            {
                "name": m,
                "a": round(mer_a.get(m, 0), 2),
                "b": round(mer_b.get(m, 0), 2),
                "delta": round(mer_b.get(m, 0) - mer_a.get(m, 0), 2),
            }
            for m in set(mer_a) | set(mer_b)
        ],
        key=lambda x: abs(x["delta"]),
        reverse=True,
    )

    new_merchants = sorted(
        [m for m in merchants if m["a"] == 0 and m["b"] > 0],
        key=lambda x: -x["b"],
    )
    vanished = sorted(
        [m for m in merchants if m["b"] == 0 and m["a"] > 0],
        key=lambda x: -x["a"],
    )

    tot_a = sum(p["charge"] for p in data_a["payments"])
    tot_b = sum(p["charge"] for p in data_b["payments"])

    return {
        "a_title": data_a["title"],
        "a_source": data_a["source"],
        "a_total": round(tot_a, 2),
        "a_count": len(data_a["payments"]),
        "b_title": data_b["title"],
        "b_source": data_b["source"],
        "b_total": round(tot_b, 2),
        "b_count": len(data_b["payments"]),
        "delta": round(tot_b - tot_a, 2),
        "delta_pct": round((tot_b - tot_a) / tot_a * 100, 1) if tot_a else 0,
        "categories": categories,
        "merchants": merchants,
        "new_merchants": new_merchants[:20],
        "vanished": vanished[:20],
    }


COMPARE_HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="he" dir="rtl" data-theme="light">
<head>
<meta charset="UTF-8">
<title>השוואת תקופות</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
  :root {
    --bg: #f5f5f7; --card: #fff; --text: #222; --muted: #666; --soft: #888;
    --border: #eee; --border-strong: #ddd; --hover: #f8f8fb; --th-bg: #fafafa;
    --shadow: 0 1px 3px rgba(0,0,0,0.08);
    --up: #c62828; --down: #2e7d32; --a-color: #2196f3; --b-color: #ef6c00;
  }
  [data-theme="dark"] {
    --bg: #111418; --card: #1c2128; --text: #e6edf3; --muted: #9aa4af; --soft: #7a8591;
    --border: #2a323c; --border-strong: #394350; --hover: #232b36; --th-bg: #1a2028;
    --shadow: 0 1px 3px rgba(0,0,0,0.4);
    --up: #ef5350; --down: #66bb6a; --a-color: #64b5f6; --b-color: #ffb74d;
  }
  * { box-sizing: border-box; }
  body { font-family: -apple-system, "Segoe UI", Arial, sans-serif; margin: 0; padding: 24px;
         background: var(--bg); color: var(--text); transition: background 0.2s, color 0.2s; }
  header { display: flex; justify-content: space-between; align-items: flex-start; gap: 16px; margin-bottom: 20px; }
  header h1 { font-size: 20px; margin: 0 0 6px; }
  .src { color: var(--soft); font-size: 12px; }
  .theme-toggle { background: var(--card); color: var(--text); border: 1px solid var(--border-strong);
                  width: 40px; height: 40px; border-radius: 10px; cursor: pointer; font-size: 18px; box-shadow: var(--shadow); }
  .cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 12px; margin-bottom: 20px; }
  .card { background: var(--card); padding: 16px; border-radius: 10px; box-shadow: var(--shadow); }
  .card .label { font-size: 11px; color: var(--muted); letter-spacing: 0.5px; }
  .card .value { font-size: 22px; font-weight: 600; margin-top: 4px; }
  .card .sub { font-size: 12px; color: var(--soft); margin-top: 2px; }
  .section { background: var(--card); padding: 16px 20px; border-radius: 10px; box-shadow: var(--shadow); margin-bottom: 20px; }
  .section > h2 { margin: 0 0 12px; font-size: 16px; }
  .chart-wrap { position: relative; height: 320px; }
  table { width: 100%; border-collapse: collapse; font-size: 14px; }
  th, td { padding: 8px 10px; border-bottom: 1px solid var(--border); text-align: right; }
  th { background: var(--th-bg); font-weight: 600; font-size: 13px; }
  tr:hover td { background: var(--hover); }
  .num { font-variant-numeric: tabular-nums; white-space: nowrap; }
  .up { color: var(--up); font-weight: 600; }
  .down { color: var(--down); font-weight: 600; }
  .grid-two { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
  @media (max-width: 700px) { .grid-two { grid-template-columns: 1fr; } }
  .empty { color: var(--soft); font-style: italic; padding: 8px 0; }
  .backlink { font-size: 12px; color: var(--muted); text-decoration: none; }
  .backlink:hover { color: var(--a-color); }
</style>
</head>
<body>
<header>
  <div>
    <h1>השוואת תקופות</h1>
    <div class="src"><a class="backlink" href="/">← חזרה להעלאת קובץ</a></div>
  </div>
  <button class="theme-toggle" id="theme-toggle" title="מצב כהה/בהיר">🌙</button>
</header>

<div class="cards" id="cards"></div>

<div class="section">
  <h2>חלוקה לפי ענף</h2>
  <div class="chart-wrap"><canvas id="chart-categories"></canvas></div>
</div>

<div class="section">
  <h2>שינויים לפי ענף</h2>
  <table id="categories-table">
    <thead><tr><th>ענף</th><th>חודש א' (₪)</th><th>חודש ב' (₪)</th><th>שינוי (₪)</th></tr></thead>
    <tbody></tbody>
  </table>
</div>

<div class="section">
  <h2>השינויים הגדולים ביותר לפי בית עסק</h2>
  <div class="grid-two">
    <div>
      <h3 style="font-size:13px;color:var(--muted);margin:0 0 8px;">עליות</h3>
      <table id="increases-table">
        <thead><tr><th>בית עסק</th><th>לפני</th><th>אחרי</th><th>+שינוי</th></tr></thead>
        <tbody></tbody>
      </table>
    </div>
    <div>
      <h3 style="font-size:13px;color:var(--muted);margin:0 0 8px;">ירידות</h3>
      <table id="decreases-table">
        <thead><tr><th>בית עסק</th><th>לפני</th><th>אחרי</th><th>-שינוי</th></tr></thead>
        <tbody></tbody>
      </table>
    </div>
  </div>
</div>

<div class="section">
  <div class="grid-two">
    <div>
      <h2>חדשים בחודש ב' <span class="num" id="new-count" style="color:var(--soft);font-size:12px;"></span></h2>
      <table id="new-table">
        <thead><tr><th>בית עסק</th><th>סכום (₪)</th></tr></thead>
        <tbody></tbody>
      </table>
    </div>
    <div>
      <h2>הפסיקו בחודש ב' <span class="num" id="vanished-count" style="color:var(--soft);font-size:12px;"></span></h2>
      <table id="vanished-table">
        <thead><tr><th>בית עסק</th><th>סכום קודם (₪)</th></tr></thead>
        <tbody></tbody>
      </table>
    </div>
  </div>
</div>

<script>
const DATA = __DATA__;

const fmt = n => new Intl.NumberFormat('he-IL', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(n);
const fmt0 = n => new Intl.NumberFormat('he-IL', { maximumFractionDigits: 0 }).format(n);
const signed = n => (n >= 0 ? '+' : '') + fmt(n);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[c]));
const deltaClass = n => n > 0.01 ? 'up' : n < -0.01 ? 'down' : '';

function applyTheme(t) {
  document.documentElement.dataset.theme = t;
  document.getElementById('theme-toggle').textContent = t === 'dark' ? '☀️' : '🌙';
  renderChart();
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

function renderCards() {
  const deltaCls = deltaClass(DATA.delta);
  const cards = [
    ['חודש א\'', '₪' + fmt(DATA.a_total), DATA.a_count + ' עסקאות · ' + DATA.a_source],
    ['חודש ב\'', '₪' + fmt(DATA.b_total), DATA.b_count + ' עסקאות · ' + DATA.b_source],
    ['שינוי', `<span class="${deltaCls}">${signed(DATA.delta)}</span>`, (DATA.delta_pct >= 0 ? '+' : '') + DATA.delta_pct + '%'],
  ];
  document.getElementById('cards').innerHTML = cards
    .map(([l, v, sub]) => `<div class="card"><div class="label">${esc(l)}</div><div class="value">${v}</div><div class="sub">${esc(sub)}</div></div>`)
    .join('');
}

let chart;
function renderChart() {
  if (typeof Chart === 'undefined') return;
  if (chart) chart.destroy();
  const text = getComputedStyle(document.documentElement).getPropertyValue('--text').trim() || '#222';
  const grid = getComputedStyle(document.documentElement).getPropertyValue('--border').trim() || '#eee';
  const aColor = getComputedStyle(document.documentElement).getPropertyValue('--a-color').trim() || '#2196f3';
  const bColor = getComputedStyle(document.documentElement).getPropertyValue('--b-color').trim() || '#ef6c00';
  Chart.defaults.color = text;
  Chart.defaults.borderColor = grid;

  const cats = DATA.categories.slice(0, 12);
  chart = new Chart(document.getElementById('chart-categories'), {
    type: 'bar',
    data: {
      labels: cats.map(c => c.name),
      datasets: [
        { label: 'חודש א\'', data: cats.map(c => c.a), backgroundColor: aColor },
        { label: 'חודש ב\'', data: cats.map(c => c.b), backgroundColor: bColor },
      ],
    },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: 'top' } },
      scales: { x: { ticks: { callback: v => '₪' + fmt0(v) } } },
    },
  });
}

function renderCategoriesTable() {
  document.querySelector('#categories-table tbody').innerHTML = DATA.categories.map(c => `
    <tr>
      <td>${esc(c.name)}</td>
      <td class="num">${fmt(c.a)}</td>
      <td class="num">${fmt(c.b)}</td>
      <td class="num ${deltaClass(c.delta)}">${signed(c.delta)}</td>
    </tr>
  `).join('') || `<tr><td colspan="4" class="empty">אין נתונים</td></tr>`;
}

function renderChanges() {
  const increases = DATA.merchants.filter(m => m.delta > 0.01).slice(0, 15);
  const decreases = [...DATA.merchants].filter(m => m.delta < -0.01).sort((a, b) => a.delta - b.delta).slice(0, 15);

  document.querySelector('#increases-table tbody').innerHTML = increases.map(m => `
    <tr>
      <td>${esc(m.name)}</td>
      <td class="num">${fmt(m.a)}</td>
      <td class="num">${fmt(m.b)}</td>
      <td class="num up">${signed(m.delta)}</td>
    </tr>
  `).join('') || `<tr><td colspan="4" class="empty">—</td></tr>`;

  document.querySelector('#decreases-table tbody').innerHTML = decreases.map(m => `
    <tr>
      <td>${esc(m.name)}</td>
      <td class="num">${fmt(m.a)}</td>
      <td class="num">${fmt(m.b)}</td>
      <td class="num down">${signed(m.delta)}</td>
    </tr>
  `).join('') || `<tr><td colspan="4" class="empty">—</td></tr>`;
}

function renderNewVanished() {
  document.getElementById('new-count').textContent = `(${DATA.new_merchants.length})`;
  document.getElementById('vanished-count').textContent = `(${DATA.vanished.length})`;
  document.querySelector('#new-table tbody').innerHTML = DATA.new_merchants.map(m => `
    <tr><td>${esc(m.name)}</td><td class="num">${fmt(m.b)}</td></tr>
  `).join('') || `<tr><td colspan="2" class="empty">—</td></tr>`;
  document.querySelector('#vanished-table tbody').innerHTML = DATA.vanished.map(m => `
    <tr><td>${esc(m.name)}</td><td class="num">${fmt(m.a)}</td></tr>
  `).join('') || `<tr><td colspan="2" class="empty">—</td></tr>`;
}

initTheme();
renderCards();
renderCategoriesTable();
renderChanges();
renderNewVanished();
renderChart();
</script>
</body>
</html>
"""


def generate_comparison_html(data_a: dict, data_b: dict) -> str:
    comparison = build_comparison(data_a, data_b)
    return COMPARE_HTML_TEMPLATE.replace("__DATA__", json.dumps(comparison, ensure_ascii=False))


# ---------------------------------------------------------------------------
# Multi-month comparison (up to 12 files)
# ---------------------------------------------------------------------------

def build_multi(months_data: list[dict], month_urls: list[str] | None = None) -> dict:
    """Build a multi-month comparison payload from a list of parsed payment dicts."""

    # Sort months by their earliest transaction date
    def _month_key(d):
        dates = [p["date"] for p in d["payments"] if p["date"]]
        return min(dates) if dates else d.get("source", "")

    # Sort both months_data and month_urls together
    if month_urls and len(month_urls) == len(months_data):
        paired = sorted(zip(months_data, month_urls), key=lambda x: _month_key(x[0]))
        months_data, month_urls = [p[0] for p in paired], [p[1] for p in paired]
    else:
        months_data = sorted(months_data, key=_month_key)
        month_urls = None

    months = []
    all_cat_names: set = set()
    all_mer_names: set = set()

    for d in months_data:
        pays = d["payments"]
        total = round(sum(p["charge"] for p in pays), 2)
        cat_totals: dict = defaultdict(float)
        mer_totals: dict = defaultdict(float)
        for p in pays:
            cat_totals[p["category"] or "ללא קטגוריה"] += p["charge"]
            mer_totals[p.get("canonical") or p["merchant"]] += p["charge"]
        all_cat_names |= set(cat_totals.keys())
        all_mer_names |= set(mer_totals.keys())
        months.append({
            "label": d.get("title") or d.get("source") or "—",
            "source": d.get("source", ""),
            "total": total,
            "count": len(pays),
            "cat": {k: round(v, 2) for k, v in cat_totals.items()},
            "mer": {k: round(v, 2) for k, v in mer_totals.items()},
            "url": month_urls[len(months)] if month_urls else "",
        })

    # Top categories by total across all months
    cat_grand = {c: sum(m["cat"].get(c, 0) for m in months) for c in all_cat_names}
    top_cats = [k for k, _ in sorted(cat_grand.items(), key=lambda x: -x[1])][:10]

    # Category matrix: list of {name, totals:[...per month], grand_total}
    cat_matrix = [
        {
            "name": c,
            "totals": [round(m["cat"].get(c, 0), 2) for m in months],
            "grand": round(cat_grand[c], 2),
        }
        for c in top_cats
    ]

    # Top merchants by total across all months
    mer_grand = {m: sum(mo["mer"].get(m, 0) for mo in months) for m in all_mer_names}
    top_mers = [k for k, _ in sorted(mer_grand.items(), key=lambda x: -x[1])][:20]

    mer_matrix = [
        {
            "name": m,
            "totals": [round(mo["mer"].get(m, 0), 2) for mo in months],
            "grand": round(mer_grand[m], 2),
            "last_delta": round(
                months[-1]["mer"].get(m, 0) - months[-2]["mer"].get(m, 0), 2
            ) if len(months) >= 2 else 0,
        }
        for m in top_mers
    ]

    grand_total = round(sum(m["total"] for m in months), 2)
    avg = round(grand_total / len(months), 2) if months else 0
    peak = max(months, key=lambda m: m["total"]) if months else None
    low  = min(months, key=lambda m: m["total"]) if months else None

    return {
        "months": months,
        "cat_matrix": cat_matrix,
        "mer_matrix": mer_matrix,
        "grand_total": grand_total,
        "avg_monthly": avg,
        "peak_label": peak["label"] if peak else "",
        "peak_total": peak["total"] if peak else 0,
        "low_label":  low["label"]  if low  else "",
        "low_total":  low["total"]  if low  else 0,
        "high_threshold": HIGH_THRESHOLD,
    }


MULTI_HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="he" dir="rtl" data-theme="light">
<head>
<meta charset="UTF-8">
<title>השוואת חודשים מרובים</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
  :root {
    --bg:#f5f5f7;--card:#fff;--text:#222;--muted:#666;--soft:#888;
    --border:#eee;--border-strong:#ddd;--hover:#f8f8fb;--th-bg:#fafafa;
    --shadow:0 1px 3px rgba(0,0,0,.08);
    --primary:#2196f3;--up:#c62828;--down:#2e7d32;
  }
  [data-theme="dark"]{
    --bg:#111418;--card:#1c2128;--text:#e6edf3;--muted:#9aa4af;--soft:#7a8591;
    --border:#2a323c;--border-strong:#394350;--hover:#232b36;--th-bg:#1a2028;
    --shadow:0 1px 3px rgba(0,0,0,.4);
    --primary:#64b5f6;--up:#ef5350;--down:#66bb6a;
  }
  *{box-sizing:border-box;}
  body{font-family:-apple-system,"Segoe UI",Arial,sans-serif;margin:0;padding:24px;
       background:var(--bg);color:var(--text);transition:background .2s,color .2s;}
  header{display:flex;justify-content:space-between;align-items:center;gap:16px;margin-bottom:20px;}
  header h1{font-size:20px;margin:0;}
  .btn{background:var(--card);color:var(--text);border:1px solid var(--border-strong);
       height:36px;border-radius:8px;cursor:pointer;font-size:13px;font-weight:600;
       padding:0 14px;box-shadow:var(--shadow);text-decoration:none;display:inline-flex;align-items:center;}
  .btn:hover{background:var(--hover);}
  .cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;margin-bottom:20px;}
  .card{background:var(--card);padding:16px;border-radius:10px;box-shadow:var(--shadow);}
  .card .lbl{font-size:11px;color:var(--muted);letter-spacing:.5px;}
  .card .val{font-size:22px;font-weight:600;margin-top:4px;}
  .card .sub{font-size:12px;color:var(--soft);margin-top:2px;}
  .section{background:var(--card);padding:16px 20px;border-radius:10px;box-shadow:var(--shadow);margin-bottom:20px;}
  .section>h2{margin:0 0 14px;font-size:16px;cursor:pointer;user-select:none;}
  .section.collapsed>:not(h2){display:none;}
  .section>h2::before{content:"▾";display:inline-block;width:1em;font-size:11px;color:var(--muted);transform:scaleX(-1);}
  .section.collapsed>h2::before{content:"▸";}
  .chart-wrap{position:relative;height:300px;}
  .chart-wrap.tall{height:380px;}
  table{width:100%;border-collapse:collapse;font-size:13px;}
  th,td{padding:7px 10px;border-bottom:1px solid var(--border);text-align:right;white-space:nowrap;}
  th{background:var(--th-bg);font-weight:600;font-size:12px;position:sticky;top:0;z-index:2;}
  tr:hover td{background:var(--hover);}
  /* Sticky first column — name stays visible while scrolling horizontally */
  td.name,th.name-hdr{
    position:sticky;right:0;background:var(--card);z-index:3;
    min-width:130px;max-width:180px;white-space:normal;
    border-left:1px solid var(--border-strong);
    font-weight:600;
  }
  tr:hover td.name{background:var(--hover);}
  .sum-row td.name{background:var(--th-bg);}
  th.name-hdr{z-index:4;background:var(--th-bg);}
  td.num-month{min-width:88px;color:var(--muted);font-size:12px;}
  td.num-month.has-val{color:var(--text);}
  .num{font-variant-numeric:tabular-nums;}
  .up{color:var(--up);font-weight:700;}
  .down{color:var(--down);font-weight:700;}
  .flat{color:var(--soft);}
  .sum-row td{border-top:2px solid var(--border-strong);border-bottom:none;font-weight:700;background:var(--th-bg);}
  .grand{font-weight:700;color:var(--primary);}
  .empty{color:var(--soft);font-style:italic;padding:8px 0;}
  .tbl-wrap{overflow-x:auto;border-radius:6px;}
  .insights-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(270px,1fr));gap:10px;}
  .ic{display:flex;align-items:flex-start;gap:10px;padding:10px 14px;border-radius:8px;
      border-right:4px solid transparent;background:var(--bg);font-size:14px;line-height:1.5;}
  .ic.ok{border-color:#43a047;} .ic.warn{border-color:#fb8c00;} .ic.alert{border-color:var(--up);}
  .ic .emoji{font-size:20px;flex-shrink:0;margin-top:1px;}
  .ic .body{flex:1;} .ic .body strong{color:var(--primary);font-weight:700;}
</style>
</head>
<body>
<header>
  <h1>השוואת חודשים מרובים</h1>
  <div style="display:flex;gap:8px;">
    <a href="/" class="btn">← קובץ חדש</a>
    <a href="/multi" class="btn">📂 העלאה חדשה</a>
    <button class="btn" id="btn-save-html">💾 שמור HTML</button>
    <button class="btn" id="theme-toggle">🌙</button>
  </div>
</header>

<div class="cards" id="cards"></div>

<div class="section" style="display:flex;gap:24px;align-items:flex-start;flex-wrap:wrap;">
  <div style="flex:1;min-width:260px;">
    <h2 style="font-size:16px;margin:0 0 14px;">חלוקת הוצאות לפי קטגוריה</h2>
    <div style="position:relative;height:300px;"><canvas id="chart-cat-doughnut"></canvas></div>
  </div>
  <div style="flex:1;min-width:220px;">
    <h2 style="font-size:16px;margin:0 0 14px;">סיכום חודשים</h2>
    <table id="month-summary-table" style="width:100%;border-collapse:collapse;font-size:14px;">
      <thead><tr>
        <th style="text-align:right;padding:6px 8px;border-bottom:2px solid var(--border-strong);font-size:12px;">חודש</th>
        <th style="text-align:right;padding:6px 8px;border-bottom:2px solid var(--border-strong);font-size:12px;">עסקאות</th>
        <th style="text-align:right;padding:6px 8px;border-bottom:2px solid var(--border-strong);font-size:12px;">סה"כ (₪)</th>
      </tr></thead>
      <tbody id="month-summary-body"></tbody>
    </table>
  </div>
</div>

<div class="section" id="sec-insights">
  <h2>תובנות 🧠</h2>
  <div class="insights-grid" id="insights-grid"></div>
</div>

<div class="section">
  <h2>טבלת קטגוריות לפי חודש</h2>
  <div class="tbl-wrap"><table id="cat-table"></table></div>
</div>

<div class="section">
  <h2>Top בתי עסק לפי חודש</h2>
  <input id="mer-search" type="text" placeholder="חיפוש בית עסק..."
    style="width:100%;padding:8px 12px;margin-bottom:12px;border:1px solid var(--border-strong);
           border-radius:6px;font-size:14px;font-family:inherit;background:var(--card);color:var(--text);">
  <div class="tbl-wrap"><table id="mer-table"></table></div>
</div>

<script>
const DATA = __DATA__;
const fmt  = n => new Intl.NumberFormat('he-IL',{minimumFractionDigits:2,maximumFractionDigits:2}).format(n);
const fmt0 = n => new Intl.NumberFormat('he-IL',{maximumFractionDigits:0}).format(n);
const esc  = s => String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));

const COLORS = ['#2196f3','#ef6c00','#43a047','#8e24aa','#e53935','#00acc1','#fb8c00','#6d4c41','#546e7a','#d81b60','#7cb342','#3949ab'];

let charts = {};
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
    localStorage.setItem('payments-theme',next); applyTheme(next);
  });
}

// ── Cards ──────────────────────────────────────────────────────────────────
function renderCards(){
  const items=[
    ['חודשים',DATA.months.length,'קבצים שהועלו'],
    ['סה"כ כולל','₪'+fmt(DATA.grand_total),DATA.months.reduce((s,m)=>s+m.count,0)+' עסקאות'],
    ['ממוצע חודשי','₪'+fmt(DATA.avg_monthly),''],
    ['חודש יקר ביותר','₪'+fmt(DATA.peak_total),esc(DATA.peak_label)],
    ['חודש זול ביותר','₪'+fmt(DATA.low_total),esc(DATA.low_label)],
  ];
  document.getElementById('cards').innerHTML=items.map(([l,v,s])=>
    `<div class="card"><div class="lbl">${esc(l)}</div><div class="val">${v}</div><div class="sub">${s}</div></div>`
  ).join('');
}

// ── Charts ─────────────────────────────────────────────────────────────────
function renderCharts(){
  if(typeof Chart==='undefined') return;
  destroyCharts();
  const text=getComputedStyle(document.documentElement).getPropertyValue('--text').trim()||'#222';
  const grid=getComputedStyle(document.documentElement).getPropertyValue('--border').trim()||'#eee';
  Chart.defaults.color=text; Chart.defaults.borderColor=grid;

  // Category doughnut — aggregate totals across all months
  const cats=DATA.cat_matrix;
  if(cats.length){
    charts.doughnut=new Chart(document.getElementById('chart-cat-doughnut'),{
      type:'doughnut',
      data:{
        labels:cats.map(c=>c.name),
        datasets:[{data:cats.map(c=>c.grand),backgroundColor:COLORS.slice(0,cats.length),borderWidth:2}],
      },
      options:{responsive:true,maintainAspectRatio:false,
        plugins:{
          legend:{position:'right',labels:{boxWidth:13,font:{size:12},padding:10}},
          tooltip:{callbacks:{label:ctx=>{
            const pct=Math.round(ctx.parsed/DATA.grand_total*100);
            return ` ${ctx.label}: ₪${fmt(ctx.parsed)} (${pct}%)`;
          }}},
        },
      },
    });
  }

  // Month summary table
  const tbody=document.getElementById('month-summary-body');
  if(tbody){
    const avg=DATA.avg_monthly;
    tbody.innerHTML=DATA.months.map(m=>{
      const diff=m.total-avg;
      const diffStr=diff>=0?`<span style="color:var(--up);font-size:11px;">+₪${fmt0(diff)}</span>`
                          :`<span style="color:var(--down);font-size:11px;">-₪${fmt0(Math.abs(diff))}</span>`;
      const labelCell=m.url
        ? `<a href="${m.url}" target="_blank" style="color:var(--primary);text-decoration:none;font-weight:600;display:flex;align-items:center;gap:5px;">
             ${esc(m.label)} <span style="font-size:10px;opacity:0.7;">↗</span>
           </a>`
        : esc(m.label);
      return `<tr style="${m.url?'cursor:pointer;':''}">
        <td style="padding:6px 8px;border-bottom:1px solid var(--border);">${labelCell}</td>
        <td style="padding:6px 8px;border-bottom:1px solid var(--border);font-variant-numeric:tabular-nums;">${m.count}</td>
        <td style="padding:6px 8px;border-bottom:1px solid var(--border);font-variant-numeric:tabular-nums;">₪${fmt(m.total)} ${diffStr}</td>
      </tr>`;
    }).join('') + `<tr style="font-weight:700;background:var(--th-bg);">
      <td style="padding:6px 8px;">סה"כ</td>
      <td style="padding:6px 8px;">${DATA.months.reduce((s,m)=>s+m.count,0)}</td>
      <td style="padding:6px 8px;">₪${fmt(DATA.grand_total)}</td>
    </tr>`;
  }
}

// ── Category table ──────────────────────────────────────────────────────────
function shortLabel(label){
  // "דף פירוט דיגיטלי כאל 01-25" → "01-25", or keep as-is if short
  const m = label.match(/(\d{2}-\d{2,4})$/);
  return m ? m[1] : (label.length > 10 ? label.slice(-7) : label);
}

function renderCatTable(){
  const months=DATA.months; const cats=DATA.cat_matrix;
  if(!cats.length){document.getElementById('cat-table').innerHTML='<tr><td class="empty">אין נתונים</td></tr>';return;}
  const hdrs=`<th class="name-hdr">קטגוריה</th>${months.map(m=>`<th title="${esc(m.label)}">${esc(shortLabel(m.label))}</th>`).join('')}<th class="grand">סה"כ</th>`;
  const rows=cats.map(c=>{
    const cells=c.totals.map(t=>`<td class="num num-month ${t>0?'has-val':''}">${t>0?fmt(t):'—'}</td>`).join('');
    return `<tr><td class="name">${esc(c.name)}</td>${cells}<td class="num grand">${fmt(c.grand)}</td></tr>`;
  });
  const colTotals=months.map((_,i)=>cats.reduce((s,c)=>s+c.totals[i],0));
  const grandSum=cats.reduce((s,c)=>s+c.grand,0);
  const sumRow=`<tr class="sum-row"><td class="name">סה"כ</td>${colTotals.map(t=>`<td class="num">${fmt(t)}</td>`).join('')}<td class="num grand">${fmt(grandSum)}</td></tr>`;
  document.getElementById('cat-table').innerHTML=`<thead><tr>${hdrs}</tr></thead><tbody>${rows.join('')}${sumRow}</tbody>`;
}

// ── Merchant table ──────────────────────────────────────────────────────────
function renderMerTable(q=''){
  const months=DATA.months;
  let mers=DATA.mer_matrix;
  if(q) mers=mers.filter(m=>m.name.toLowerCase().includes(q.toLowerCase()));
  if(!mers.length){document.getElementById('mer-table').innerHTML=`<tr><td colspan="${months.length+3}" class="empty">${q?'לא נמצאו תוצאות':'אין נתונים'}</td></tr>`;return;}
  const hdrs=`<th class="name-hdr">בית עסק</th>${months.map(m=>`<th title="${esc(m.label)}">${esc(shortLabel(m.label))}</th>`).join('')}<th class="grand">סה"כ</th><th>מגמה</th>`;
  const rows=mers.map(m=>{
    const cells=m.totals.map(t=>`<td class="num num-month ${t>0?'has-val':''}">${t>0?fmt(t):'—'}</td>`).join('');
    const d=m.last_delta;
    const trend=DATA.months.length<2?'<span class="flat">—</span>':
      d>0.5?`<span class="up">↑ ₪${fmt0(Math.abs(d))}</span>`:
      d<-0.5?`<span class="down">↓ ₪${fmt0(Math.abs(d))}</span>`:'<span class="flat">→</span>';
    return `<tr><td class="name">${esc(m.name)}</td>${cells}<td class="num grand">${fmt(m.grand)}</td><td>${trend}</td></tr>`;
  });
  const colTotals=months.map((_,i)=>mers.reduce((s,m)=>s+m.totals[i],0));
  const grandSum=mers.reduce((s,m)=>s+m.grand,0);
  const sumRow=`<tr class="sum-row"><td class="name">סה"כ</td>${colTotals.map(t=>`<td class="num">${fmt(t)}</td>`).join('')}<td class="num grand">${fmt(grandSum)}</td><td></td></tr>`;
  document.getElementById('mer-table').innerHTML=`<thead><tr>${hdrs}</tr></thead><tbody>${rows.join('')}${sumRow}</tbody>`;
}

// ── Insights ────────────────────────────────────────────────────────────────
function renderInsights(){
  const months=DATA.months; const cats=DATA.cat_matrix; const mers=DATA.mer_matrix;
  const n=months.length;
  const items=[];
  if(n<2) return;

  const totals=months.map(m=>m.total);
  const avg=DATA.avg_monthly;

  // ── 1. Annual projection ────────────────────────────────────────────────
  const annualProjection=Math.round(avg*12);
  items.push({l:'ok',e:'🗓️',
    h:`קצב שנתי משוער: <strong>₪${fmt(annualProjection)}</strong> (ממוצע ₪${fmt(avg)}/חודש על בסיס ${n} חודשים)`});

  // ── 2. Spending trend (linear regression slope) ─────────────────────────
  if(n>=3){
    const xs=months.map((_,i)=>i), meanX=(n-1)/2, meanY=avg;
    const num=xs.reduce((s,x,i)=>s+(x-meanX)*(totals[i]-meanY),0);
    const den=xs.reduce((s,x)=>s+(x-meanX)**2,0);
    const slope=den?num/den:0;
    const pctPerMonth=Math.abs(Math.round(slope/avg*100));
    if(Math.abs(slope)>avg*0.02){
      items.push({l:slope>0?'warn':'ok',e:slope>0?'📈':'📉',
        h:slope>0
          ? `מגמת <strong>עלייה</strong> ב-${pctPerMonth}% לחודש לאורך התקופה — ₪${fmt0(Math.abs(slope))} יותר בכל חודש בממוצע`
          : `מגמת <strong>ירידה</strong> ב-${pctPerMonth}% לחודש לאורך התקופה — ₪${fmt0(Math.abs(slope))} פחות בכל חודש בממוצע`});
    } else {
      items.push({l:'ok',e:'➡️',h:`ההוצאה <strong>יציבה יחסית</strong> לאורך התקופה — שינוי ממוצע של ₪${fmt0(Math.abs(slope))} בלבד לחודש`});
    }
  }

  // ── 3. Last 3 months vs first 3 months ─────────────────────────────────
  if(n>=6){
    const firstHalf=totals.slice(0,3).reduce((a,b)=>a+b,0)/3;
    const lastHalf=totals.slice(-3).reduce((a,b)=>a+b,0)/3;
    const delta=lastHalf-firstHalf, pct=Math.round(Math.abs(delta)/firstHalf*100);
    items.push({l:delta>firstHalf*0.15?'warn':delta<-firstHalf*0.1?'ok':'ok',
      e:delta>0?'📊':'📊',
      h:`3 חודשים אחרונים לעומת 3 ראשונים: ממוצע <strong>${delta>=0?'+':''}₪${fmt0(delta)}</strong> לחודש (${delta>=0?'+':''}${pct}%)`});
  }

  // ── 4. Outlier months (>1.5 std dev from mean) ─────────────────────────
  if(n>=4){
    const variance=totals.reduce((s,t)=>s+(t-avg)**2,0)/n;
    const std=Math.sqrt(variance);
    const outliers=months.filter(m=>Math.abs(m.total-avg)>1.5*std)
      .sort((a,b)=>Math.abs(b.total-avg)-Math.abs(a.total-avg));
    if(outliers.length){
      const o=outliers[0], diff=o.total-avg;
      items.push({l:'warn',e:'⚡',
        h:`חודש חריג: <strong>${esc(shortLabel(o.label))}</strong> — ₪${fmt(o.total)} (${diff>0?'+':''}₪${fmt0(diff)} מהממוצע, ${Math.round(Math.abs(diff)/std*10)/10} סטיות תקן)`});
    }
  }

  // ── 5. Merchants present in ALL months (loyal/fixed expenses) ───────────
  if(n>=3){
    const loyal=mers.filter(m=>m.totals.every(t=>t>0));
    const loyalTotal=loyal.reduce((s,m)=>s+m.grand,0);
    const loyalPct=Math.round(loyalTotal/DATA.grand_total*100);
    if(loyal.length){
      items.push({l:loyalPct>40?'warn':'ok',e:'🔒',
        h:`<strong>${loyal.length} בתי עסק</strong> חויבו בכל ${n} החודשים — ₪${fmt(loyalTotal)} סה"כ (${loyalPct}% מההוצאה הכוללת).<br>
           <span style="font-size:12px;color:var(--muted)">${loyal.slice(0,4).map(m=>esc(m.name)).join(' · ')}${loyal.length>4?' ועוד…':''}</span>`});
    }
  }

  // ── 6. New merchants in last month (didn't appear before) ──────────────
  if(n>=2){
    const prevMonthsMers=new Set(mers.flatMap(m=>
      m.totals.slice(0,-1).some(t=>t>0)?[m.name]:[]));
    const newMers=mers.filter(m=>m.totals[n-1]>0 && !prevMonthsMers.has(m.name));
    if(newMers.length){
      const newTotal=newMers.reduce((s,m)=>s+m.totals[n-1],0);
      items.push({l:'ok',e:'🆕',
        h:`<strong>${newMers.length} בתי עסק חדשים</strong> בחודש האחרון — ₪${fmt(newTotal)} סה"כ.<br>
           <span style="font-size:12px;color:var(--muted)">${newMers.slice(0,4).map(m=>esc(m.name)).join(' · ')}${newMers.length>4?' ועוד…':''}</span>`});
    }
  }

  // ── 7. Vanished merchants (were in first half, gone in last month) ──────
  if(n>=3){
    const hadBefore=mers.filter(m=>m.totals.slice(0,-2).some(t=>t>0) && m.totals[n-1]===0);
    if(hadBefore.length){
      items.push({l:'ok',e:'👻',
        h:`<strong>${hadBefore.length} בתי עסק</strong> שנעלמו בחודש האחרון — לא חויבו לאחרונה.<br>
           <span style="font-size:12px;color:var(--muted)">${hadBefore.slice(0,4).map(m=>esc(m.name)).join(' · ')}${hadBefore.length>4?' ועוד…':''}</span>`});
    }
  }

  // ── 8. Category that grew the most over the full period ────────────────
  if(n>=3 && cats.length){
    let maxGrowth=-Infinity, growthCat='', growthFrom=0, growthTo=0;
    for(const c of cats){
      const first=c.totals.find(t=>t>0)||0;
      const last=c.totals[n-1];
      if(first>50 && last>0){
        const growth=(last-first)/first;
        if(growth>maxGrowth){maxGrowth=growth;growthCat=c.name;growthFrom=first;growthTo=last;}
      }
    }
    if(growthCat && maxGrowth>0.2){
      items.push({l:maxGrowth>0.5?'warn':'ok',e:'🚀',
        h:`הקטגוריה עם הצמיחה הגדולה ביותר: <strong>${esc(growthCat)}</strong> — מ-₪${fmt0(growthFrom)} ל-₪${fmt0(growthTo)} (+${Math.round(maxGrowth*100)}% מחודש ראשון לאחרון)`});
    }
  }

  // ── 9. Subscription price creep (standing orders that increased) ────────
  if(n>=3){
    let creepMer='', creepFrom=0, creepTo=0;
    for(const m of mers){
      // look for merchants that appear regularly and the last value is notably higher
      const active=m.totals.filter(t=>t>0);
      if(active.length>=Math.floor(n*0.6) && active.length>=3){
        const firstActive=active[0], lastActive=active[active.length-1];
        const creep=(lastActive-firstActive)/firstActive;
        if(creep>0.15 && lastActive-firstActive > creepTo-creepFrom){
          creepMer=m.name; creepFrom=firstActive; creepTo=lastActive;
        }
      }
    }
    if(creepMer){
      items.push({l:'warn',e:'💸',
        h:`זחילת מחיר: <strong>${esc(creepMer)}</strong> — מ-₪${fmt(creepFrom)} ל-₪${fmt(creepTo)} (+${Math.round((creepTo-creepFrom)/creepFrom*100)}%) לאורך התקופה`});
    }
  }

  // ── 10. Most volatile category (highest coefficient of variation) ───────
  if(n>=3 && cats.length){
    let maxCV=0, volCat='', volStd=0, volMean=0;
    for(const c of cats){
      const active=c.totals.filter(t=>t>0);
      if(active.length<2) continue;
      const mean=active.reduce((a,b)=>a+b,0)/active.length;
      const std=Math.sqrt(active.reduce((s,t)=>s+(t-mean)**2,0)/active.length);
      const cv=mean>50?std/mean:0;
      if(cv>maxCV){maxCV=cv;volCat=c.name;volStd=std;volMean=mean;}
    }
    if(volCat && maxCV>0.3){
      items.push({l:'ok',e:'🌊',
        h:`הקטגוריה הכי תנודתית: <strong>${esc(volCat)}</strong> — סטיית תקן ₪${fmt0(volStd)} סביב ממוצע ₪${fmt0(volMean)} לחודש`});
    }
  }

  document.getElementById('insights-grid').innerHTML=items.map(it=>
    `<div class="ic ${it.l}"><span class="emoji">${it.e}</span><span class="body">${it.h}</span></div>`
  ).join('');
}

// ── Collapsible sections ────────────────────────────────────────────────────
document.querySelectorAll('.section>h2').forEach(h=>{
  h.addEventListener('click',()=>h.parentElement.classList.toggle('collapsed'));
});

document.getElementById('mer-search').addEventListener('input', e => {
  renderMerTable(e.target.value.trim());
});

document.getElementById('btn-save-html').addEventListener('click', () => {
  const html = '<!DOCTYPE html>' + document.documentElement.outerHTML;
  const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  const months = DATA.months.map(m => shortLabel(m.label)).join('_');
  a.download = `השוואה_${months}.html`;
  a.click();
  URL.revokeObjectURL(a.href);
});

initTheme();
renderCards();
renderInsights();
renderCatTable();
renderMerTable();
</script>
</body>
</html>
"""


def generate_multi_html(months_data: list[dict], month_urls: list[str] | None = None) -> str:
    payload = build_multi(months_data, month_urls=month_urls)
    return MULTI_HTML_TEMPLATE.replace("__DATA__", json.dumps(payload, ensure_ascii=False))


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 1

    xlsx_path = Path(sys.argv[1])
    if not xlsx_path.exists():
        print(f"error: file not found: {xlsx_path}", file=sys.stderr)
        return 1

    out_path = Path(sys.argv[2]) if len(sys.argv) > 2 and not sys.argv[2].startswith("--") else xlsx_path.with_suffix(".html")

    data = parse_payments(xlsx_path)
    out_path.write_text(generate_html(data), encoding="utf-8")

    s = build_summary(data["payments"])
    ins = build_insights(data["payments"])
    print(f"parsed {s['total_count']} payments  ·  total ₪{s['total_amount']:,.2f}")
    print(f"  regular      : {s['regular_count']:3d}  ₪{s['regular_amount']:,.2f}")
    print(f"  non-regular  : {s['non_regular_count']:3d}  ₪{s['non_regular_amount']:,.2f}")
    print(f"  high (≥{HIGH_THRESHOLD})  : {s['high_count']:3d}  ₪{s['high_amount']:,.2f}")
    print(f"  duplicates   : {len(ins['duplicates'])}")
    print(f"  subscriptions: {len(ins['subscriptions'])}")
    print(f"  installments : {len(ins['installments'])}  (future ₪{ins['total_installment_remaining']:,.2f})")
    print(f"written: {out_path}")

    if "--open" in sys.argv:
        webbrowser.open(out_path.resolve().as_uri())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
