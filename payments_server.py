#!/usr/bin/env python3
"""
payments_server.py — tiny Flask server that lets you upload a payments
Excel file through a web form and renders it using the payments_ui UI.

Usage:
    python3 payments_server.py            # http://127.0.0.1:5000
    python3 payments_server.py 8080       # custom port
"""
from __future__ import annotations

import json
import os
import sys
import tempfile
import time
import uuid
from pathlib import Path

from flask import Flask, request, redirect

from payments_ui import parse_payments, generate_html, generate_comparison_html, generate_multi_html
from bank_ui import parse_bank_statement, generate_bank_html

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB

# Filesystem-based result cache — shared across all gunicorn workers
_CACHE_DIR = Path(tempfile.gettempdir()) / "payments_results"
_CACHE_DIR.mkdir(exist_ok=True)
_CACHE_TTL = 3600  # 1 hour


def _cache_write(result_id: str, html: str) -> None:
    (_CACHE_DIR / f"{result_id}.html").write_text(html, encoding="utf-8")


def _cache_read(result_id: str) -> str | None:
    p = _CACHE_DIR / f"{result_id}.html"
    if not p.exists():
        return None
    # Expire old results
    if time.time() - p.stat().st_mtime > _CACHE_TTL:
        p.unlink(missing_ok=True)
        return None
    return p.read_text(encoding="utf-8")


def _cache_cleanup() -> None:
    """Remove results older than TTL. Called opportunistically."""
    try:
        now = time.time()
        for p in _CACHE_DIR.glob("*.html"):
            if now - p.stat().st_mtime > _CACHE_TTL:
                p.unlink(missing_ok=True)
    except OSError:
        pass

UPLOAD_FORM = """<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
<meta charset="UTF-8">
<title>Payments — ניתוח הוצאות</title>
<style>
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:-apple-system,"Segoe UI",Arial,sans-serif;background:#042C53;
     display:flex;align-items:center;justify-content:center;min-height:100vh;padding:20px;}
.page{background:#0C447C;border-radius:16px;padding:28px 24px;max-width:480px;width:100%;overflow:hidden;}
.topbar{display:flex;align-items:center;justify-content:space-between;margin-bottom:22px;}
.logo{display:flex;align-items:center;gap:10px;}
.logo-icon{width:32px;height:32px;background:#378ADD;border-radius:8px;display:flex;align-items:center;justify-content:center;}
.logo-icon svg{width:16px;height:16px;stroke:#E6F1FB;fill:none;stroke-width:2;}
.logo-name{font-size:16px;font-weight:600;color:#E6F1FB;}
.logo-sub{font-size:11px;color:#85B7EB;margin-top:1px;}
.ver{background:#185FA5;color:#B5D4F4;font-size:11px;padding:4px 10px;border-radius:20px;}
.headline{font-size:19px;font-weight:600;color:#E6F1FB;margin-bottom:3px;}
.subline{font-size:13px;color:#85B7EB;margin-bottom:18px;}
.err{background:#4A1B0C;color:#F5C4B3;padding:10px 14px;border-radius:8px;font-size:13px;margin-bottom:14px;}
.cards{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:12px;}
.feat{border-radius:12px;padding:14px 16px;cursor:pointer;outline:2px solid transparent;outline-offset:-2px;transition:outline-color .15s;}
.feat:hover{outline-color:rgba(255,255,255,.25);}
.feat.active{outline-color:rgba(255,255,255,.5);}
.feat.blue{background:#185FA5;}
.feat.teal{background:#0F6E56;}
.feat.purple{background:#534AB7;}
.feat.coral{background:#993C1D;}
.feat-icon{width:30px;height:30px;border-radius:7px;display:flex;align-items:center;justify-content:center;margin-bottom:9px;}
.feat.blue .feat-icon{background:#378ADD;}
.feat.teal .feat-icon{background:#1D9E75;}
.feat.purple .feat-icon{background:#7F77DD;}
.feat.coral .feat-icon{background:#D85A30;}
.feat-icon svg{width:15px;height:15px;fill:none;stroke-width:1.8;}
.feat.blue .feat-icon svg{stroke:#E6F1FB;}
.feat.teal .feat-icon svg{stroke:#E1F5EE;}
.feat.purple .feat-icon svg{stroke:#EEEDFE;}
.feat.coral .feat-icon svg{stroke:#FAECE7;}
.feat-name{font-size:12px;font-weight:600;color:#E6F1FB;margin-bottom:2px;}
.feat.blue .feat-hint{color:#85B7EB;}
.feat.teal .feat-hint{color:#9FE1CB;}
.feat.purple .feat-hint{color:#AFA9EC;}
.feat.coral .feat-hint{color:#F5C4B3;}
.feat-hint{font-size:11px;}
.drop-zone{background:#042C53;border:1.5px dashed #378ADD;border-radius:12px;padding:20px;text-align:center;cursor:pointer;transition:border-color .15s;}
.drop-zone.drag,.drop-zone:hover{border-color:#85B7EB;}
.drop-zone input{display:none;}
.drop-icon{width:38px;height:38px;background:#0C447C;border-radius:50%;margin:0 auto 10px;display:flex;align-items:center;justify-content:center;}
.drop-icon svg{width:17px;height:17px;fill:none;stroke:#378ADD;stroke-width:1.8;}
.drop-title{font-size:13px;font-weight:600;color:#B5D4F4;margin-bottom:3px;}
.drop-hint{font-size:11px;color:#378ADD;}
.fname{font-size:12px;color:#9FE1CB;margin-top:8px;min-height:16px;}
.pills{display:flex;justify-content:center;gap:6px;margin-top:10px;}
.pill{background:#0C447C;color:#85B7EB;font-size:11px;padding:3px 9px;border-radius:20px;}
.footer{display:flex;justify-content:space-between;align-items:center;margin-top:14px;}
.submit-btn{background:#378ADD;color:#042C53;font-size:14px;font-weight:700;padding:10px 24px;
            border:none;border-radius:8px;cursor:pointer;transition:background .15s;}
.submit-btn:hover:not(:disabled){background:#85B7EB;}
.submit-btn:disabled{background:#185FA5;color:#378ADD;cursor:not-allowed;}
</style>
</head>
<body>
<div class="page">
  <div class="topbar">
    <div class="logo">
      <div class="logo-icon">
        <svg viewBox="0 0 24 24"><path d="M12 2L2 7l10 5 10-5-10-5zM2 17l10 5 10-5M2 12l10 5 10-5"/></svg>
      </div>
      <div>
        <div class="logo-name">Payments</div>
        <div class="logo-sub">ניתוח הוצאות</div>
      </div>
    </div>
    <div class="ver">v2.0</div>
  </div>

  <div class="headline">מה תרצה לנתח?</div>
  <div class="subline">בחר סוג ניתוח, גרור קובץ והתוצאות יוצגו מיד</div>

  __ERROR__

  <div class="cards">
    <div class="feat blue active" data-mode="credit" onclick="setMode(this,'credit','/upload','file','.xlsx,.json,.pdf')">
      <div class="feat-icon">
        <svg viewBox="0 0 24 24"><rect x="2" y="5" width="20" height="14" rx="2"/><path d="M2 10h20"/></svg>
      </div>
      <div class="feat-name">כרטיס אשראי</div>
      <div class="feat-hint">Cal · Isracard · xlsx · pdf</div>
    </div>
    <div class="feat teal" data-mode="bank" onclick="setMode(this,'bank','/bank','bank-file','.xls,.xlsx,.pdf')">
      <div class="feat-icon">
        <svg viewBox="0 0 24 24"><path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg>
      </div>
      <div class="feat-name">תנועות בנק</div>
      <div class="feat-hint">Fibi · xls · pdf</div>
    </div>
    <div class="feat purple" data-mode="multi" onclick="setMode(this,'multi',null,null,null)">
      <div class="feat-icon">
        <svg viewBox="0 0 24 24"><path d="M18 20V10M12 20V4M6 20v-6"/></svg>
      </div>
      <div class="feat-name">השוואת חודשים</div>
      <div class="feat-hint">עד 12 חודשים</div>
    </div>
    <div class="feat coral" data-mode="compare" onclick="setMode(this,'compare','/compare','cmp-file','.xlsx,.json,.pdf')">
      <div class="feat-icon">
        <svg viewBox="0 0 24 24"><path d="M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01"/></svg>
      </div>
      <div class="feat-name">השוואת 2 חודשים</div>
      <div class="feat-hint">לפני / אחרי</div>
    </div>
  </div>

  <form id="main-form" action="/upload" method="POST" enctype="multipart/form-data">
    <label class="drop-zone" id="drop" for="file-input">
      <input type="file" id="file-input" name="file" accept=".xlsx,.json,.pdf" required>
      <div class="drop-icon">
        <svg viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M17 8l-5-5-5 5M12 3v12"/></svg>
      </div>
      <div class="drop-title">גרור קובץ לכאן</div>
      <div class="drop-hint">או לחץ לבחירה מהמחשב</div>
      <div class="fname" id="fname"></div>
      <div class="pills" id="pills">
        <span class="pill">xlsx</span><span class="pill">pdf</span>
        <span class="pill">json</span>
      </div>
    </label>

    <div class="footer">
      <span style="font-size:12px;color:#85B7EB;" id="mode-label">כרטיס אשראי — פירוט חודשי</span>
      <button type="submit" id="submit" disabled class="submit-btn">העלה קובץ</button>
    </div>
  </form>
</div>

<script>
const modes = {
  credit:  {action:'/upload', accept:'.xlsx,.json,.pdf', label:'כרטיס אשראי — פירוט חודשי',  pills:['xlsx','pdf','json']},
  bank:    {action:'/bank',   accept:'.xls,.xlsx,.pdf',  label:'תנועות בנק — הכנסות vs הוצאות', pills:['xls','pdf']},
  multi:   {action:'/multi',  accept:'.xlsx,.json,.pdf', label:'השוואת עד 12 חודשים',         pills:['xlsx','pdf','json'], multi:true},
  compare: {action:'/compare',accept:'.xlsx,.json,.pdf', label:'השוואת 2 חודשים',              pills:['xlsx','pdf','json']},
};
let curMode = 'credit';
const form = document.getElementById('main-form');
const fileInput = document.getElementById('file-input');
const drop = document.getElementById('drop');
const fname = document.getElementById('fname');
const submit = document.getElementById('submit');
const pillsEl = document.getElementById('pills');
const modeLabel = document.getElementById('mode-label');

function setMode(el, mode) {
  document.querySelectorAll('.feat').forEach(f => f.classList.remove('active'));
  el.classList.add('active');
  curMode = mode;
  const m = modes[mode];
  form.action = m.action;
  fileInput.accept = m.accept;
  fileInput.value = '';
  fname.textContent = '';
  modeLabel.textContent = m.label;
  pillsEl.innerHTML = m.pills.map(p => `<span class="pill">${p}</span>`).join('');
  if (mode === 'multi') {
    // Multi has its own multi-file uploader page
    drop.style.display = 'none';
    submit.disabled = false;
    submit.textContent = 'עבור לדף ההשוואה ←';
  } else {
    drop.style.display = '';
    submit.textContent = 'העלה קובץ';
    submit.disabled = !fileInput.files.length;
  }
}

function update() {
  if (fileInput.files.length) {
    fname.textContent = fileInput.files[0].name;
    submit.disabled = false;
  }
}
fileInput.addEventListener('change', update);
drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('drag'); });
drop.addEventListener('dragleave', () => drop.classList.remove('drag'));
drop.addEventListener('drop', e => {
  e.preventDefault(); drop.classList.remove('drag');
  if (e.dataTransfer.files.length) { fileInput.files = e.dataTransfer.files; update(); }
});
form.addEventListener('submit', e => {
  if (curMode === 'multi') { e.preventDefault(); window.location.href = '/multi'; return; }
  submit.disabled = true; submit.textContent = 'מעבד...';
});
</script>
</body>
</html>
"""


COMPARE_FORM = """<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
<meta charset="UTF-8">
<title>Payments — השוואת שני חודשים</title>
<style>
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:-apple-system,"Segoe UI",Arial,sans-serif;background:#042C53;
     display:flex;align-items:center;justify-content:center;min-height:100vh;padding:20px;}
.page{background:#0C447C;border-radius:16px;padding:28px 24px;max-width:500px;width:100%;overflow:hidden;}
.topbar{display:flex;align-items:center;justify-content:space-between;margin-bottom:20px;}
.back-btn{color:#85B7EB;font-size:13px;text-decoration:none;}
.mode-icon{width:32px;height:32px;background:#D85A30;border-radius:8px;display:flex;align-items:center;justify-content:center;}
.mode-icon svg{width:16px;height:16px;fill:none;stroke:#FAECE7;stroke-width:2;}
.ttl{font-size:18px;font-weight:600;color:#E6F1FB;margin-bottom:4px;}
.sub{font-size:13px;color:#85B7EB;margin-bottom:16px;}
.err{background:#4A1B0C;color:#F5C4B3;padding:10px 14px;border-radius:8px;font-size:13px;margin-bottom:14px;}
.row{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:14px;}
.slot-lbl{font-size:11px;font-weight:600;color:#85B7EB;margin-bottom:6px;letter-spacing:.5px;}
.drop{display:block;background:#042C53;border:1.5px dashed #378ADD;border-radius:10px;
      padding:20px 12px;text-align:center;cursor:pointer;transition:border-color .15s;}
.drop:hover,.drop.drag{border-color:#85B7EB;}
.drop input{display:none;}
.di{width:32px;height:32px;background:#0C447C;border-radius:50%;margin:0 auto 8px;display:flex;align-items:center;justify-content:center;}
.di svg{width:14px;height:14px;fill:none;stroke:#378ADD;stroke-width:1.8;}
.dl{font-size:12px;color:#B5D4F4;font-weight:500;margin-bottom:2px;}
.ds{font-size:11px;color:#378ADD;}
.fname{font-size:11px;color:#9FE1CB;margin-top:6px;min-height:14px;text-align:center;}
.footer{display:flex;justify-content:space-between;align-items:center;}
.sbtn{background:#378ADD;color:#042C53;font-size:14px;font-weight:700;padding:10px 24px;
      border:none;border-radius:8px;cursor:pointer;}
.sbtn:hover:not(:disabled){background:#85B7EB;}
.sbtn:disabled{background:#185FA5;color:#378ADD;cursor:not-allowed;}
</style>
</head>
<body>
<div class="page">
  <div class="topbar">
    <a class="back-btn" href="/">← חזרה</a>
    <div class="mode-icon"><svg viewBox="0 0 24 24"><path d="M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01"/></svg></div>
  </div>
  <div class="ttl">השוואת שני חודשים</div>
  <div class="sub">העלו שני קבצים — לפני ואחרי — ותקבלו השוואה מלאה</div>
  __ERROR__
  <form id="form" action="/compare" method="POST" enctype="multipart/form-data">
    <div class="row">
      <div>
        <div class="slot-lbl">חודש א׳</div>
        <label class="drop" id="drop-a">
          <input type="file" name="file_a" id="file_a" accept=".xlsx,.json,.pdf" required>
          <div class="di"><svg viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M17 8l-5-5-5 5M12 3v12"/></svg></div>
          <div class="dl">בחר קובץ</div><div class="ds">xlsx · pdf · json</div>
        </label>
        <div class="fname" id="fname_a"></div>
      </div>
      <div>
        <div class="slot-lbl">חודש ב׳</div>
        <label class="drop" id="drop-b">
          <input type="file" name="file_b" id="file_b" accept=".xlsx,.json,.pdf" required>
          <div class="di"><svg viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M17 8l-5-5-5 5M12 3v12"/></svg></div>
          <div class="dl">בחר קובץ</div><div class="ds">xlsx · pdf · json</div>
        </label>
        <div class="fname" id="fname_b"></div>
      </div>
    </div>
    <div class="footer">
      <span style="font-size:12px;color:#85B7EB;">השוואת 2 חודשים</span>
      <button type="submit" id="submit" disabled class="sbtn">השווה</button>
    </div>
  </form>
</div>
<script>
  const fA=document.getElementById('file_a'),fB=document.getElementById('file_b'),sb=document.getElementById('submit');
  function upd(){document.getElementById('fname_a').textContent=fA.files[0]?.name||'';document.getElementById('fname_b').textContent=fB.files[0]?.name||'';sb.disabled=!(fA.files.length&&fB.files.length);}
  [fA,fB].forEach(f=>f.addEventListener('change',upd));
  [['drop-a','file_a'],['drop-b','file_b']].forEach(([dId,fId])=>{
    const d=document.getElementById(dId),inp=document.getElementById(fId);
    d.addEventListener('dragover',e=>{e.preventDefault();d.classList.add('drag');});
    d.addEventListener('dragleave',()=>d.classList.remove('drag'));
    d.addEventListener('drop',e=>{e.preventDefault();d.classList.remove('drag');if(e.dataTransfer.files.length){inp.files=e.dataTransfer.files;upd();}});
  });
  document.getElementById('form').addEventListener('submit',()=>{sb.disabled=true;sb.textContent='מעבד...';});
</script>
</body>
</html>
"""

def render_form(error: str | None = None) -> str:
    err_html = f'<div class="err">{error}</div>' if error else ""
    return UPLOAD_FORM.replace("__ERROR__", err_html)


def render_compare_form(error: str | None = None) -> str:
    err_html = f'<div class="err">{error}</div>' if error else ""
    return COMPARE_FORM.replace("__ERROR__", err_html)


def _parse_upload(f) -> dict:
    """Save an uploaded file to a temp path, parse it, and clean up.
    Accepts .xlsx (Excel), .json (saved via Save JSON), or .pdf (Cal digital statement).
    """
    fname = f.filename.lower()
    if fname.endswith(".json"):
        data = json.loads(f.read().decode("utf-8"))
        if "payments" not in data:
            raise ValueError("JSON חסר מפתח 'payments'")
        data.setdefault("title", Path(f.filename).stem)
        data.setdefault("source", f.filename)
        data.setdefault("issuer", "")
        from payments_ui import _normalize_merchant
        for p in data["payments"]:
            if "canonical" not in p:
                p["canonical"] = _normalize_merchant(p.get("merchant", ""))
        return data

    suffix = ".pdf" if fname.endswith(".pdf") else ".xlsx"
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        f.save(tmp.name)
        tmp_path = Path(tmp.name)
    try:
        data = parse_payments(tmp_path)
        data["title"] = Path(f.filename).stem.replace("_", " ")
        data["source"] = f.filename
        return data
    finally:
        tmp_path.unlink(missing_ok=True)


@app.get("/health")
def health():
    import importlib
    pkgs = ["flask", "openpyxl", "pdfplumber", "pdfminer"]
    status = {}
    for pkg in pkgs:
        try:
            m = importlib.import_module(pkg.replace(".", "_").split("_")[0])
            status[pkg] = getattr(m, "__version__", "ok")
        except ImportError as e:
            status[pkg] = f"MISSING: {e}"
    return status


@app.get("/")
def index():
    return render_form()


@app.get("/compare")
def compare_form():
    return render_compare_form()


@app.post("/compare")
def compare():
    fa = request.files.get("file_a")
    fb = request.files.get("file_b")
    if not fa or not fa.filename or not fb or not fb.filename:
        return render_compare_form("יש לבחור שני קבצים."), 400
    for f in (fa, fb):
        if not f.filename.lower().endswith((".xlsx", ".json", ".pdf")):
            return render_compare_form("יש להעלות קבצי .xlsx, .json או .pdf בלבד."), 400

    try:
        data_a = _parse_upload(fa)
        data_b = _parse_upload(fb)
        html = generate_comparison_html(data_a, data_b)
    except Exception as e:
        return render_compare_form(f"כשל בקריאת הקבצים: {e}"), 400

    return html


MULTI_FORM = """<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
<meta charset="UTF-8">
<title>Payments — השוואת חודשים מרובים</title>
<style>
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:-apple-system,"Segoe UI",Arial,sans-serif;background:#042C53;
     display:flex;align-items:center;justify-content:center;min-height:100vh;padding:20px;}
.page{background:#0C447C;border-radius:16px;padding:28px 24px;max-width:500px;width:100%;overflow:hidden;}
.topbar{display:flex;align-items:center;justify-content:space-between;margin-bottom:20px;}
.back-btn{color:#85B7EB;font-size:13px;text-decoration:none;}
.mode-icon{width:32px;height:32px;background:#534AB7;border-radius:8px;display:flex;align-items:center;justify-content:center;}
.mode-icon svg{width:16px;height:16px;fill:none;stroke:#EEEDFE;stroke-width:2;}
.ttl{font-size:18px;font-weight:600;color:#E6F1FB;margin-bottom:4px;}
.sub{font-size:13px;color:#85B7EB;margin-bottom:16px;}
.err{background:#4A1B0C;color:#F5C4B3;padding:10px 14px;border-radius:8px;font-size:13px;margin-bottom:14px;display:none;}
.drop-zone{background:#042C53;border:1.5px dashed #378ADD;border-radius:12px;
           padding:24px;text-align:center;cursor:pointer;transition:border-color .15s;margin-bottom:12px;}
.drop-zone:hover,.drop-zone.drag{border-color:#85B7EB;}
.drop-zone input{display:none;}
.di{width:38px;height:38px;background:#0C447C;border-radius:50%;margin:0 auto 10px;display:flex;align-items:center;justify-content:center;}
.di svg{width:17px;height:17px;fill:none;stroke:#378ADD;stroke-width:1.8;}
.dl{font-size:13px;font-weight:600;color:#B5D4F4;margin-bottom:3px;}
.ds{font-size:11px;color:#378ADD;}
.file-list{display:flex;flex-direction:column;gap:6px;margin-bottom:8px;}
.file-item{display:flex;align-items:center;justify-content:space-between;
           background:#042C53;border-radius:8px;padding:8px 12px;font-size:12px;color:#B5D4F4;}
.file-item button{background:none;border:none;color:#85B7EB;cursor:pointer;font-size:14px;padding:0 2px;}
.file-item button:hover{color:#F5C4B3;}
.counter{font-size:11px;color:#85B7EB;text-align:center;margin-bottom:12px;min-height:14px;}
.footer{display:flex;justify-content:space-between;align-items:center;}
.sbtn{background:#7F77DD;color:#EEEDFE;font-size:14px;font-weight:700;padding:10px 24px;
      border:none;border-radius:8px;cursor:pointer;}
.sbtn:hover:not(:disabled){background:#AFA9EC;}
.sbtn:disabled{background:#3C3489;color:#7F77DD;cursor:not-allowed;}
</style>
</head>
<body>
<div class="page">
  <div class="topbar">
    <a class="back-btn" href="/">← חזרה</a>
    <div class="mode-icon"><svg viewBox="0 0 24 24"><path d="M18 20V10M12 20V4M6 20v-6"/></svg></div>
  </div>
  <div class="ttl">השוואת חודשים מרובים</div>
  <div class="sub">גרור עד 12 קבצים — ותקבל השוואה מלאה בין כל החודשים</div>
  __ERROR__
  <div id="drop" class="drop-zone">
    <input type="file" name="files" id="files" accept=".xlsx,.json,.pdf" multiple>
    <div class="di"><svg viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M17 8l-5-5-5 5M12 3v12"/></svg></div>
    <div class="dl">גרור קבצים לכאן</div>
    <div class="ds">xlsx · pdf · json · עד 12 קבצים</div>
  </div>
  <div class="file-list" id="file-list"></div>
  <div class="counter" id="counter"></div>
  <div class="err" id="err-msg"></div>
  <div class="footer">
    <span style="font-size:12px;color:#85B7EB;" id="mode-hint">0 קבצים נבחרו</span>
    <button type="button" id="submit" disabled class="sbtn" onclick="submitFiles()">יצירת השוואה</button>
  </div>
</div>
<script>
  const MAX=12;
  let selected=new DataTransfer();
  const dropEl=document.getElementById('drop'),filesEl=document.getElementById('files'),
        listEl=document.getElementById('file-list'),counterEl=document.getElementById('counter'),
        submitEl=document.getElementById('submit'),errEl=document.getElementById('err-msg'),
        hintEl=document.getElementById('mode-hint');

  function refreshUI(){
    const files=[...selected.files];
    listEl.innerHTML=files.map((f,i)=>`<div class="file-item"><span>${f.name}</span><button type="button" data-i="${i}">✕</button></div>`).join('');
    const cnt=files.length;
    counterEl.textContent=cnt?`${cnt} / ${MAX} קבצים`:'';
    hintEl.textContent=`${cnt} קבצים נבחרו`;
    submitEl.disabled=cnt<2;
    listEl.querySelectorAll('button[data-i]').forEach(btn=>{
      btn.addEventListener('click',()=>{
        const idx=parseInt(btn.dataset.i),next=new DataTransfer();
        [...selected.files].forEach((f,i)=>{if(i!==idx)next.items.add(f);});
        selected=next;refreshUI();
      });
    });
  }
  function addFiles(newFiles){
    for(const f of newFiles){
      if(selected.files.length>=MAX)break;
      const ext=f.name.toLowerCase();
      if(!ext.endsWith('.xlsx')&&!ext.endsWith('.json')&&!ext.endsWith('.pdf'))continue;
      if([...selected.files].some(e=>e.name===f.name))continue;
      selected.items.add(f);
    }
    refreshUI();
  }
  async function submitFiles(){
    if(selected.files.length<2)return;
    submitEl.disabled=true;submitEl.textContent='מעבד...';errEl.style.display='none';
    const fd=new FormData();
    for(const f of selected.files)fd.append('files',f);
    try{
      const resp=await fetch('/multi',{method:'POST',body:fd,redirect:'follow'});
      if(resp.ok){window.location.href=resp.url;}
      else{
        const text=await resp.text();
        const m=text.match(/class="err">([\\s\\S]*?)<\/div>/);
        errEl.textContent=m?m[1]:`שגיאה ${resp.status}`;errEl.style.display='block';
        submitEl.disabled=false;submitEl.textContent='יצירת השוואה';
      }
    }catch(e){
      errEl.textContent='שגיאת רשת: '+e.message;errEl.style.display='block';
      submitEl.disabled=false;submitEl.textContent='יצירת השוואה';
    }
  }
  filesEl.addEventListener('change',()=>{addFiles(filesEl.files);filesEl.value='';});
  dropEl.addEventListener('click',()=>filesEl.click());
  dropEl.addEventListener('dragover',e=>{e.preventDefault();dropEl.classList.add('drag');});
  dropEl.addEventListener('dragleave',()=>dropEl.classList.remove('drag'));
  dropEl.addEventListener('drop',e=>{e.preventDefault();dropEl.classList.remove('drag');addFiles(e.dataTransfer.files);});
</script>
</body>
</html>
"""

def render_multi_form(error: str | None = None) -> str:
    err_html = f'<div class="err">{error}</div>' if error else ""
    return MULTI_FORM.replace("__ERROR__", err_html)


@app.get("/multi")
def multi_form():
    return render_multi_form()


@app.post("/multi")
def multi():
    files = request.files.getlist("files")
    files = [f for f in files if f and f.filename]
    if len(files) < 2:
        return render_multi_form("יש לבחור לפחות 2 קבצים."), 400
    if len(files) > 12:
        return render_multi_form("ניתן להעלות עד 12 קבצים בלבד."), 400
    for f in files:
        if not f.filename.lower().endswith((".xlsx", ".json", ".pdf")):
            return render_multi_form("יש להעלות קבצי .xlsx, .json או .pdf בלבד."), 400
    try:
        months_data = [_parse_upload(f) for f in files]

        # Cache individual month dashboards so they can be opened in a new tab
        month_urls = []
        for d in months_data:
            mid = str(uuid.uuid4())
            _cache_write(mid, generate_html(d))
            month_urls.append(f"/multi/result/{mid}")

        html = generate_multi_html(months_data, month_urls=month_urls)
        result_id = str(uuid.uuid4())
        _cache_write(result_id, html)
        _cache_cleanup()
        return redirect(f"/multi/result/{result_id}")
    except Exception as e:
        return render_multi_form(f"כשל בקריאת הקבצים: {e}"), 400


@app.get("/multi/result/<result_id>")
def multi_result(result_id: str):
    html = _cache_read(result_id)
    if not html:
        return redirect("/multi")
    return html


@app.post("/upload")
def upload():
    f = request.files.get("file")
    if not f or not f.filename:
        return render_form("לא נבחר קובץ."), 400
    if not f.filename.lower().endswith((".xlsx", ".json", ".pdf")):
        return render_form("יש להעלות קובץ .xlsx, .json או .pdf בלבד."), 400

    try:
        data = _parse_upload(f)
        return generate_html(data)
    except Exception as e:
        return render_form(f"כשל בקריאת הקובץ: {e}"), 400


BANK_FORM = """<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
<meta charset="UTF-8">
<title>ניתוח תנועות בנק</title>
<style>
  *{box-sizing:border-box;}
  body{font-family:-apple-system,"Segoe UI",Arial,sans-serif;background:#f5f5f7;
       color:#222;display:flex;align-items:center;justify-content:center;
       min-height:100vh;margin:0;padding:20px;}
  .card{background:#fff;padding:32px 36px;border-radius:12px;
        box-shadow:0 2px 16px rgba(0,0,0,.08);max-width:460px;width:100%;}
  h1{font-size:20px;margin:0 0 6px;}
  p{color:#666;font-size:14px;margin:0 0 22px;}
  .drop{display:block;border:2px dashed #ccc;border-radius:8px;
        padding:36px 16px;text-align:center;cursor:pointer;
        transition:all .15s;background:#fafafa;}
  .drop:hover,.drop.drag{border-color:#2196f3;background:#e3f2fd;}
  .drop input{display:none;}
  .drop .icon{font-size:36px;line-height:1;margin-bottom:10px;}
  .drop .lbl{font-size:14px;color:#444;}
  .drop .sub{font-size:12px;color:#888;margin-top:4px;}
  .fname{margin-top:10px;font-size:13px;color:#1976d2;text-align:center;min-height:18px;}
  button{margin-top:14px;width:100%;padding:12px;background:#2196f3;color:#fff;
         border:0;border-radius:6px;font-size:15px;font-weight:600;cursor:pointer;}
  button:hover:not(:disabled){background:#1976d2;}
  button:disabled{background:#ccc;cursor:not-allowed;}
  .err{background:#ffebee;color:#c62828;padding:10px 12px;border-radius:6px;
       font-size:13px;margin-bottom:14px;}
  .back{display:block;text-align:center;margin-top:16px;font-size:13px;color:#2196f3;text-decoration:none;}
  .info{background:#e3f2fd;color:#1565c0;padding:10px 12px;border-radius:6px;
        font-size:13px;margin-bottom:14px;line-height:1.6;}
</style>
</head>
<body>
<div class="card">
  <h1>📊 ניתוח תנועות בנק</h1>
  <p>העלו קובץ תנועות מהבנק הבינלאומי — ניתוח הכנסות מול הוצאות, מגמות ותובנות.</p>
  <div class="info">
    ✓ קובץ Excel (.xls) — ייצוא מאתר FibiSave<br>
    ✓ קובץ PDF — דף פירוט מהבנק
  </div>
  __ERROR__
  <form id="form" action="/bank" method="POST" enctype="multipart/form-data">
    <label class="drop" id="drop">
      <div class="icon">🏦</div>
      <div class="lbl">גררו קובץ או לחצו לבחירה</div>
      <div class="sub">.xls / .pdf · עד 16MB</div>
      <input type="file" name="file" id="file" accept=".xls,.xlsx,.pdf" required>
    </label>
    <div class="fname" id="fname"></div>
    <button type="submit" id="submit" disabled>ניתוח תנועות</button>
  </form>
  <a class="back" href="/">← חזרה לדף הראשי</a>
</div>
<script>
  const drop=document.getElementById('drop'),file=document.getElementById('file'),
        fname=document.getElementById('fname'),submit=document.getElementById('submit'),
        form=document.getElementById('form');
  function update(){if(file.files.length){fname.textContent=file.files[0].name;submit.disabled=false;}}
  file.addEventListener('change',update);
  drop.addEventListener('dragover',e=>{e.preventDefault();drop.classList.add('drag');});
  drop.addEventListener('dragleave',()=>drop.classList.remove('drag'));
  drop.addEventListener('drop',e=>{e.preventDefault();drop.classList.remove('drag');
    if(e.dataTransfer.files.length){file.files=e.dataTransfer.files;update();}});
  form.addEventListener('submit',()=>{submit.disabled=true;submit.textContent='מנתח...';});
</script>
</body>
</html>
"""


def render_bank_form(error: str | None = None) -> str:
    err_html = f'<div class="err">{error}</div>' if error else ""
    return BANK_FORM.replace("__ERROR__", err_html)


@app.get("/bank")
def bank_form():
    return render_bank_form()


@app.post("/bank")
def bank_upload():
    f = request.files.get("file")
    if not f or not f.filename:
        return render_bank_form("לא נבחר קובץ."), 400
    if not f.filename.lower().endswith((".xls", ".xlsx", ".pdf")):
        return render_bank_form("יש להעלות קובץ .xls, .xlsx או .pdf בלבד."), 400
    try:
        suffix = Path(f.filename).suffix.lower()
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
            f.save(tmp.name)
            tmp_path = Path(tmp.name)
        try:
            data = parse_bank_statement(tmp_path)
            data["source"] = f.filename
            data["title"] = Path(f.filename).stem
        finally:
            tmp_path.unlink(missing_ok=True)
        result_id = str(uuid.uuid4())
        _cache_write(result_id, generate_bank_html(data))
        _cache_cleanup()
        return redirect(f"/bank/result/{result_id}")
    except Exception as e:
        return render_bank_form(f"כשל בקריאת הקובץ: {e}"), 400


@app.get("/bank/result/<result_id>")
def bank_result(result_id: str):
    html = _cache_read(result_id)
    if not html:
        return redirect("/bank")
    return html


def main() -> int:
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    port = int(args[0]) if args else 5000
    url = f"http://127.0.0.1:{port}"
    print(f"Serving on {url}  (Ctrl-C to stop)")

    if "--open" in sys.argv:
        import threading
        import webbrowser
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()

    app.run(host="0.0.0.0", port=port, debug=False)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
