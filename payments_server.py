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
import sys
import tempfile
import uuid
from collections import OrderedDict
from pathlib import Path

from flask import Flask, request, redirect

from payments_ui import parse_payments, generate_html, generate_comparison_html, generate_multi_html

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB

# In-memory store for multi-month results (max 20 entries)
_RESULT_CACHE: OrderedDict[str, str] = OrderedDict()

UPLOAD_FORM = """<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
<meta charset="UTF-8">
<title>Payments UI — העלאת קובץ</title>
<style>
  * { box-sizing: border-box; }
  body { font-family: -apple-system, "Segoe UI", Arial, sans-serif; background: #f5f5f7;
         color: #222; display: flex; align-items: center; justify-content: center;
         min-height: 100vh; margin: 0; padding: 20px; }
  .card { background: #fff; padding: 32px 36px; border-radius: 12px;
          box-shadow: 0 2px 16px rgba(0,0,0,0.08); max-width: 460px; width: 100%; }
  h1 { font-size: 20px; margin: 0 0 6px; }
  p  { color: #666; font-size: 14px; margin: 0 0 22px; }
  .drop { display: block; border: 2px dashed #ccc; border-radius: 8px;
          padding: 36px 16px; text-align: center; cursor: pointer;
          transition: all 0.15s; background: #fafafa; }
  .drop:hover, .drop.drag { border-color: #2196f3; background: #e3f2fd; }
  .drop input { display: none; }
  .drop .icon  { font-size: 36px; line-height: 1; margin-bottom: 10px; }
  .drop .label { font-size: 14px; color: #444; }
  .drop .sub   { font-size: 12px; color: #888; margin-top: 4px; }
  .filename { margin-top: 10px; font-size: 13px; color: #1976d2; text-align: center; min-height: 18px; }
  button { margin-top: 14px; width: 100%; padding: 12px; background: #2196f3; color: #fff;
           border: 0; border-radius: 6px; font-size: 15px; font-weight: 600; cursor: pointer; }
  button:hover:not(:disabled) { background: #1976d2; }
  button:disabled { background: #ccc; cursor: not-allowed; }
  .err { background: #ffebee; color: #c62828; padding: 10px 12px; border-radius: 6px;
         font-size: 13px; margin-bottom: 14px; }
</style>
</head>
<body>
  <div class="card">
    <h1>העלאת קובץ עסקאות</h1>
    <p>בחרו קובץ Excel (.xlsx), PDF של פירוט כאל, או JSON שנשמר קודם — כדי לייצר תצוגה אינטראקטיבית.</p>
    __ERROR__
    <form id="form" action="/upload" method="POST" enctype="multipart/form-data">
      <label class="drop" id="drop">
        <div class="icon">📄</div>
        <div class="label">גררו קובץ או לחצו לבחירה</div>
        <div class="sub">.xlsx / .json / .pdf · עד 16MB</div>
        <input type="file" name="file" id="file" accept=".xlsx,.json,.pdf" required>
      </label>
      <div class="filename" id="filename"></div>
      <button type="submit" id="submit" disabled>יצירת תצוגה</button>
    </form>
    <div style="text-align:center;margin-top:16px;font-size:13px;display:flex;flex-direction:column;gap:8px;">
      <a href="/compare" style="color:#2196f3;text-decoration:none;">השוואה בין שני חודשים →</a>
      <a href="/multi" style="color:#2196f3;text-decoration:none;">השוואת עד 12 חודשים →</a>
    </div>
  </div>
<script>
  const drop = document.getElementById('drop');
  const file = document.getElementById('file');
  const name = document.getElementById('filename');
  const submit = document.getElementById('submit');
  const form = document.getElementById('form');

  function update() {
    if (file.files.length) {
      name.textContent = file.files[0].name;
      submit.disabled = false;
    }
  }
  file.addEventListener('change', update);
  drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('drag'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('drag'));
  drop.addEventListener('drop', e => {
    e.preventDefault();
    drop.classList.remove('drag');
    if (e.dataTransfer.files.length) {
      file.files = e.dataTransfer.files;
      update();
    }
  });
  form.addEventListener('submit', () => {
    submit.disabled = true;
    submit.textContent = 'מעבד...';
  });
</script>
</body>
</html>
"""


COMPARE_FORM = """<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
<meta charset="UTF-8">
<title>Payments UI — השוואת חודשים</title>
<style>
  * { box-sizing: border-box; }
  body { font-family: -apple-system, "Segoe UI", Arial, sans-serif; background: #f5f5f7;
         color: #222; display: flex; align-items: center; justify-content: center;
         min-height: 100vh; margin: 0; padding: 20px; }
  .card { background: #fff; padding: 32px 36px; border-radius: 12px;
          box-shadow: 0 2px 16px rgba(0,0,0,0.08); max-width: 520px; width: 100%; }
  h1 { font-size: 20px; margin: 0 0 6px; }
  p  { color: #666; font-size: 14px; margin: 0 0 22px; }
  .row { display: flex; gap: 12px; }
  .slot { flex: 1; }
  .slot label { display: block; font-size: 12px; color: #555; margin-bottom: 4px; font-weight: 600; }
  .drop { display: block; border: 2px dashed #ccc; border-radius: 8px;
          padding: 28px 12px; text-align: center; cursor: pointer;
          transition: all 0.15s; background: #fafafa; }
  .drop:hover, .drop.drag { border-color: #2196f3; background: #e3f2fd; }
  .drop input { display: none; }
  .drop .icon  { font-size: 28px; line-height: 1; margin-bottom: 6px; }
  .drop .label { font-size: 13px; color: #444; }
  .drop .sub   { font-size: 11px; color: #888; margin-top: 2px; }
  .fname { margin-top: 6px; font-size: 12px; color: #1976d2; text-align: center; min-height: 16px; }
  button { margin-top: 18px; width: 100%; padding: 12px; background: #2196f3; color: #fff;
           border: 0; border-radius: 6px; font-size: 15px; font-weight: 600; cursor: pointer; }
  button:hover:not(:disabled) { background: #1976d2; }
  button:disabled { background: #ccc; cursor: not-allowed; }
  .err { background: #ffebee; color: #c62828; padding: 10px 12px; border-radius: 6px;
         font-size: 13px; margin-bottom: 14px; }
  .back { display: block; text-align: center; margin-top: 16px; font-size: 13px; color: #2196f3; text-decoration: none; }
</style>
</head>
<body>
  <div class="card">
    <h1>השוואת שני חודשים</h1>
    <p>העלו שני קבצי .xlsx, .json או .pdf (למשל ינואר ופברואר) כדי לראות שינויים לפי ענף ובית עסק.</p>
    __ERROR__
    <form id="form" action="/compare" method="POST" enctype="multipart/form-data">
      <div class="row">
        <div class="slot">
          <label>חודש א'</label>
          <label class="drop" data-for="file_a">
            <div class="icon">📄</div>
            <div class="label">בחרו קובץ</div>
            <div class="sub">.xlsx / .json</div>
            <input type="file" name="file_a" id="file_a" accept=".xlsx,.json,.pdf" required>
          </label>
          <div class="fname" id="fname_a"></div>
        </div>
        <div class="slot">
          <label>חודש ב'</label>
          <label class="drop" data-for="file_b">
            <div class="icon">📄</div>
            <div class="label">בחרו קובץ</div>
            <div class="sub">.xlsx / .json</div>
            <input type="file" name="file_b" id="file_b" accept=".xlsx,.json,.pdf" required>
          </label>
          <div class="fname" id="fname_b"></div>
        </div>
      </div>
      <button type="submit" id="submit" disabled>השוואה</button>
    </form>
    <a class="back" href="/">← חזרה לקובץ בודד</a>
  </div>
<script>
  const fileA = document.getElementById('file_a');
  const fileB = document.getElementById('file_b');
  const submit = document.getElementById('submit');

  function update() {
    document.getElementById('fname_a').textContent = fileA.files[0]?.name || '';
    document.getElementById('fname_b').textContent = fileB.files[0]?.name || '';
    submit.disabled = !(fileA.files.length && fileB.files.length);
  }
  [fileA, fileB].forEach(f => f.addEventListener('change', update));

  document.querySelectorAll('.drop').forEach(d => {
    const input = document.getElementById(d.dataset.for);
    d.addEventListener('dragover', e => { e.preventDefault(); d.classList.add('drag'); });
    d.addEventListener('dragleave', () => d.classList.remove('drag'));
    d.addEventListener('drop', e => {
      e.preventDefault();
      d.classList.remove('drag');
      if (e.dataTransfer.files.length) { input.files = e.dataTransfer.files; update(); }
    });
  });

  document.getElementById('form').addEventListener('submit', () => {
    submit.disabled = true;
    submit.textContent = 'מעבד...';
  });
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
  *{box-sizing:border-box;}
  body{font-family:-apple-system,"Segoe UI",Arial,sans-serif;background:#f5f5f7;
       color:#222;display:flex;align-items:center;justify-content:center;
       min-height:100vh;margin:0;padding:20px;}
  .card{background:#fff;padding:32px 36px;border-radius:12px;
        box-shadow:0 2px 16px rgba(0,0,0,.08);max-width:560px;width:100%;}
  h1{font-size:20px;margin:0 0 6px;}
  p{color:#666;font-size:14px;margin:0 0 20px;}
  .drop-zone{border:2px dashed #ccc;border-radius:8px;padding:32px 16px;
             text-align:center;cursor:pointer;transition:.15s;background:#fafafa;}
  .drop-zone:hover,.drop-zone.drag{border-color:#2196f3;background:#e3f2fd;}
  .drop-zone input{display:none;}
  .drop-zone .icon{font-size:36px;line-height:1;margin-bottom:10px;}
  .drop-zone .lbl{font-size:14px;color:#444;}
  .drop-zone .sub{font-size:12px;color:#888;margin-top:4px;}
  .file-list{margin-top:12px;display:flex;flex-direction:column;gap:6px;}
  .file-item{display:flex;align-items:center;justify-content:space-between;
             padding:6px 10px;background:#f0f4ff;border-radius:6px;font-size:13px;}
  .file-item button{background:none;border:none;color:#999;cursor:pointer;font-size:16px;padding:0 4px;}
  .file-item button:hover{color:#e53935;}
  .counter{font-size:12px;color:#888;margin-top:6px;text-align:center;}
  button[type=submit]{margin-top:16px;width:100%;padding:12px;background:#2196f3;
    color:#fff;border:0;border-radius:6px;font-size:15px;font-weight:600;cursor:pointer;}
  button[type=submit]:hover:not(:disabled){background:#1976d2;}
  button[type=submit]:disabled{background:#ccc;cursor:not-allowed;}
  .err{background:#ffebee;color:#c62828;padding:10px 12px;border-radius:6px;
       font-size:13px;margin-bottom:14px;}
  .back{display:block;text-align:center;margin-top:16px;font-size:13px;color:#2196f3;text-decoration:none;}
</style>
</head>
<body>
<div class="card">
  <h1>השוואת חודשים מרובים</h1>
  <p>העלו בין 2 ל-12 קבצי Excel (.xlsx), PDF כאל, או JSON כדי לראות השוואה מלאה בין החודשים.</p>
  __ERROR__
  <div id="drop" class="drop-zone">
    <div class="icon">📂</div>
    <div class="lbl">גררו קבצים או לחצו לבחירה</div>
    <div class="sub">.xlsx / .pdf / .json · עד 12 קבצים</div>
    <input type="file" name="files" id="files" accept=".xlsx,.json,.pdf" multiple>
  </div>
  <div class="file-list" id="file-list"></div>
  <div class="counter" id="counter"></div>
  <div class="err" id="err-msg" style="display:none;margin-top:12px;"></div>
  <button type="button" id="submit" disabled onclick="submitFiles()">יצירת השוואה</button>
  <a class="back" href="/">← חזרה</a>
</div>
<script>
  const MAX = 12;
  let selected = new DataTransfer();

  const dropEl   = document.getElementById('drop');
  const filesEl  = document.getElementById('files');
  const listEl   = document.getElementById('file-list');
  const counterEl= document.getElementById('counter');
  const submitEl = document.getElementById('submit');
  const errEl    = document.getElementById('err-msg');

  function refreshUI() {
    const files = [...selected.files];
    listEl.innerHTML = files.map((f, i) => `
      <div class="file-item">
        <span>📄 ${f.name}</span>
        <button type="button" data-i="${i}" title="הסר">✕</button>
      </div>`).join('');
    counterEl.textContent = files.length ? `${files.length} / ${MAX} קבצים נבחרו` : '';
    submitEl.disabled = files.length < 2;
    listEl.querySelectorAll('button[data-i]').forEach(btn => {
      btn.addEventListener('click', () => {
        const idx = parseInt(btn.dataset.i);
        const next = new DataTransfer();
        [...selected.files].forEach((f, i) => { if (i !== idx) next.items.add(f); });
        selected = next;
        refreshUI();
      });
    });
  }

  function addFiles(newFiles) {
    for (const f of newFiles) {
      if (selected.files.length >= MAX) break;
      const ext = f.name.toLowerCase();
      if (!ext.endsWith('.xlsx') && !ext.endsWith('.json') && !ext.endsWith('.pdf')) continue;
      if ([...selected.files].some(e => e.name === f.name)) continue;
      selected.items.add(f);
    }
    refreshUI();
  }

  async function submitFiles() {
    if (selected.files.length < 2) return;
    submitEl.disabled = true;
    submitEl.textContent = 'מעבד...';
    errEl.style.display = 'none';

    const fd = new FormData();
    for (const f of selected.files) fd.append('files', f);

    try {
      const resp = await fetch('/multi', { method: 'POST', body: fd, redirect: 'follow' });
      if (resp.ok) {
        window.location.href = resp.url;
      } else {
        const text = await resp.text();
        const m = text.match(/class="err">([\s\S]*?)<\/div>/);
        errEl.textContent = m ? m[1] : `שגיאה ${resp.status}`;
        errEl.style.display = 'block';
        submitEl.disabled = false;
        submitEl.textContent = 'יצירת השוואה';
      }
    } catch(e) {
      errEl.textContent = 'שגיאת רשת: ' + e.message;
      errEl.style.display = 'block';
      submitEl.disabled = false;
      submitEl.textContent = 'יצירת השוואה';
    }
  }

  filesEl.addEventListener('change', () => { addFiles(filesEl.files); filesEl.value=''; });
  dropEl.addEventListener('click', () => filesEl.click());
  dropEl.addEventListener('dragover', e => { e.preventDefault(); dropEl.classList.add('drag'); });
  dropEl.addEventListener('dragleave', () => dropEl.classList.remove('drag'));
  dropEl.addEventListener('drop', e => {
    e.preventDefault(); dropEl.classList.remove('drag');
    addFiles(e.dataTransfer.files);
  });
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
        html = generate_multi_html(months_data)
        # Store result with a UUID so the browser gets a real URL it can navigate from
        result_id = str(uuid.uuid4())
        _RESULT_CACHE[result_id] = html
        if len(_RESULT_CACHE) > 20:          # keep last 20 results only
            _RESULT_CACHE.popitem(last=False)
        return redirect(f"/multi/result/{result_id}")
    except Exception as e:
        return render_multi_form(f"כשל בקריאת הקבצים: {e}"), 400


@app.get("/multi/result/<result_id>")
def multi_result(result_id: str):
    html = _RESULT_CACHE.get(result_id)
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
