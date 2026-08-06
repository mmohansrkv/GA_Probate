"""
app.py
====================================================================
LNGA / GA Probate Courts Tool — Flask web front-end
====================================================================
Wraps the two pipelines in core.py (Citation Search + Validation) as a
small web app, in place of the original Tkinter desktop GUI:

  /                    home page, choose a module
  /citation-search     upload a County/Start/End file -> live-scrape
                        georgiaprobaterecords.com -> download the
                        Output workbook + Result (Match/New/Missing)
                        workbook
  /validation           upload offline HTML case files + Output Excel(s)
                        + Input Excel/txt -> download the Result
                        workbook (Validation_Result / Case#_Diff_County /
                        Input_vs_Output sheets)
  /jobs/<id>/status     JSON polling endpoint (status + log tail)
  /jobs/<id>/download/<kind>   download a finished job's output file

IMPORTANT — single-process deployment:
Job state lives in an in-memory dict (JOBS). That only works correctly
if the app runs as ONE worker process (many threads is fine — see
render.yaml / the gunicorn command: `--workers 1 --threads 8`). Do not
scale this service to multiple Render instances/workers without moving
job state to something shared (Redis, a DB, etc).
====================================================================
"""

import os
import io
import uuid
import shutil
import threading
import traceback
from datetime import datetime

from flask import (
    Flask, request, jsonify, send_file, render_template_string, abort,
)
from werkzeug.utils import secure_filename

import core

# ------------------------------------------------------------------
# App / config
# ------------------------------------------------------------------
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200 MB upload cap

JOBS_DIR = os.path.join(core.DATA_DIR, "jobs")
os.makedirs(JOBS_DIR, exist_ok=True)

# job_id -> dict(kind, status, log[list], created, error, files{label:path}, stats)
JOBS = {}
JOBS_LOCK = threading.Lock()

ALLOWED_INPUT_EXT = {".xlsx", ".xls", ".txt"}
ALLOWED_HTML_EXT = {".html", ".htm"}
ALLOWED_EXCEL_EXT = {".xlsx", ".xls"}


def _ext_ok(filename, allowed):
    return os.path.splitext(filename)[1].lower() in allowed


def _new_job(kind):
    job_id = uuid.uuid4().hex[:12]
    job_dir = os.path.join(JOBS_DIR, job_id)
    os.makedirs(job_dir, exist_ok=True)
    with JOBS_LOCK:
        JOBS[job_id] = {
            "id": job_id,
            "kind": kind,
            "status": "queued",
            "log": [],
            "created": datetime.now().isoformat(timespec="seconds"),
            "error": None,
            "files": {},
            "stats": {},
            "dir": job_dir,
        }
    return job_id, job_dir


def _log_fn(job_id):
    def _log(msg):
        line = str(msg)
        with JOBS_LOCK:
            job = JOBS.get(job_id)
            if job is not None:
                for sub in line.splitlines() or [""]:
                    job["log"].append(sub)
                # keep the buffer from growing without bound
                if len(job["log"]) > 5000:
                    job["log"] = job["log"][-5000:]
    return _log


def _set(job_id, **kv):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if job is not None:
            job.update(kv)


def _save_upload(file_storage, dest_dir, allowed_ext=None):
    os.makedirs(dest_dir, exist_ok=True)
    filename = secure_filename(file_storage.filename)
    if not filename:
        return None
    if allowed_ext and not _ext_ok(filename, allowed_ext):
        return None
    dest = os.path.join(dest_dir, filename)
    file_storage.save(dest)
    return dest


# ------------------------------------------------------------------
# Citation Search
# ------------------------------------------------------------------

def _run_citation_search(job_id, input_path, compare_path, limit, delay_min, delay_max):
    log = _log_fn(job_id)
    _set(job_id, status="running")
    job_dir = JOBS[job_id]["dir"]
    try:
        success, output_path, match_path, stats = core.run_citation_search_job(
            input_file_path=input_path,
            job_dir=job_dir,
            log=log,
            headless=True,
            limit=limit,
            compare_file=compare_path,
            delay_min=delay_min,
            delay_max=delay_max,
        )
        files = {}
        if output_path and os.path.exists(output_path):
            files["output"] = output_path
        if match_path and os.path.exists(match_path):
            files["result"] = match_path
        _set(job_id, status="done" if success else "error", files=files,
             stats={k: v for k, v in stats.items() if k != "report_df"},
             error=None if success else "Pipeline reported failure — see log.")
    except Exception:
        log("[FATAL] " + traceback.format_exc())
        _set(job_id, status="error", error=traceback.format_exc().strip().splitlines()[-1])


@app.route("/citation-search", methods=["GET", "POST"])
def citation_search():
    if request.method == "GET":
        return render_template_string(CITATION_SEARCH_PAGE)

    input_file = request.files.get("input_file")
    if input_file is None or not input_file.filename:
        return render_template_string(CITATION_SEARCH_PAGE,
                                       error="Please choose a County/Start/End input file (.txt or .xlsx).")

    job_id, job_dir = _new_job("citation_search")
    uploads_dir = os.path.join(job_dir, "uploads")
    input_path = _save_upload(input_file, uploads_dir, ALLOWED_INPUT_EXT)
    if not input_path:
        return render_template_string(
            CITATION_SEARCH_PAGE,
            error="Input file must be .txt, .xlsx, or .xls.")

    compare_path = None
    compare_file = request.files.get("compare_file")
    if compare_file and compare_file.filename:
        compare_path = _save_upload(compare_file, uploads_dir, ALLOWED_EXCEL_EXT)

    limit_raw = (request.form.get("limit") or "").strip()
    limit = int(limit_raw) if limit_raw.isdigit() else None
    delay_min = float(request.form.get("delay_min") or 0) or 0.0
    delay_max = float(request.form.get("delay_max") or 0) or 0.0

    t = threading.Thread(
        target=_run_citation_search,
        args=(job_id, input_path, compare_path, limit, delay_min, delay_max),
        daemon=True,
    )
    t.start()
    return render_template_string(JOB_STATUS_PAGE, job_id=job_id, kind="Citation Search")


# ------------------------------------------------------------------
# Validation
# ------------------------------------------------------------------

def _run_validation(job_id, offline_dir, output_dir, input_dir, result_dir):
    log = _log_fn(job_id)
    _set(job_id, status="running")
    try:
        summary = core.run_validation_job(offline_dir, output_dir, input_dir, result_dir,
                                           log_callback=log)
        files = {}
        out_path = summary.get("out_path")
        if out_path and os.path.exists(out_path):
            files["result"] = out_path
        _set(job_id, status="done", files=files, stats=summary)
    except Exception:
        log("[FATAL] " + traceback.format_exc())
        _set(job_id, status="error", error=traceback.format_exc().strip().splitlines()[-1])


@app.route("/validation", methods=["GET", "POST"])
def validation():
    if request.method == "GET":
        return render_template_string(VALIDATION_PAGE)

    html_files = request.files.getlist("html_files")
    output_files = request.files.getlist("output_files")
    input_files = request.files.getlist("input_files")

    if not any(f.filename for f in html_files):
        return render_template_string(VALIDATION_PAGE,
                                       error="Please upload at least one offline .html/.htm case file.")
    if not any(f.filename for f in output_files):
        return render_template_string(VALIDATION_PAGE,
                                       error="Please upload at least one Output Excel (.xlsx) file.")

    job_id, job_dir = _new_job("validation")
    offline_dir = os.path.join(job_dir, "Offline")
    output_dir = os.path.join(job_dir, "Output")
    input_dir = os.path.join(job_dir, "Input")
    result_dir = os.path.join(job_dir, "Result")
    for d in (offline_dir, output_dir, input_dir, result_dir):
        os.makedirs(d, exist_ok=True)

    for f in html_files:
        if f.filename:
            _save_upload(f, offline_dir, ALLOWED_HTML_EXT)
    for f in output_files:
        if f.filename:
            _save_upload(f, output_dir, ALLOWED_EXCEL_EXT)
    for f in input_files:
        if f.filename:
            _save_upload(f, input_dir, ALLOWED_INPUT_EXT)

    t = threading.Thread(
        target=_run_validation,
        args=(job_id, offline_dir, output_dir, input_dir, result_dir),
        daemon=True,
    )
    t.start()
    return render_template_string(JOB_STATUS_PAGE, job_id=job_id, kind="Validation")


# ------------------------------------------------------------------
# Job status / download
# ------------------------------------------------------------------

@app.route("/jobs/<job_id>/status")
def job_status(job_id):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if job is None:
            abort(404)
        return jsonify({
            "id": job["id"],
            "kind": job["kind"],
            "status": job["status"],
            "log": job["log"][-400:],
            "error": job["error"],
            "stats": job["stats"],
            "downloads": sorted(job["files"].keys()),
        })


@app.route("/jobs/<job_id>/download/<kind>")
def job_download(job_id, kind):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if job is None:
            abort(404)
        path = job["files"].get(kind)
    if not path or not os.path.exists(path):
        abort(404)
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))


@app.route("/jobs/<job_id>/cleanup", methods=["POST"])
def job_cleanup(job_id):
    """Optional: delete a finished job's files from disk to free space."""
    with JOBS_LOCK:
        job = JOBS.pop(job_id, None)
    if job:
        shutil.rmtree(job["dir"], ignore_errors=True)
    return jsonify({"deleted": bool(job)})


# ------------------------------------------------------------------
# Home
# ------------------------------------------------------------------

@app.route("/")
def home():
    return render_template_string(HOME_PAGE)


@app.route("/healthz")
def healthz():
    return jsonify({"status": "ok"})


# ------------------------------------------------------------------
# Inline templates (kept in this file so the deployable file set stays
# exactly app.py / core.py / requirements.txt / render.yaml / README.md)
# ------------------------------------------------------------------

BASE_CSS = """
<style>
  :root{--bg:#F7F8FA;--card:#fff;--accent:#185FA5;--accent-hover:#1D3A6A;
        --border:#E5E7EB;--text:#1A1A1A;--muted:#6B7280;--log-bg:#1E2330;
        --ok:#276221;--err:#9C0006;}
  *{box-sizing:border-box}
  body{background:var(--bg);color:var(--text);font-family:"Segoe UI",Arial,sans-serif;
       margin:0;padding:0 20px 60px;}
  header{padding:36px 0 8px;text-align:center}
  header h1{margin:0;font-size:26px}
  header p{color:var(--muted);margin:6px 0 0}
  .wrap{max-width:760px;margin:0 auto}
  .cards{display:flex;gap:24px;margin-top:28px;flex-wrap:wrap;justify-content:center}
  .card{background:var(--card);border:1px solid var(--border);border-radius:10px;
        padding:22px;flex:1 1 280px;max-width:340px}
  .card h2{margin-top:0;font-size:17px}
  .card p{color:var(--muted);font-size:13px;min-height:54px}
  a.btn,button.btn{display:inline-block;background:var(--accent);color:#fff;border:none;
      padding:10px 18px;border-radius:6px;text-decoration:none;font-weight:600;
      cursor:pointer;font-size:14px}
  a.btn:hover,button.btn:hover{background:var(--accent-hover)}
  form{background:var(--card);border:1px solid var(--border);border-radius:10px;
       padding:24px;margin-top:20px}
  label{display:block;font-weight:600;font-size:13px;margin:14px 0 6px}
  input[type=file],input[type=number],input[type=text]{width:100%;padding:8px;
      border:1px solid var(--border);border-radius:6px;font-size:13px}
  .hint{color:var(--muted);font-size:12px;margin-top:4px}
  .error{background:#FCEBEB;color:var(--err);padding:10px 14px;border-radius:6px;
         font-size:13px;margin-top:14px}
  .row2{display:flex;gap:16px}
  .row2>div{flex:1}
  #log{background:var(--log-bg);color:#D1FAE5;font-family:Consolas,monospace;
       font-size:12.5px;padding:16px;border-radius:8px;height:340px;overflow-y:auto;
       white-space:pre-wrap;margin-top:16px}
  .status{font-weight:700;margin-top:18px}
  .status.done{color:var(--ok)} .status.error{color:var(--err)} .status.running{color:var(--accent)}
  .downloads{margin-top:14px}
  .downloads a{margin-right:12px}
  nav{text-align:center;margin-top:8px}
  nav a{color:var(--accent);text-decoration:none;font-size:13px;margin:0 8px}
</style>
"""

HOME_PAGE = BASE_CSS + """
<header>
  <h1>LNGA Probate Courts</h1>
  <p>Combined Tool — web edition</p>
</header>
<div class="wrap">
  <div class="cards">
    <div class="card">
      <h2>🔎 Citation Search</h2>
      <p>Live-scrapes georgiaprobaterecords.com for every County/Start/End
         row in an uploaded file and builds Output / Result workbooks.</p>
      <a class="btn" href="/citation-search">Open ▶</a>
    </div>
    <div class="card">
      <h2>✅ Validation</h2>
      <p>Compares offline HTML case files against an Output Excel and
         flags field-by-field mismatches.</p>
      <a class="btn" href="/validation">Open ▶</a>
    </div>
  </div>
</div>
"""

CITATION_SEARCH_PAGE = BASE_CSS + """
<header><h1>🔎 Citation Search</h1><p>Upload the County / Starting Date / Ending Date file</p></header>
<nav><a href="/">&larr; Home</a></nav>
<div class="wrap">
  {% if error %}<div class="error">{{ error }}</div>{% endif %}
  <form method="post" enctype="multipart/form-data">
    <label>Input file (County / Starting Date / Ending Date — .txt or .xlsx)</label>
    <input type="file" name="input_file" accept=".txt,.xlsx,.xls" required>
    <div class="hint">Tab-separated .txt header row: County&#9;Starting Date&#9;Ending Date
      (dates DD-MM-YYYY), or an .xlsx with the same columns.</div>

    <label>Existing case-number workbook to compare against (optional)</label>
    <input type="file" name="compare_file" accept=".xlsx,.xls">
    <div class="hint">An .xlsx with a "Case #" column. If omitted, this run's own results
      are saved as a fresh baseline for next time.</div>

    <div class="row2">
      <div>
        <label>Limit counties (optional, for a quick test run)</label>
        <input type="number" name="limit" min="1" placeholder="e.g. 3">
      </div>
      <div>
        <label>Delay between counties, seconds (optional)</label>
        <input type="text" name="delay_min" placeholder="min, e.g. 1">
      </div>
      <div>
        <input type="text" name="delay_max" placeholder="max, e.g. 3" style="margin-top:27px">
      </div>
    </div>

    <div style="margin-top:20px"><button class="btn" type="submit">Run Citation Search ▶</button></div>
  </form>
</div>
"""

VALIDATION_PAGE = BASE_CSS + """
<header><h1>✅ Validation</h1><p>Offline HTML case files vs Output Excel</p></header>
<nav><a href="/">&larr; Home</a></nav>
<div class="wrap">
  {% if error %}<div class="error">{{ error }}</div>{% endif %}
  <form method="post" enctype="multipart/form-data">
    <label>Offline case HTML files (.html/.htm) — select all that apply</label>
    <input type="file" name="html_files" accept=".html,.htm" multiple required>

    <label>Output Excel file(s) (.xlsx)</label>
    <input type="file" name="output_files" accept=".xlsx,.xls" multiple required>

    <label>Input Excel/txt file(s) (optional — enables Input_vs_Output check)</label>
    <input type="file" name="input_files" accept=".xlsx,.xls,.txt" multiple>

    <div style="margin-top:20px"><button class="btn" type="submit">Run Validation ▶</button></div>
  </form>
</div>
"""

JOB_STATUS_PAGE = BASE_CSS + """
<header><h1>{{ kind }}</h1><p>Job <code>{{ job_id }}</code></p></header>
<nav><a href="/">&larr; Home</a></nav>
<div class="wrap">
  <div id="status" class="status">Starting…</div>
  <div id="downloads" class="downloads"></div>
  <div id="log">Waiting for log output…</div>
</div>
<script>
const jobId = "{{ job_id }}";
const statusEl = document.getElementById("status");
const logEl = document.getElementById("log");
const dlEl = document.getElementById("downloads");

async function poll() {
  try {
    const res = await fetch(`/jobs/${jobId}/status`);
    const data = await res.json();
    statusEl.textContent = "Status: " + data.status + (data.error ? " — " + data.error : "");
    statusEl.className = "status " + data.status;
    logEl.textContent = data.log.join("\\n") || "Waiting for log output…";
    logEl.scrollTop = logEl.scrollHeight;

    dlEl.innerHTML = "";
    (data.downloads || []).forEach(kind => {
      const a = document.createElement("a");
      a.className = "btn";
      a.href = `/jobs/${jobId}/download/${kind}`;
      a.textContent = "Download " + kind;
      dlEl.appendChild(a);
    });

    if (data.status === "running" || data.status === "queued") {
      setTimeout(poll, 1500);
    }
  } catch (e) {
    statusEl.textContent = "Lost connection to server — retrying…";
    setTimeout(poll, 3000);
  }
}
poll();
</script>
"""


if __name__ == "__main__":
    # Local dev only. On Render, gunicorn serves the app (see render.yaml).
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
