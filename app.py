"""
app.py — LNGA GA Probate Courts web app.

Wraps core.py's two pipelines (Citation Search live-scrape, and
offline Validation) behind:
  - a small web UI (upload a file, click Run, watch the log, download
    the result workbook), and
  - a JSON API doing the same thing, for scripted/automated use.

Runs are asynchronous: POSTing a job returns a job_id immediately; the
scrape/validation runs in a background thread while the UI polls
/api/jobs/<id> for log lines and status. This matters because a full
Citation Search run over 20+ counties, each paging through a live
ASP.NET grid, can take several minutes — a synchronous request would
just time out.

NOTE: the Citation Search feature drives a real (headless) Chrome
browser against https://georgiaprobaterecords.com/Traffic/SearchCitations.aspx.
It needs Chrome/Chromium + a matching driver installed in the
container (see Dockerfile) and outbound network access to that site.
The Validation feature is fully offline (HTML/Excel files you upload)
and has no such requirement.
"""

import io
import os
import shutil
import tempfile
import threading
import traceback
import uuid
import zipfile
from datetime import datetime

from flask import (
    Flask, request, jsonify, send_file, render_template_string, abort
)
from werkzeug.utils import secure_filename

import core

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200 MB uploads

# Where per-job scratch folders + results live on disk. On Render this
# should point at a mounted persistent disk if you want results to survive
# a restart; by default it's ephemeral container storage, which is fine
# since results are meant to be downloaded right after the run.
DATA_ROOT = os.environ.get("LNGA_DATA_ROOT", "/tmp/lnga_jobs")
os.makedirs(DATA_ROOT, exist_ok=True)

# In-memory job registry: {job_id: {...}}. Fine for a single-process
# deployment (Render's free/starter web service tier); if you scale to
# multiple instances, swap this for Redis or a DB-backed job table.
JOBS = {}
JOBS_LOCK = threading.Lock()


# ─────────────────────────────────────────────────────────────────────────
# Job helpers
# ─────────────────────────────────────────────────────────────────────────
def _new_job(kind):
    job_id = uuid.uuid4().hex[:12]
    job_dir = os.path.join(DATA_ROOT, job_id)
    os.makedirs(job_dir, exist_ok=True)
    with JOBS_LOCK:
        JOBS[job_id] = {
            "id": job_id,
            "kind": kind,
            "status": "pending",  # pending -> running -> done | error
            "log": [],
            "created": datetime.utcnow().isoformat() + "Z",
            "dir": job_dir,
            "downloads": {},   # label -> absolute path
            "stats": {},
            "error": None,
        }
    return job_id


def _log(job_id, line):
    with JOBS_LOCK:
        JOBS[job_id]["log"].append(str(line))


def _set_status(job_id, status, **extra):
    with JOBS_LOCK:
        JOBS[job_id]["status"] = status
        JOBS[job_id].update(extra)


def _get_job(job_id):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if job is None:
            return None
        return dict(job)  # shallow copy for safe read outside the lock


def _extract_upload_to_dir(file_storage, dest_dir):
    """Saves an uploaded file into dest_dir. If it's a .zip, extracts its
    contents into dest_dir instead (so callers can upload a whole folder
    of offline HTML dumps / Excel files as one .zip)."""
    os.makedirs(dest_dir, exist_ok=True)
    filename = secure_filename(file_storage.filename or "upload")
    if filename.lower().endswith(".zip"):
        data = io.BytesIO(file_storage.read())
        with zipfile.ZipFile(data) as zf:
            zf.extractall(dest_dir)
    else:
        file_storage.save(os.path.join(dest_dir, filename))
    return dest_dir


# ─────────────────────────────────────────────────────────────────────────
# Citation Search (live scrape) job
# ─────────────────────────────────────────────────────────────────────────
def _run_citation_search_job(job_id, input_path, compare_path, opts):
    def log(msg):
        _log(job_id, msg)

    def on_done(success, output_path, match_path, stats):
        if success:
            downloads = {}
            if output_path:
                downloads["result_workbook"] = output_path
            if match_path:
                downloads["match_status_workbook"] = match_path
            _set_status(job_id, "done", downloads=downloads, stats=stats)
        else:
            _set_status(job_id, "error", error="Pipeline failed — see log.")

    job = _get_job(job_id)
    cfg = {
        "input_file": input_path,
        "output_folder": os.path.join(job["dir"], "Output"),
        "offline_folder": os.path.join(job["dir"], "Offline"),
        "tool_input_folder": os.path.join(job["dir"], "Tool_Input"),
        "compare_file": compare_path,
        "headless": opts.get("headless", True),
        "limit": opts.get("limit") or 0,
        "delay_seconds_minimum": opts.get("delay_min", 0),
        "delay_seconds_maximum": opts.get("delay_max", 0),
    }
    core.ensure_folder(cfg["output_folder"])
    core.ensure_folder(cfg["offline_folder"])
    core.ensure_folder(cfg["tool_input_folder"])

    _set_status(job_id, "running")
    try:
        core.run_live_pipeline(cfg, log, on_done)
    except Exception:
        _log(job_id, traceback.format_exc())
        _set_status(job_id, "error", error="Unhandled exception — see log.")


# ─────────────────────────────────────────────────────────────────────────
# Validation job
# ─────────────────────────────────────────────────────────────────────────
def _run_validation_job(job_id, offline_dir, output_dir, input_dir):
    _set_status(job_id, "running")
    try:
        job = _get_job(job_id)
        result_base = os.path.join(job["dir"], "Result")
        _log(job_id, "Starting validation...")
        info = core.run_validation(offline_dir, output_dir, input_dir, result_base)
        _log(job_id, f"Done. {info['total_cases']} case(s) processed, "
                      f"{info['matched_cases']} matched, {info['mismatched_cases']} mismatched, "
                      f"{info['new_case_count']} new, {info['not_found_cases']} not found.")
        _set_status(job_id, "done",
                    downloads={"validation_result": info["out_path"]},
                    stats=info)
    except Exception:
        _log(job_id, traceback.format_exc())
        _set_status(job_id, "error", error="Validation failed — see log.")


# ─────────────────────────────────────────────────────────────────────────
# API routes
# ─────────────────────────────────────────────────────────────────────────
@app.route("/api/citation-search", methods=["POST"])
def api_citation_search():
    """multipart/form-data:
       input_file    (required) - .txt or .xlsx: County / Starting Date / Ending Date
       compare_file  (optional) - existing case-number .xlsx to validate the scrape against
       headless      (optional, default true)
       limit         (optional) - only process the first N counties (testing)
       delay_min / delay_max (optional) - seconds to sleep between counties
    """
    if not core.SELENIUM_AVAILABLE:
        return jsonify(error=("selenium is not installed in this deployment — "
                               "the Citation Search feature is unavailable.")), 503
    if "input_file" not in request.files or request.files["input_file"].filename == "":
        return jsonify(error="input_file is required (.txt or .xlsx)"), 400

    job_id = _new_job("citation_search")
    job = _get_job(job_id)

    input_path = os.path.join(job["dir"], secure_filename(request.files["input_file"].filename))
    request.files["input_file"].save(input_path)

    compare_path = None
    if "compare_file" in request.files and request.files["compare_file"].filename:
        compare_path = os.path.join(job["dir"], secure_filename(request.files["compare_file"].filename))
        request.files["compare_file"].save(compare_path)

    opts = {
        "headless": request.form.get("headless", "true").lower() != "false",
        "limit": request.form.get("limit", type=int) or 0,
        "delay_min": request.form.get("delay_min", type=float) or 0,
        "delay_max": request.form.get("delay_max", type=float) or 0,
    }

    t = threading.Thread(target=_run_citation_search_job,
                          args=(job_id, input_path, compare_path, opts), daemon=True)
    t.start()
    return jsonify(job_id=job_id), 202


@app.route("/api/validate", methods=["POST"])
def api_validate():
    """multipart/form-data:
       offline  (required) - .zip of debug_<county>.html offline dumps
       output   (required) - .zip of Output Excel workbook(s) (or a single .xlsx)
       input    (required) - .zip of Input Excel workbook(s) (or a single .xlsx)
    """
    for field in ("offline", "output", "input"):
        if field not in request.files or request.files[field].filename == "":
            return jsonify(error=f"'{field}' file is required"), 400

    job_id = _new_job("validation")
    job = _get_job(job_id)

    offline_dir = _extract_upload_to_dir(request.files["offline"], os.path.join(job["dir"], "Offline"))
    output_dir = _extract_upload_to_dir(request.files["output"], os.path.join(job["dir"], "Output"))
    input_dir = _extract_upload_to_dir(request.files["input"], os.path.join(job["dir"], "Input"))

    t = threading.Thread(target=_run_validation_job,
                          args=(job_id, offline_dir, output_dir, input_dir), daemon=True)
    t.start()
    return jsonify(job_id=job_id), 202


@app.route("/api/jobs/<job_id>", methods=["GET"])
def api_job_status(job_id):
    job = _get_job(job_id)
    if job is None:
        abort(404)
    return jsonify({
        "id": job["id"],
        "kind": job["kind"],
        "status": job["status"],
        "log": job["log"],
        "stats": job["stats"],
        "error": job["error"],
        "downloads": sorted(job["downloads"].keys()),
    })


@app.route("/api/jobs/<job_id>/download/<label>", methods=["GET"])
def api_job_download(job_id, label):
    job = _get_job(job_id)
    if job is None or label not in job["downloads"]:
        abort(404)
    path = job["downloads"][label]
    if not os.path.isfile(path):
        abort(404)
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))


@app.route("/healthz")
def healthz():
    return jsonify(status="ok", selenium_available=core.SELENIUM_AVAILABLE)


# ─────────────────────────────────────────────────────────────────────────
# Web UI (single page, vanilla JS polling — no separate templates/ dir)
# ─────────────────────────────────────────────────────────────────────────
INDEX_HTML = """
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>LNGA GA Probate Courts</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
  :root { --accent:#185FA5; --ink:#1A1A1A; --sub:#6B7280; --bg:#F7F8FA; --card:#fff; --border:#E5E7EB; }
  body { font-family: -apple-system, Segoe UI, Arial, sans-serif; background:var(--bg); color:var(--ink); margin:0; }
  header { background:#1D3A6A; color:#fff; padding:18px 28px; }
  header h1 { margin:0; font-size:20px; }
  main { max-width:920px; margin:24px auto; padding:0 16px; }
  .tabs { display:flex; gap:8px; margin-bottom:16px; }
  .tab { padding:10px 18px; border-radius:8px 8px 0 0; background:var(--border); color:var(--sub); cursor:pointer; font-weight:600; }
  .tab.active { background:var(--card); color:var(--accent); }
  .card { background:var(--card); border:1px solid var(--border); border-radius:0 8px 8px 8px; padding:20px; }
  .panel { display:none; } .panel.active { display:block; }
  label { display:block; font-size:13px; font-weight:600; margin:14px 0 4px; }
  input[type=file], input[type=number], input[type=text] { width:100%; padding:8px; border:1px solid var(--border); border-radius:6px; box-sizing:border-box; }
  .row { display:flex; gap:16px; } .row > div { flex:1; }
  button { margin-top:18px; background:var(--accent); color:#fff; border:none; padding:10px 20px; border-radius:6px; font-weight:600; cursor:pointer; }
  button:disabled { opacity:.5; cursor:default; }
  .hint { color:var(--sub); font-size:12px; margin-top:4px; }
  pre#log { background:#1E2330; color:#D1FAE5; padding:14px; border-radius:8px; max-height:320px; overflow:auto; font-size:12px; margin-top:18px; white-space:pre-wrap; }
  #downloads a { display:inline-block; margin:10px 10px 0 0; background:#EAF3DE; color:#276221; padding:8px 14px; border-radius:6px; text-decoration:none; font-weight:600; }
  .status { font-weight:700; margin-top:14px; }
  .status.error { color:#9C0006; } .status.done { color:#276221; } .status.running,.status.pending { color:#7F3F00; }
  .checkline { display:flex; align-items:center; gap:8px; margin-top:14px; }
</style>
</head>
<body>
<header><h1>LNGA GA Probate Courts</h1></header>
<main>
  <div class="tabs">
    <div class="tab active" data-tab="citation">Citation Search</div>
    <div class="tab" data-tab="validation">Validation</div>
  </div>

  <div class="card">
    <div class="panel active" id="panel-citation">
      <form id="form-citation">
        <label>Input file (.txt or .xlsx — County / Starting Date / Ending Date)</label>
        <input type="file" name="input_file" required>
        <div class="hint">Tab-separated .txt or Excel, same columns as your weekly county list.</div>

        <label>Compare / existing case-number file (optional .xlsx)</label>
        <input type="file" name="compare_file">
        <div class="hint">If provided, the scrape is checked against it for Match / New / Missing case numbers.</div>

        <div class="row">
          <div>
            <label>Limit counties (optional, for testing)</label>
            <input type="number" name="limit" min="0" placeholder="e.g. 3">
          </div>
          <div>
            <label>Delay between counties: min / max seconds</label>
            <div class="row">
              <input type="number" name="delay_min" min="0" step="0.5" placeholder="0">
              <input type="number" name="delay_max" min="0" step="0.5" placeholder="0">
            </div>
          </div>
        </div>

        <div class="checkline">
          <input type="checkbox" name="headless" id="headless" checked>
          <label for="headless" style="margin:0;">Run headless (unchecked = only relevant for local debugging)</label>
        </div>

        <button type="submit">Run Citation Search</button>
      </form>
    </div>

    <div class="panel" id="panel-validation">
      <form id="form-validation">
        <label>Offline HTML folder (.zip of debug_&lt;county&gt;.html files)</label>
        <input type="file" name="offline" accept=".zip" required>

        <label>Output Excel folder (.zip, or a single .xlsx)</label>
        <input type="file" name="output" required>

        <label>Input Excel folder (.zip, or a single .xlsx)</label>
        <input type="file" name="input" required>

        <button type="submit">Run Validation</button>
      </form>
    </div>

    <div class="status" id="status"></div>
    <div id="downloads"></div>
    <pre id="log" style="display:none;"></pre>
  </div>
</main>

<script>
document.querySelectorAll(".tab").forEach(t => t.addEventListener("click", () => {
  document.querySelectorAll(".tab").forEach(x => x.classList.remove("active"));
  document.querySelectorAll(".panel").forEach(x => x.classList.remove("active"));
  t.classList.add("active");
  document.getElementById("panel-" + t.dataset.tab).classList.add("active");
}));

let poller = null;

function resetOutput() {
  document.getElementById("status").className = "status";
  document.getElementById("status").textContent = "";
  document.getElementById("downloads").innerHTML = "";
  document.getElementById("log").style.display = "none";
  document.getElementById("log").textContent = "";
  if (poller) clearInterval(poller);
}

function poll(jobId) {
  poller = setInterval(async () => {
    const r = await fetch("/api/jobs/" + jobId);
    const data = await r.json();
    document.getElementById("log").style.display = "block";
    document.getElementById("log").textContent = data.log.join("\\n");
    document.getElementById("log").scrollTop = 1e9;
    const statusEl = document.getElementById("status");
    statusEl.textContent = "Status: " + data.status;
    statusEl.className = "status " + data.status;
    if (data.status === "done" || data.status === "error") {
      clearInterval(poller);
      const dl = document.getElementById("downloads");
      dl.innerHTML = "";
      (data.downloads || []).forEach(label => {
        const a = document.createElement("a");
        a.href = "/api/jobs/" + jobId + "/download/" + label;
        a.textContent = "Download: " + label;
        dl.appendChild(a);
      });
      if (data.error) {
        statusEl.textContent += " — " + data.error;
      }
    }
  }, 1500);
}

async function submitForm(formId, url) {
  const form = document.getElementById(formId);
  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    resetOutput();
    const btn = form.querySelector("button");
    btn.disabled = true;
    try {
      const fd = new FormData(form);
      if (form.querySelector("[name=headless]")) {
        fd.set("headless", form.querySelector("[name=headless]").checked ? "true" : "false");
      }
      const r = await fetch(url, { method: "POST", body: fd });
      const data = await r.json();
      if (!r.ok) {
        document.getElementById("status").className = "status error";
        document.getElementById("status").textContent = data.error || "Request failed.";
        return;
      }
      poll(data.job_id);
    } finally {
      btn.disabled = false;
    }
  });
}
submitForm("form-citation", "/api/citation-search");
submitForm("form-validation", "/api/validate");
</script>
</body>
</html>
"""


@app.route("/")
def index():
    return render_template_string(INDEX_HTML)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
