"""
MCH Nursing Tracer Summary Tool - local web application

Run:   double-click run.bat      (or:  python app.py)
The tool starts a local server on this computer (127.0.0.1) and opens it in
Google Chrome. Files never leave this machine.

All Excel work happens in Python with openpyxl:
    tracer_processor.py  - reads each unit tracer (unit, final score, 0/1 findings)
    summary_writer.py    - fills the uploaded Tracer Summary template
    validator.py         - re-opens the output and checks it
The browser is only the upload / progress / download interface.
"""
import os
import re
import shutil
import socket
import sys
import threading
import time
import traceback
import uuid
import webbrowser
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file

from excel_utils import file_sha1, fmt_pct
from tracer_processor import process_tracer
from summary_writer import (LAYOUT_PER_FINDING, LAYOUT_PER_UNIT, TemplateError,
                            pick_template_sheet, write_summary)
from validator import validate_output

import openpyxl

APP_NAME = "MCH Tracer Summary"
BASE_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
DATA_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = DATA_DIR / "uploads"
OUTPUT_DIR = DATA_DIR / "outputs"
# Create the working folders at import time too, so the app also works under
# Gunicorn (e.g. on Render), where the __main__ block below is not executed.
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
DEFAULT_DOWNLOAD_NAME = "TRACER SUMMARY"
INVALID_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|]')
PREFERRED_PORT = 5051  # the Chart Review tool uses 5050, so both can run together

app = Flask(__name__, template_folder=str(BASE_DIR / "templates"), static_folder=str(BASE_DIR / "static"))
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024

# Single-user local tool: in-memory state is enough.
BATCHES = {}   # batch_id -> {"tracers": [record...], "template": record|None}
JOBS = {}      # job_id   -> job dict
LOCK = threading.Lock()


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
def _safe_name(filename):
    base = os.path.basename(filename or "file.xlsx")
    return "".join(c for c in base if c.isalnum() or c in "._- ()&") or "file.xlsx"


def _sanitize_output_filename(raw):
    name = (raw or "").strip()
    name = re.sub(r"\.xlsx?$", "", name, flags=re.IGNORECASE)
    name = INVALID_FILENAME_CHARS.sub("", name)
    name = re.sub(r"\s+", " ", name).strip().rstrip(". ")[:150]
    return name or DEFAULT_DOWNLOAD_NAME


def _batch(batch_id=None, create=True):
    with LOCK:
        if batch_id and batch_id in BATCHES:
            return batch_id, BATCHES[batch_id]
        if not create:
            return None, None
        bid = batch_id or uuid.uuid4().hex
        BATCHES[bid] = {"tracers": [], "template": None}
        (UPLOAD_DIR / bid).mkdir(parents=True, exist_ok=True)
        return bid, BATCHES[bid]


def _analyze_tracer(rec):
    """Run the real extraction once at upload time so the list shows unit / score / findings."""
    res = process_tracer(rec["path"])
    rec.update({
        "valid": res.ok,
        "unit": res.unit,
        "score": res.final_score,
        "score_text": fmt_pct(res.final_score) if res.final_score is not None else "",
        "findings": len(res.findings) if res.ok else 0,
        "items": len(res.items),
        "errors": list(res.errors),
        "warnings": list(res.warnings),
    })


def _refresh_flags(batch):
    """Duplicate-file and same-unit flags (warnings - the user decides to keep or remove)."""
    seen_hash, seen_unit = {}, {}
    for rec in batch["tracers"]:
        rec["duplicate_of"] = None
        rec["same_unit_as"] = None
        if rec["sha1"] in seen_hash:
            rec["duplicate_of"] = seen_hash[rec["sha1"]]
        else:
            seen_hash[rec["sha1"]] = rec["filename"]
        if rec.get("valid") and rec.get("unit"):
            key = rec["unit"].strip().upper()
            if key in seen_unit and not rec["duplicate_of"]:
                rec["same_unit_as"] = seen_unit[key]
            seen_unit.setdefault(key, rec["filename"])


def _public_batch(bid, batch):
    _refresh_flags(batch)
    tracers = []
    for r in batch["tracers"]:
        tracers.append({k: r.get(k) for k in ("id", "filename", "valid", "unit", "score_text", "findings",
                                               "items", "errors", "warnings", "duplicate_of", "same_unit_as")})
    return {"batch_id": bid, "tracers": tracers, "template": batch["template"] and {
        k: batch["template"].get(k) for k in ("filename", "valid", "error", "sheet")}}


# ---------------------------------------------------------------------------
# pages / API
# ---------------------------------------------------------------------------
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/batch/<bid>", methods=["GET"])
def get_batch(bid):
    bid, batch = _batch(bid, create=False)
    if not batch:
        return jsonify({"error": "Session expired - please reload the page."}), 404
    return jsonify(_public_batch(bid, batch))


@app.route("/api/tracers", methods=["POST"])
def upload_tracers():
    bid, batch = _batch(request.form.get("batch_id"))
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "No files received."}), 400
    rejected = []
    for f in files:
        name = f.filename or "file.xlsx"
        if os.path.splitext(name)[1].lower() != ".xlsx":
            rejected.append("%s (not an .xlsx file)" % name)
            continue
        fid = uuid.uuid4().hex[:12]
        path = UPLOAD_DIR / bid / ("%s__%s" % (fid, _safe_name(name)))
        f.save(str(path))
        rec = {"id": fid, "filename": os.path.basename(name), "path": str(path), "sha1": file_sha1(str(path))}
        try:
            _analyze_tracer(rec)
        except Exception as e:  # never crash the upload
            rec.update({"valid": False, "errors": ["Could not read file: %s" % e], "warnings": [],
                        "unit": "", "score_text": "", "findings": 0, "items": 0})
        batch["tracers"].append(rec)
    out = _public_batch(bid, batch)
    out["rejected"] = rejected
    return jsonify(out)


@app.route("/api/tracers/<bid>/<fid>", methods=["DELETE"])
def delete_tracer(bid, fid):
    bid, batch = _batch(bid, create=False)
    if not batch:
        return jsonify({"error": "Session expired - please reload the page."}), 404
    for r in list(batch["tracers"]):
        if r["id"] == fid:
            batch["tracers"].remove(r)
            try:
                os.remove(r["path"])
            except OSError:
                pass
    return jsonify(_public_batch(bid, batch))


@app.route("/api/tracers/<bid>/clear", methods=["POST"])
def clear_tracers(bid):
    bid, batch = _batch(bid, create=False)
    if batch:
        for r in batch["tracers"]:
            try:
                os.remove(r["path"])
            except OSError:
                pass
        batch["tracers"] = []
        return jsonify(_public_batch(bid, batch))
    return jsonify({"error": "Session expired - please reload the page."}), 404


@app.route("/api/tracers/<bid>/order", methods=["POST"])
def reorder(bid):
    bid, batch = _batch(bid, create=False)
    if not batch:
        return jsonify({"error": "Session expired - please reload the page."}), 404
    order = (request.get_json(silent=True) or {}).get("order", [])
    pos = {fid: i for i, fid in enumerate(order)}
    batch["tracers"].sort(key=lambda r: pos.get(r["id"], 10 ** 6))
    return jsonify(_public_batch(bid, batch))


@app.route("/api/template", methods=["POST"])
def upload_template():
    bid, batch = _batch(request.form.get("batch_id"))
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "No file received."}), 400
    name = f.filename or "template.xlsx"
    if os.path.splitext(name)[1].lower() != ".xlsx":
        return jsonify({"error": "The template must be an .xlsx file."}), 400
    path = UPLOAD_DIR / bid / ("template__%s" % _safe_name(name))
    f.save(str(path))
    rec = {"filename": os.path.basename(name), "path": str(path), "valid": True, "error": None, "sheet": None}
    try:
        ws = pick_template_sheet(openpyxl.load_workbook(str(path)))
        rec["sheet"] = ws.title
    except TemplateError as e:
        rec.update(valid=False, error=str(e))
    except Exception as e:
        rec.update(valid=False, error="Could not open workbook (%s)." % e)
    batch["template"] = rec
    return jsonify(_public_batch(bid, batch))


@app.route("/api/template/<bid>", methods=["DELETE"])
def delete_template(bid):
    bid, batch = _batch(bid, create=False)
    if batch:
        batch["template"] = None
        return jsonify(_public_batch(bid, batch))
    return jsonify({"error": "Session expired - please reload the page."}), 404


# ---------------------------------------------------------------------------
# generation job (runs in a thread; the page polls /api/job/<id>)
# ---------------------------------------------------------------------------
@app.route("/api/generate", methods=["POST"])
def generate():
    data = request.get_json(silent=True) or {}
    bid, batch = _batch(data.get("batch_id"), create=False)
    if not batch:
        return jsonify({"error": "Session expired - please reload the page."}), 404
    if not batch["tracers"]:
        return jsonify({"error": "Please upload at least one Tracer Excel file."}), 400
    tpl = batch["template"]
    if not tpl:
        return jsonify({"error": "Please upload the Tracer Summary Template."}), 400
    if not tpl["valid"]:
        return jsonify({"error": "The Summary template is not valid: %s" % tpl["error"]}), 400
    layout = LAYOUT_PER_UNIT  # ONE ROW = ONE UNIT/WARD (all findings of the unit in one cell)
    month = None
    m = (data.get("month") or "").strip()  # "YYYY-MM" from <input type=month>
    if m:
        try:
            y, mo = m.split("-")[:2]
            import datetime as dt
            month = dt.datetime(int(y), int(mo), 1)
        except Exception:
            return jsonify({"error": "MONTH\\YEAR not understood."}), 400
    job_id = uuid.uuid4().hex
    job = {
        "id": job_id, "status": "running", "log": [], "current": 0, "total": len(batch["tracers"]),
        "current_file": "", "stage": "processing", "result": None,
        "download_name": _sanitize_output_filename(data.get("output_name")) + ".xlsx",
    }
    JOBS[job_id] = job
    tracers = [dict(r) for r in batch["tracers"]]
    threading.Thread(target=_run_job, args=(job, tracers, dict(tpl), layout, month), daemon=True).start()
    return jsonify({"job_id": job_id})


def _run_job(job, tracers, tpl, layout, month):
    def log(msg, kind="info"):
        job["log"].append({"t": kind, "m": msg})

    try:
        ok, failed = [], []
        n = len(tracers)
        for i, rec in enumerate(tracers, 1):
            job.update(current=i, current_file=rec["filename"])
            log("Processing %d of %d — %s" % (i, n, rec["filename"]), "head")
            res = process_tracer(rec["path"])
            res.path = rec["filename"]  # show the original name, not the stored one
            for w in res.warnings:
                log("WARNING: " + w, "warn")
            if not res.ok:
                for e in res.errors:
                    log("ERROR: %s — %s" % (rec["filename"], e), "err")
                failed.append({"filename": rec["filename"], "errors": res.errors})
                continue
            log("✓ Unit detected: %s" % res.unit, "ok")
            log("✓ Final score detected: %s" % fmt_pct(res.final_score), "ok")
            log("✓ %d findings extracted  (score 0: %d · score 1: %d · excluded 2: %d · NA: %d)"
                % (len(res.findings), res.count(0), res.count(1), res.count(2), res.count("NA")), "ok")
            if rec.get("duplicate_of"):
                log("WARNING: Duplicate file detected: %s (same as %s) — kept as requested"
                    % (rec["filename"], rec["duplicate_of"]), "warn")
            ok.append(res)

        if not ok:
            job.update(status="failed", stage="done")
            job["result"] = {"success": False, "files_processed": 0, "findings": 0,
                             "errors": len(failed), "failed": failed, "checks": [],
                             "message": "No tracer file could be processed."}
            return

        job.update(stage="writing", current_file="")
        log("Writing the Tracer Summary into the template…", "head")
        out_path = OUTPUT_DIR / ("%s.xlsx" % job["id"])
        info = write_summary(tpl["path"], ok, str(out_path), layout=layout, month_value=month,
                             log=lambda m: log(m, "info"))
        job.update(stage="validating")
        log("Validating the generated workbook…", "head")
        checks = validate_output(str(out_path), tpl["path"], ok, info)
        for passed, msg in checks:
            log(("✓ " if passed else "✗ FAILED: ") + msg, "ok" if passed else "err")
        n_fail = sum(1 for p, _ in checks if not p)
        warnings = sum(1 for e in job["log"] if e["t"] == "warn")
        job["result"] = {
            "success": n_fail == 0,
            "files_processed": len(ok),
            "findings": sum(len(r.findings) for r in ok),
            "rows": len(info["rows"]),
            "errors": len(failed),
            "warnings": warnings,
            "failed": failed,
            "checks_passed": len(checks) - n_fail,
            "checks_total": len(checks),
            "checks": [{"ok": p, "m": m} for p, m in checks],
            "units": [{"unit": r.unit, "score": fmt_pct(r.final_score), "findings": len(r.findings),
                       "file": r.file_name} for r in ok],
            "download_url": "/api/download/%s" % job["id"],
            "download_name": job["download_name"],
        }
        job.update(status="done", stage="done")
    except PermissionError:
        log("ERROR: could not save the output file (permission denied).", "err")
        job.update(status="failed", stage="done", result={"success": False, "message": "Could not save the output."})
    except Exception as e:
        log("ERROR: " + str(e), "err")
        log(traceback.format_exc(), "err")
        job.update(status="failed", stage="done", result={"success": False, "message": str(e)})


@app.route("/api/job/<job_id>")
def job_status(job_id):
    job = JOBS.get(job_id)
    if not job:
        return jsonify({"error": "Unknown job."}), 404
    since = int(request.args.get("since", 0))
    out = {k: job[k] for k in ("status", "current", "total", "current_file", "stage", "result")}
    out.update(log=job["log"][since:], log_len=len(job["log"]))
    return jsonify(out)


@app.route("/api/download/<job_id>")
def download(job_id):
    job = JOBS.get(job_id)
    path = OUTPUT_DIR / ("%s.xlsx" % job_id)
    if not job or not path.exists():
        return jsonify({"error": "File not found - please generate again."}), 404
    return send_file(str(path), as_attachment=True, download_name=job["download_name"],
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ---------------------------------------------------------------------------
# startup
# ---------------------------------------------------------------------------
def _clean_work_dirs():
    for d in (UPLOAD_DIR, OUTPUT_DIR):
        shutil.rmtree(d, ignore_errors=True)
        d.mkdir(parents=True, exist_ok=True)


def _free_port(preferred):
    for port in [preferred] + list(range(preferred + 1, preferred + 30)):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    return 0


def _open_browser(url):
    """Open Google Chrome if installed, otherwise the default browser."""
    if os.environ.get("TRACER_NO_BROWSER"):
        return
    if sys.platform == "win32":
        candidates = [
            os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"), r"Google\Chrome\Application\chrome.exe"),
            os.path.join(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"), r"Google\Chrome\Application\chrome.exe"),
            os.path.join(os.environ.get("LOCALAPPDATA", ""), r"Google\Chrome\Application\chrome.exe"),
        ]
        for c in candidates:
            if c and os.path.isfile(c):
                import subprocess
                subprocess.Popen([c, url])
                return
    webbrowser.open(url)


if __name__ == "__main__":
    _clean_work_dirs()
    port = _free_port(PREFERRED_PORT)
    url = "http://127.0.0.1:%d" % port
    print("=" * 64)
    print("  %s is running at  %s" % (APP_NAME, url))
    print("  Your browser will open automatically.")
    print("  Keep this window open while you use the tool.")
    print("  Close this window to stop the tool.")
    print("=" * 64)
    import logging
    import flask.cli
    logging.getLogger("werkzeug").setLevel(logging.ERROR)   # no request log / dev-server notice
    flask.cli.show_server_banner = lambda *a, **k: None
    threading.Timer(1.2, _open_browser, args=(url,)).start()
    app.run(host="127.0.0.1", port=port, debug=False, threaded=True)
