"""
app.py
Browser-based version of the pipeline. Open a URL, drag in your Excel/CSV,
click one button -- the report is built and (optionally) emailed. No command
line needed once this is running.

Also supports:
  - Managing recipient emails from the page itself (no .env editing needed)
  - Scheduling a daily automatic run (uses whichever file was uploaded most
    recently) so you can set it and forget it, right from the portal
  - A quick data preview (detected columns + row count) once a file is read
  - A history of past reports, downloadable again within their retention window
  - Automatic cleanup of old generated files so disk usage doesn't grow forever
  - Optional Postgres-backed settings storage (DATABASE_URL) so recipients and
    schedule survive a redeploy -- falls back to a local JSON file if unset
  - Optional Sentry error monitoring (SENTRY_DSN)
  - Basic login rate-limiting and an upload size cap

Run locally:   python app.py            -> open http://127.0.0.1:5000
Deployment:    see README.md "Deploy as a web app" section (Render, no CLI needed)
"""

import os
import re
import sys
import json
import time
import uuid
import shutil
import threading
import traceback
from functools import wraps
from datetime import datetime, timedelta

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from flask import Flask, request, jsonify, render_template, send_file, url_for, session, redirect
from dotenv import load_dotenv
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
from apscheduler.triggers.interval import IntervalTrigger

from src import ingest, process, charts, report_builder
from src.email_sender import send_report

load_dotenv()

# ---------------------------------------------------------------------------
# Optional error monitoring (Sentry) -- only activates if SENTRY_DSN is set.
# Free tier at sentry.io is plenty for a tool like this.
# ---------------------------------------------------------------------------
SENTRY_DSN = os.getenv("SENTRY_DSN")
if SENTRY_DSN:
    try:
        import sentry_sdk
        from sentry_sdk.integrations.flask import FlaskIntegration
        sentry_sdk.init(dsn=SENTRY_DSN, integrations=[FlaskIntegration()], traces_sample_rate=0.1)
    except ImportError:
        print("SENTRY_DSN is set but sentry-sdk isn't installed -- run: pip install sentry-sdk")

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", os.urandom(24).hex())
APP_PASSWORD = os.getenv("APP_PASSWORD")  # set this to lock the page down; leave unset for local-only use

# Session cookies are non-permanent (expire when the browser is closed) and
# HttpOnly (not readable by JS). Combined with the no-store headers below,
# this stops a logged-out session from appearing "still logged in" via the
# browser's back/forward cache.
app.config["SESSION_PERMANENT"] = False
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"

# Reject uploads bigger than this before they ever reach the pipeline (default 20 MB).
app.config["MAX_CONTENT_LENGTH"] = int(os.getenv("MAX_UPLOAD_MB", "20")) * 1024 * 1024

UPLOAD_DIR = os.path.join("output", "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)
CONFIG_PATH = os.path.join("output", "config.json")

# Generated files older than this get cleaned up automatically so disk usage
# doesn't grow forever on a long-running server.
FILE_RETENTION_HOURS = int(os.getenv("FILE_RETENTION_HOURS", "24"))

# In-memory job tracking -- fine for a single-user internal tool. Report
# metadata (for history) is persisted separately so it survives across jobs.
JOBS = {}

STEP_ORDER = ["reading", "cleaning", "charting", "building", "emailing", "done"]


# ---------------------------------------------------------------------------
# Config storage: recipients, schedule, and report history.
#
# If DATABASE_URL is set (e.g. a free Postgres instance from Supabase/Render),
# settings are stored there and survive redeploys. Otherwise they fall back to
# a local JSON file, which works fine for local/desktop use but resets when a
# host with an ephemeral filesystem (like Render's free tier) redeploys.
# ---------------------------------------------------------------------------
DATABASE_URL = os.getenv("DATABASE_URL")
_db_engine = None

if DATABASE_URL:
    try:
        from sqlalchemy import create_engine, text
        _db_engine = create_engine(DATABASE_URL, pool_pre_ping=True)
        with _db_engine.begin() as conn:
            conn.execute(text("CREATE TABLE IF NOT EXISTS app_config (key TEXT PRIMARY KEY, value TEXT)"))
        print("Using Postgres-backed settings storage (DATABASE_URL is set).")
    except Exception as e:
        print(f"Could not connect to DATABASE_URL ({e}) -- falling back to local JSON file.")
        _db_engine = None


def default_config():
    env_recipients = [r.strip() for r in os.getenv("REPORT_RECIPIENTS", "").split(",") if r.strip()]
    return {
        "recipients": env_recipients,
        "schedule_enabled": False,
        "schedule_time": "09:00",
        "latest_file": None,
        "last_run_at": None,
        "last_run_status": None,
        "reports": [],  # history: [{id, filename, path, generated_at, size_kb}, ...]
    }


def load_config():
    merged = default_config()
    if _db_engine is not None:
        from sqlalchemy import text
        try:
            with _db_engine.begin() as conn:
                row = conn.execute(text("SELECT value FROM app_config WHERE key = 'main'")).fetchone()
            if row:
                merged.update(json.loads(row[0]))
            return merged
        except Exception as e:
            print(f"DB read failed ({e}), falling back to defaults for this request.")
            return merged

    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH) as f:
                merged.update(json.load(f))
        except (json.JSONDecodeError, OSError):
            pass
    return merged


def save_config(cfg):
    if _db_engine is not None:
        from sqlalchemy import text
        try:
            with _db_engine.begin() as conn:
                conn.execute(
                    text(
                        "INSERT INTO app_config (key, value) VALUES ('main', :v) "
                        "ON CONFLICT (key) DO UPDATE SET value = :v"
                    ),
                    {"v": json.dumps(cfg, default=str)},
                )
            return
        except Exception as e:
            print(f"DB write failed ({e}) -- settings for this run were not persisted.")
            return

    os.makedirs(os.path.dirname(CONFIG_PATH), exist_ok=True)
    with open(CONFIG_PATH, "w") as f:
        json.dump(cfg, f, indent=2, default=str)


def record_report(output_path, send_email):
    """Adds a completed report to the history list, kept to the most recent 30."""
    cfg = load_config()
    reports = cfg.get("reports", [])
    reports.insert(0, {
        "id": uuid.uuid4().hex[:10],
        "filename": os.path.basename(output_path),
        "path": os.path.abspath(output_path),
        "generated_at": datetime.now().isoformat(timespec="minutes"),
        "size_kb": round(os.path.getsize(output_path) / 1024, 1) if os.path.exists(output_path) else 0,
        "emailed": send_email,
    })
    cfg["reports"] = reports[:30]
    save_config(cfg)


# ---------------------------------------------------------------------------
# Disk cleanup -- deletes generated reports/charts/uploads older than
# FILE_RETENTION_HOURS so the disk doesn't fill up on a long-running server.
# ---------------------------------------------------------------------------

def cleanup_old_files():
    cutoff = time.time() - FILE_RETENTION_HOURS * 3600
    removed = 0
    for folder, patterns in [
        ("output", ("report_",)),
        (UPLOAD_DIR, ("",)),  # every upload except the "latest*" file the scheduler needs
        (os.path.join("output", "charts"), ("",)),
    ]:
        if not os.path.isdir(folder):
            continue
        for name in os.listdir(folder):
            if name.startswith("latest"):
                continue  # keep the file the scheduler relies on
            path = os.path.join(folder, name)
            if not os.path.isfile(path):
                continue
            if any(name.startswith(p) for p in patterns) or patterns == ("",):
                try:
                    if os.path.getmtime(path) < cutoff:
                        os.remove(path)
                        removed += 1
                except OSError:
                    pass
    if removed:
        print(f"Cleanup: removed {removed} file(s) older than {FILE_RETENTION_HOURS}h.")

    # Also drop history entries whose files are gone, so "Past Reports" doesn't
    # show dead links.
    cfg = load_config()
    reports = cfg.get("reports", [])
    kept = [r for r in reports if os.path.exists(r.get("path", ""))]
    if len(kept) != len(reports):
        cfg["reports"] = kept
        save_config(cfg)


# ---------------------------------------------------------------------------
# Pipeline runner -- shared by manual uploads and the scheduler
# ---------------------------------------------------------------------------

def run_job(job_id, file_path, send_email, recipients_override=None):
    def set_step(step, **extra):
        JOBS[job_id] = {**JOBS.get(job_id, {}), "status": "running", "step": step, **extra}

    try:
        set_step("reading")
        raw_df = ingest.from_any_file(file_path)

        set_step("cleaning")
        df = process.clean_data(raw_df)
        kpis = process.compute_kpis(df)

        # Data preview -- shown in the UI so the person can see at a glance
        # what was detected, without waiting for the full deck to build.
        preview = {
            "detected_columns": list(df.columns),
            "row_count": len(df),
            "sample_rows": df.head(3).astype(str).to_dict(orient="records"),
        }
        set_step("cleaning", preview=preview)

        set_step("charting", preview=preview)
        chart_paths = {}
        region_df = process.region_summary(df)
        if not region_df.empty:
            chart_paths["region"] = charts.region_bar_chart(region_df)
        trend_df = process.trend_summary(df)
        if not trend_df.empty:
            chart_paths["trend"] = charts.trend_line_chart(trend_df)
        sp_df = process.salesperson_summary(df)
        if not sp_df.empty:
            chart_paths["leaderboard"] = charts.salesperson_leaderboard_chart(sp_df)

        set_step("building", preview=preview)
        insights = process.generate_insights(kpis, region_df, trend_df, sp_df)
        output_path = report_builder.build_report(
            kpis, chart_paths, output_path=f"output/report_{job_id}.pptx",
            region_df=region_df, sp_df=sp_df, insights=insights
        )

        if send_email:
            set_step("emailing", preview=preview)
            send_report(output_path, subject="Business Report", recipients_override=recipients_override)

        JOBS[job_id] = {"status": "done", "step": "done", "output_path": os.path.abspath(output_path), "preview": preview}
        record_report(output_path, send_email)

        cfg = load_config()
        cfg["last_run_at"] = datetime.now().isoformat(timespec="minutes")
        cfg["last_run_status"] = "success"
        save_config(cfg)

    except Exception as e:
        traceback.print_exc()
        JOBS[job_id] = {"status": "error", "step": None, "error": str(e)}
        cfg = load_config()
        cfg["last_run_at"] = datetime.now().isoformat(timespec="minutes")
        cfg["last_run_status"] = f"error: {e}"
        save_config(cfg)


def scheduled_run():
    """Called by APScheduler at the configured daily time."""
    cfg = load_config()
    if not cfg.get("schedule_enabled"):
        return
    latest = cfg.get("latest_file")
    if not latest or not os.path.exists(latest):
        print("Scheduled run skipped -- no file has been uploaded yet.")
        cfg["last_run_at"] = datetime.now().isoformat(timespec="minutes")
        cfg["last_run_status"] = "skipped: no file uploaded yet"
        save_config(cfg)
        return
    job_id = "scheduled-" + uuid.uuid4().hex[:6]
    JOBS[job_id] = {"status": "queued", "step": None}
    run_job(job_id, latest, send_email=True, recipients_override=cfg.get("recipients"))


scheduler = BackgroundScheduler(timezone="Asia/Kolkata")


def reschedule():
    cfg = load_config()
    for job in scheduler.get_jobs():
        if job.id == "daily_report":
            job.remove()
    if cfg.get("schedule_enabled"):
        try:
            hh, mm = cfg.get("schedule_time", "09:00").split(":")
            scheduler.add_job(scheduled_run, CronTrigger(hour=int(hh), minute=int(mm)), id="daily_report")
        except ValueError:
            pass


scheduler.start()
reschedule()
cleanup_old_files()
scheduler.add_job(cleanup_old_files, IntervalTrigger(hours=1), id="cleanup", replace_existing=True)


# ---------------------------------------------------------------------------
# Auth (with basic brute-force rate-limiting)
# ---------------------------------------------------------------------------
LOGIN_MAX_ATTEMPTS = 5
LOGIN_LOCKOUT_MINUTES = 5
_login_attempts = {}  # ip -> {"count": int, "locked_until": datetime|None, "window_start": datetime}


def _login_is_locked(ip):
    entry = _login_attempts.get(ip)
    if not entry or not entry.get("locked_until"):
        return False, None
    if datetime.now() < entry["locked_until"]:
        remaining = int((entry["locked_until"] - datetime.now()).total_seconds() // 60) + 1
        return True, remaining
    _login_attempts.pop(ip, None)
    return False, None


def _record_failed_login(ip):
    entry = _login_attempts.setdefault(ip, {"count": 0, "locked_until": None, "window_start": datetime.now()})
    if datetime.now() - entry["window_start"] > timedelta(minutes=LOGIN_LOCKOUT_MINUTES):
        entry["count"] = 0
        entry["window_start"] = datetime.now()
    entry["count"] += 1
    if entry["count"] >= LOGIN_MAX_ATTEMPTS:
        entry["locked_until"] = datetime.now() + timedelta(minutes=LOGIN_LOCKOUT_MINUTES)


def require_auth(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if APP_PASSWORD and not session.get("authed"):
            if request.path != "/login":
                if request.path.startswith(("/upload", "/status", "/download", "/recipients", "/schedule", "/reports")):
                    return jsonify({"error": "Not authenticated"}), 401
                return redirect(url_for("login"))
        return f(*args, **kwargs)
    return wrapper


@app.route("/login", methods=["GET", "POST"])
def login():
    if not APP_PASSWORD:
        return redirect(url_for("index"))

    ip = request.remote_addr or "unknown"
    error = None

    locked, remaining_minutes = _login_is_locked(ip)
    if locked:
        error = f"Too many attempts. Try again in about {remaining_minutes} minute(s)."
    elif request.method == "POST":
        if request.form.get("password") == APP_PASSWORD:
            session["authed"] = True
            _login_attempts.pop(ip, None)
            return redirect(url_for("index"))
        _record_failed_login(ip)
        still_locked, _ = _login_is_locked(ip)
        error = f"Too many attempts. Try again in about {LOGIN_LOCKOUT_MINUTES} minute(s)." if still_locked else "Incorrect password"

    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.errorhandler(413)
def too_large(e):
    max_mb = app.config["MAX_CONTENT_LENGTH"] // (1024 * 1024)
    return jsonify({"error": f"File too large -- the limit is {max_mb} MB."}), 413


@app.after_request
def add_no_cache_headers(response):
    """
    Prevents the browser from serving a cached copy of a protected page after
    logout (e.g. via the Back button) -- every request to a protected page is
    forced to hit the server, which will bounce to /login if the session is gone.
    """
    if APP_PASSWORD:
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        response.headers["Pragma"] = "no-cache"
    return response


# ---------------------------------------------------------------------------
# Main page + manual generate
# ---------------------------------------------------------------------------

@app.route("/")
@require_auth
def index():
    return render_template("index.html", app_locked=bool(APP_PASSWORD))


@app.route("/upload", methods=["POST"])
@require_auth
def upload():
    if "file" not in request.files or request.files["file"].filename == "":
        return jsonify({"error": "No file selected"}), 400

    f = request.files["file"]
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in (".xlsx", ".xls", ".csv", ".pdf", ".docx"):
        return jsonify({"error": "Please upload a .xlsx, .xls, .csv, .pdf, or .docx file"}), 400

    send_email = request.form.get("send_email", "true") == "true"

    job_id = uuid.uuid4().hex[:10]
    saved_path = os.path.join(UPLOAD_DIR, f"{job_id}{ext}")
    f.save(saved_path)

    # Keep a copy as "the latest file" so the scheduler has something to use
    latest_path = os.path.join(UPLOAD_DIR, f"latest{ext}")
    shutil.copy(saved_path, latest_path)
    cfg = load_config()
    cfg["latest_file"] = latest_path
    save_config(cfg)

    JOBS[job_id] = {"status": "queued", "step": None}
    recipients = load_config().get("recipients")
    thread = threading.Thread(target=run_job, args=(job_id, saved_path, send_email, recipients), daemon=True)
    thread.start()

    return jsonify({"job_id": job_id})


@app.route("/status/<job_id>")
@require_auth
def status(job_id):
    job = JOBS.get(job_id)
    if not job:
        return jsonify({"status": "error", "error": "Unknown job"}), 404

    response = {"status": job["status"], "step": job.get("step")}
    if job.get("preview"):
        response["preview"] = job["preview"]
    if job["status"] == "error":
        response["error"] = job.get("error")
    if job["status"] == "done":
        response["download_url"] = url_for("download", job_id=job_id)
    return jsonify(response)


@app.route("/download/<job_id>")
@require_auth
def download(job_id):
    job = JOBS.get(job_id)
    if not job or job.get("status") != "done":
        return "Report not ready", 404
    return send_file(job["output_path"], as_attachment=True,
                      download_name="Business_Report.pptx")


# ---------------------------------------------------------------------------
# Report history
# ---------------------------------------------------------------------------

@app.route("/reports")
@require_auth
def list_reports():
    cfg = load_config()
    reports = []
    for r in cfg.get("reports", []):
        reports.append({
            "id": r["id"],
            "filename": r["filename"],
            "generated_at": r["generated_at"],
            "size_kb": r["size_kb"],
            "emailed": r.get("emailed", False),
            "available": os.path.exists(r.get("path", "")),
        })
    return jsonify({"reports": reports, "retention_hours": FILE_RETENTION_HOURS})


@app.route("/reports/<report_id>/download")
@require_auth
def download_report(report_id):
    cfg = load_config()
    match = next((r for r in cfg.get("reports", []) if r["id"] == report_id), None)
    if not match or not os.path.exists(match.get("path", "")):
        return "This report has expired or is no longer available.", 404
    return send_file(match["path"], as_attachment=True, download_name=match["filename"])


# ---------------------------------------------------------------------------
# Recipients management
# ---------------------------------------------------------------------------

@app.route("/recipients", methods=["GET"])
@require_auth
def get_recipients():
    return jsonify({"recipients": load_config().get("recipients", [])})


@app.route("/recipients", methods=["POST"])
@require_auth
def update_recipients():
    data = request.get_json(force=True, silent=True) or {}
    action = data.get("action")

    cfg = load_config()
    recipients = cfg.get("recipients", [])

    if action == "add":
        email = (data.get("email") or "").strip()
        if not email or "@" not in email:
            return jsonify({"error": "Enter a valid email address"}), 400
        if email not in recipients:
            recipients.append(email)

    elif action == "remove":
        email = (data.get("email") or "").strip()
        recipients = [r for r in recipients if r != email]

    elif action == "bulk_add":
        # Accepts emails separated by comma, semicolon, newline, or whitespace --
        # handles pasting a list copied from Excel, Outlook, or anywhere else.
        raw_text = data.get("text") or ""
        candidates = re.split(r"[,;\n\r\t ]+", raw_text)
        added, skipped = [], []
        for candidate in candidates:
            candidate = candidate.strip()
            if not candidate:
                continue
            if "@" in candidate and "." in candidate.split("@")[-1]:
                if candidate not in recipients:
                    recipients.append(candidate)
                    added.append(candidate)
            else:
                skipped.append(candidate)
        cfg["recipients"] = recipients
        save_config(cfg)
        return jsonify({"recipients": recipients, "added": added, "skipped": skipped})

    elif action == "clear_all":
        recipients = []

    else:
        return jsonify({"error": "Unknown action"}), 400

    cfg["recipients"] = recipients
    save_config(cfg)
    return jsonify({"recipients": recipients})


# ---------------------------------------------------------------------------
# Schedule management
# ---------------------------------------------------------------------------

@app.route("/schedule", methods=["GET"])
@require_auth
def get_schedule():
    cfg = load_config()
    next_run = None
    job = scheduler.get_job("daily_report")
    if job and job.next_run_time:
        next_run = job.next_run_time.strftime("%d %b %Y, %I:%M %p")
    return jsonify({
        "enabled": cfg.get("schedule_enabled", False),
        "time": cfg.get("schedule_time", "09:00"),
        "has_source_file": bool(cfg.get("latest_file") and os.path.exists(cfg.get("latest_file") or "")),
        "next_run": next_run,
        "last_run_at": cfg.get("last_run_at"),
        "last_run_status": cfg.get("last_run_status"),
    })


@app.route("/schedule", methods=["POST"])
@require_auth
def update_schedule():
    data = request.get_json(force=True, silent=True) or {}
    cfg = load_config()

    if "enabled" in data:
        cfg["schedule_enabled"] = bool(data["enabled"])
    if "time" in data:
        time_str = str(data["time"])
        try:
            hh, mm = time_str.split(":")
            int(hh); int(mm)
            cfg["schedule_time"] = time_str
        except ValueError:
            return jsonify({"error": "Time must be in HH:MM format"}), 400

    save_config(cfg)
    reschedule()

    job = scheduler.get_job("daily_report")
    next_run = job.next_run_time.strftime("%d %b %Y, %I:%M %p") if job and job.next_run_time else None
    return jsonify({"enabled": cfg["schedule_enabled"], "time": cfg["schedule_time"], "next_run": next_run})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
