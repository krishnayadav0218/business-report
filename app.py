"""
app.py
Browser-based version of the pipeline. Open a URL, drag in your Excel/CSV,
click one button -- the report is built and (optionally) emailed. No command
line needed once this is running.

Also supports:
  - Managing recipient emails from the page itself (no .env editing needed)
  - Scheduling a daily automatic run (uses whichever file was uploaded most
    recently) so you can set it and forget it, right from the portal

Run locally:   python app.py            -> open http://127.0.0.1:5000
Deployment:    see README.md "Deploy as a web app" section (Render, no CLI needed)
"""

import os
import re
import sys
import json
import uuid
import shutil
import threading
import traceback
from functools import wraps
from datetime import datetime

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from flask import Flask, request, jsonify, render_template, send_file, url_for, session, redirect
from dotenv import load_dotenv
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger

from src import ingest, process, charts, report_builder
from src.email_sender import send_report

load_dotenv()

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

UPLOAD_DIR = os.path.join("output", "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)
CONFIG_PATH = os.path.join("output", "config.json")

# In-memory job tracking -- fine for a single-user internal tool.
JOBS = {}

STEP_ORDER = ["reading", "cleaning", "charting", "building", "emailing", "done"]


# ---------------------------------------------------------------------------
# Config: recipients + schedule, persisted to a small JSON file so they
# survive across requests (and across restarts, as long as the disk sticks
# around -- on Render's free tier this resets on redeploy, which is fine
# since you'd typically re-check settings after a deploy anyway).
# ---------------------------------------------------------------------------

def default_config():
    env_recipients = [r.strip() for r in os.getenv("REPORT_RECIPIENTS", "").split(",") if r.strip()]
    return {
        "recipients": env_recipients,
        "schedule_enabled": False,
        "schedule_time": "09:00",
        "latest_file": None,
        "last_run_at": None,
        "last_run_status": None,
    }


def load_config():
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH) as f:
                cfg = json.load(f)
            merged = default_config()
            merged.update(cfg)
            return merged
        except (json.JSONDecodeError, OSError):
            pass
    return default_config()


def save_config(cfg):
    os.makedirs(os.path.dirname(CONFIG_PATH), exist_ok=True)
    with open(CONFIG_PATH, "w") as f:
        json.dump(cfg, f, indent=2)


# ---------------------------------------------------------------------------
# Pipeline runner -- shared by manual uploads and the scheduler
# ---------------------------------------------------------------------------

def run_job(job_id, file_path, send_email, recipients_override=None):
    def set_step(step, **extra):
        JOBS[job_id] = {"status": "running", "step": step, **extra}

    try:
        set_step("reading")
        raw_df = ingest.from_any_file(file_path)

        set_step("cleaning")
        df = process.clean_data(raw_df)
        kpis = process.compute_kpis(df)

        set_step("charting")
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

        set_step("building")
        insights = process.generate_insights(kpis, region_df, trend_df, sp_df)
        output_path = report_builder.build_report(
            kpis, chart_paths, output_path=f"output/report_{job_id}.pptx",
            region_df=region_df, sp_df=sp_df, insights=insights
        )

        if send_email:
            set_step("emailing")
            send_report(output_path, subject="Business Report", recipients_override=recipients_override)

        JOBS[job_id] = {"status": "done", "step": "done", "output_path": output_path}
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
        job.remove()
    if cfg.get("schedule_enabled"):
        try:
            hh, mm = cfg.get("schedule_time", "09:00").split(":")
            scheduler.add_job(scheduled_run, CronTrigger(hour=int(hh), minute=int(mm)), id="daily_report")
        except ValueError:
            pass


scheduler.start()
reschedule()


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------

def require_auth(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if APP_PASSWORD and not session.get("authed"):
            if request.path != "/login":
                if request.path.startswith(("/upload", "/status", "/download", "/recipients", "/schedule")):
                    return jsonify({"error": "Not authenticated"}), 401
                return redirect(url_for("login"))
        return f(*args, **kwargs)
    return wrapper


@app.route("/login", methods=["GET", "POST"])
def login():
    if not APP_PASSWORD:
        return redirect(url_for("index"))
    error = None
    if request.method == "POST":
        if request.form.get("password") == APP_PASSWORD:
            session["authed"] = True
            return redirect(url_for("index"))
        error = "Incorrect password"
    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


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
