"""
app.py
Browser-based version of the pipeline. Open a URL, drag in your Excel/CSV,
click one button -- the report is built and (optionally) emailed. No command
line needed once this is running.

Run locally:   python app.py            -> open http://127.0.0.1:5000
Host it for good (no CLI ever again):   see README.md "Deploy as a web app"
"""

import os
import sys
import uuid
import threading
import traceback
from functools import wraps

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from flask import Flask, request, jsonify, render_template, send_file, url_for, session, redirect
from dotenv import load_dotenv

from src import ingest, process, charts, report_builder
from src.email_sender import send_report

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", os.urandom(24).hex())
APP_PASSWORD = os.getenv("APP_PASSWORD")  # set this to lock the page down; leave unset for local-only use

UPLOAD_DIR = os.path.join("output", "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

# In-memory job tracking -- fine for a single-user internal tool.
# Each job: {status, step, output_path, error}
JOBS = {}

STEPS = ["reading", "cleaning", "charting", "building", "emailing", "done"]
STEP_LABELS = {
    "reading": "Reading your file",
    "cleaning": "Cleaning data & detecting columns",
    "charting": "Building charts",
    "building": "Assembling the PowerPoint",
    "emailing": "Sending email",
    "done": "Done",
}


def run_job(job_id, file_path, send_email):
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
        output_path = report_builder.build_report(
            kpis, chart_paths, output_path=f"output/report_{job_id}.pptx"
        )

        if send_email:
            set_step("emailing")
            send_report(output_path, subject="Business Report")

        JOBS[job_id] = {"status": "done", "step": "done", "output_path": output_path}

    except Exception as e:
        traceback.print_exc()
        JOBS[job_id] = {"status": "error", "step": None, "error": str(e)}


def require_auth(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if APP_PASSWORD and not session.get("authed"):
            if request.path.startswith("/upload") or request.path.startswith("/status") or request.path.startswith("/download"):
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


@app.route("/")
@require_auth
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
@require_auth
def upload():
    if "file" not in request.files or request.files["file"].filename == "":
        return jsonify({"error": "No file selected"}), 400

    f = request.files["file"]
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in (".xlsx", ".xls", ".csv"):
        return jsonify({"error": "Please upload a .xlsx, .xls, or .csv file"}), 400

    send_email = request.form.get("send_email", "true") == "true"

    job_id = uuid.uuid4().hex[:10]
    saved_path = os.path.join(UPLOAD_DIR, f"{job_id}{ext}")
    f.save(saved_path)

    JOBS[job_id] = {"status": "queued", "step": None}
    thread = threading.Thread(target=run_job, args=(job_id, saved_path, send_email), daemon=True)
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


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
