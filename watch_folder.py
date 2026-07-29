"""
watch_folder.py
Fully automatic "drop and forget" mode.

Leave this running in the background. Every time you add or REPLACE a file in
the incoming/ folder, it automatically:
    1. detects the file is done being saved/copied (waits for its size to stop changing)
    2. reads it (any column names -- auto-detected, see src/process.py)
    3. builds the PPT report
    4. emails it to REPORT_RECIPIENTS

You never have to click anything after this is started -- just overwrite or
add a new file in incoming/ each day.

Run:      python watch_folder.py
Stop:     Ctrl+C
Windows:  double-click "Start Auto-Watch.bat" instead
"""

import os
import sys
import time
from datetime import datetime

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from dotenv import load_dotenv
from src import ingest, process, charts, report_builder
from src.email_sender import send_report

load_dotenv()

INCOMING_DIR = "incoming"
CHECK_INTERVAL_SECONDS = 20   # how often to check the folder for changes
STABLE_WAIT_SECONDS = 2       # how long a file's size must stay unchanged before we trust it's fully saved
VALID_EXT = (".xlsx", ".xls", ".csv", ".pdf", ".docx")


def _log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")


def get_latest_file():
    os.makedirs(INCOMING_DIR, exist_ok=True)
    files = [os.path.join(INCOMING_DIR, f) for f in os.listdir(INCOMING_DIR)
             if f.lower().endswith(VALID_EXT) and not f.startswith("~$")]  # skip Excel lock files
    if not files:
        return None
    return max(files, key=os.path.getmtime)


def is_file_stable(path):
    """Guards against processing a file that's still mid-copy/mid-save."""
    try:
        size1 = os.path.getsize(path)
        time.sleep(STABLE_WAIT_SECONDS)
        size2 = os.path.getsize(path)
        return size1 == size2
    except OSError:
        return False


def process_file(path):
    _log(f"New/updated file detected: {path}")
    try:
        raw_df = ingest.from_any_file(path)
        df = process.clean_data(raw_df)
        kpis = process.compute_kpis(df)

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

        insights = process.generate_insights(kpis, region_df, trend_df, sp_df)
        output_path = report_builder.build_report(kpis, chart_paths, region_df=region_df, sp_df=sp_df, insights=insights)
        _log(f"Report built: {output_path}")

        send_report(output_path, subject="Business Report")
        _log("Emailed successfully. Waiting for next file...")

    except Exception as e:
        _log(f"ERROR while processing {path}: {e}")
        _log("Fix the issue and save/replace the file again -- it will retry automatically.")


def main():
    os.makedirs(INCOMING_DIR, exist_ok=True)
    _log(f"Watching '{INCOMING_DIR}/' every {CHECK_INTERVAL_SECONDS}s.")
    _log("Add or replace your Excel/CSV file there any time. Press Ctrl+C to stop.")

    last_processed_key = None
    while True:
        latest = get_latest_file()
        if latest:
            key = (latest, os.path.getmtime(latest), os.path.getsize(latest))
            if key != last_processed_key:
                if is_file_stable(latest):
                    fresh_key = (latest, os.path.getmtime(latest), os.path.getsize(latest))
                    if fresh_key != last_processed_key:
                        process_file(latest)
                        last_processed_key = fresh_key
        time.sleep(CHECK_INTERVAL_SECONDS)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        _log("Stopped.")
