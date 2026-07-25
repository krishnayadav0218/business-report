"""
main.py
Orchestrates the full pipeline: ingest -> clean -> analyze -> chart -> report -> (optional) email.

Run locally:      python src/main.py
Run without email: python src/main.py --no-email
"""

import sys
import argparse
from dotenv import load_dotenv

sys.path.append(".")  # allow running as `python src/main.py` from repo root

from src import ingest, process, charts, report_builder
from src.email_sender import send_report


def run_pipeline(send_email: bool = True):
    load_dotenv()

    print("1/5  Ingesting data...")
    raw_df = ingest.load_data()

    print("2/5  Cleaning + processing data...")
    df = process.clean_data(raw_df)
    kpis = process.compute_kpis(df)

    print("3/5  Generating charts...")
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

    print("4/5  Building PPTX report...")
    output_path = report_builder.build_report(kpis, chart_paths)
    print(f"     Saved: {output_path}")

    if send_email:
        print("5/5  Emailing report...")
        try:
            send_report(output_path, subject="Weekly Business Report")
        except RuntimeError as e:
            print(f"     Skipped email: {e}")
    else:
        print("5/5  Skipping email (--no-email)")

    print("\nDone. KPIs:")
    for k, v in kpis.items():
        print(f"  {k}: {v}")

    return output_path


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--no-email", action="store_true", help="Generate the report but skip sending it")
    args = parser.parse_args()
    run_pipeline(send_email=not args.no_email)
