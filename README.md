# Automated Business Reporting Pipeline

An end-to-end Python pipeline that pulls raw business data, cleans and analyzes it with
Pandas, builds charts, generates a branded PowerPoint report, and emails it out — on a
weekly schedule, with zero manual work.

Built to demonstrate a real-world **data automation** workflow: ingestion → analytics →
reporting → scheduled delivery.

## Architecture

```
 Data Source          Processing            Reporting            Delivery
┌────────────┐      ┌──────────────┐      ┌───────────────┐    ┌─────────────┐
│ CSV / SQL / │ ---> │ Pandas clean │ ---> │ Matplotlib     │--> │ SMTP email  │
│ Google Sheet│      │ + KPIs       │      │ charts + PPTX  │    │ attachment  │
│ / REST API  │      │              │      │ (python-pptx)  │    │             │
└────────────┘      └──────────────┘      └───────────────┘    └─────────────┘
                                                                        ^
                                                             GitHub Actions cron
                                                             (every Monday 9 AM)
```

## What it does

1. **Ingests** data from a configurable source — CSV out of the box, or Google Sheets /
   PostgreSQL / MySQL / any JSON API by flipping one `.env` variable.
2. **Cleans and analyzes** the data with Pandas: handles Indian-formatted currency
   (`₹1,20,000`), messy dates, and duplicates. Computes total collection, target
   achievement %, profit/loss, average order value, and period-on-period growth.
3. **Builds charts** with Matplotlib — regional performance, collection trend,
   salesperson leaderboard.
4. **Generates a PowerPoint report** with `python-pptx` — a clean, consistently branded
   deck built programmatically (title slide, KPI summary, chart slides).
5. **Emails the report** as an attachment via SMTP.
6. **Runs on a schedule** with a GitHub Actions workflow (`.github/workflows/weekly_report.yml`)
   — no server or manual trigger needed.

## Quick start

```bash
git clone <your-repo-url>
cd automated-business-reporting-pipeline
pip install -r requirements.txt
cp .env.example .env          # fill in your data source + SMTP details
python src/main.py            # generates + emails the report
python src/main.py --no-email # generates the report only
```

The bundled `data/sample_sales_data.csv` lets the pipeline run immediately with zero
configuration — useful for demoing the project.

## Switching data sources

Set `DATA_SOURCE` in `.env`:

| Value    | Needs |
|----------|-------|
| `csv`    | Nothing extra — works out of the box |
| `gsheet` | `pip install gspread oauth2client` + a `service_account.json` shared with the sheet |
| `sql`    | `pip install sqlalchemy psycopg2-binary` (Postgres) or `pymysql` (MySQL) + DB creds in `.env` |
| `api`    | `pip install requests` + the API URL/headers |

`src/ingest.py` is the only file that changes — everything downstream (`process.py`,
`charts.py`, `report_builder.py`) works on a plain DataFrame regardless of source.

## Automating it

The included GitHub Actions workflow runs the pipeline every Monday at 9 AM IST. To use it:

1. Push this repo to GitHub.
2. Add your credentials as **repository secrets** (Settings → Secrets and variables →
   Actions) — `SMTP_USER`, `SMTP_PASSWORD`, `REPORT_RECIPIENTS`, `DATA_SOURCE`, etc.
3. The workflow runs automatically, and also uploads the generated `.pptx` as a
   downloadable workflow artifact even if email isn't configured.

You can trigger it manually anytime from the **Actions** tab (`workflow_dispatch`).

## Project structure

```
├── data/
│   └── sample_sales_data.csv      # demo dataset
├── src/
│   ├── ingest.py                  # data source connectors
│   ├── process.py                 # cleaning + KPI calculations
│   ├── charts.py                  # matplotlib chart generation
│   ├── report_builder.py          # python-pptx report generation
│   ├── email_sender.py            # SMTP delivery
│   └── main.py                    # pipeline orchestrator
├── .github/workflows/
│   └── weekly_report.yml          # scheduled automation
├── output/                        # generated charts + reports (gitignored)
├── .env.example
└── requirements.txt
```

## Tech stack

Python · Pandas · Matplotlib · python-pptx · SMTP · GitHub Actions

## Possible extensions

- PDF output via ReportLab/WeasyPrint as an alternative to PPTX
- Slack delivery instead of / in addition to email
- Anomaly detection on week-on-week metrics (flag unusual dips automatically)
- Multi-tenant config (one pipeline run, several branded reports for different teams)
