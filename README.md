# Automated Business Reporting Pipeline

A Python pipeline that turns a company Excel/CSV file into a branded PowerPoint report
and emails it out. Handles **any Excel layout automatically** (column names don't
need to match anything). Comes as a **browser-based web app** (no command line, ever,
once deployed), plus a fully automatic "drop and forget" folder mode, a desktop tool,
and scheduled cloud automation.

## Four ways to use it

### 1. Web app — browser only, no command line at all (recommended)
Open a URL, upload your Excel, click a button, get the PPT (downloaded and/or emailed).
Manage everything from the page itself:
- **Recipients** — add/remove the emails that get the report, no `.env` editing needed
- **Automatic Daily Sending** — turn on a daily time to auto-send using whichever file
  you uploaded most recently, or trigger a one-off run manually any time

**Try it locally first (needs one-time setup):**
```bash
pip install -r requirements.txt
cp .env.example .env
python app.py
```
Then open `http://127.0.0.1:5000` in your browser.

**Deploy it so you have a permanent URL — entirely through websites, no CLI:**
See "Deploy as a web app" below.

### 2. Fully automatic — "drop and forget"
Start it once, leave it running in the background. Every time you add or **replace**
a file in the `incoming/` folder, it automatically reads it, builds the PPT, and
emails it — no clicking required.

```bash
python watch_folder.py
```
On Windows, double-click **`Start Auto-Watch.bat`** and keep that window open (it can
sit minimized). It waits until the file is fully saved/copied before processing it,
so it won't grab a half-written file.

### 3. One-click desktop tool
```bash
python desktop_app.py
```
Click **"Select Excel File"** or **"Use Latest from incoming/"**, then
**"Generate & Send Report"**. On Windows, double-click `Generate Report.bat` instead.

### 4. Scheduled automation via GitHub Actions
For a cloud-hosted weekly schedule instead of a live web app — see
`.github/workflows/weekly_report.yml`. Add your credentials as GitHub repo secrets
(Settings → Secrets and variables → Actions).

```bash
python src/main.py            # generates + emails the report
python src/main.py --no-email # generates the report only
```

## Deploy as a web app (no command line needed)

1. **Put the code on GitHub without git commands:**
   - Go to github.com → **New repository** → give it a name → Create.
   - Click **Add file → Upload files** → drag your entire project folder's contents in → **Commit changes**.
2. **Host it on Render (free tier):**
   - Go to **render.com** → sign up with your GitHub account.
   - **New +** → **Web Service** → pick your repo.
   - Render auto-detects `Procfile` and `requirements.txt` — defaults are fine.
3. **Add your settings** under the service's **Environment** tab (all through the
   Render website, click "Add Environment Variable" for each):
   `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASSWORD`, `REPORT_RECIPIENTS`,
   and `APP_PASSWORD` (pick any password — this locks the page so strangers who
   find your URL can't use your company's email to send reports).
4. Click **Deploy**. Render gives you a URL like `https://your-app.onrender.com` —
   open it, enter your `APP_PASSWORD`, upload an Excel file, done.

**Note:** Render's free tier sleeps after 15 minutes of no use — the first request
after that takes ~30-50 seconds to wake up, then it's instant. Don't use Vercel for
this: it's built for static sites/serverless functions, not a Python service that
runs for a while generating charts and a PPTX. Render, Railway, or Fly.io fit this
kind of app.

**Important — keep it at 1 worker.** The `Procfile` runs `gunicorn app:app --workers 1`
on purpose: the automatic daily scheduler runs inside the app process, and more than
one worker would each run their own copy of it, sending duplicate emails. Don't
increase the worker count unless you also move scheduling out of the app.

**Gmail SMTP "timed out" / "Network is unreachable" on Render.** Render's free
tier blocks outbound SMTP ports (25, 465, 587) to prevent spam abuse — so plain
SMTP (Gmail, Outlook, etc.) **cannot work there**, no matter how it's configured.
The fix is to send over HTTPS instead, using SendGrid's Web API:

1. Sign up free at **sendgrid.com** (100 emails/day free, no card needed).
2. **Settings → Sender Authentication → Verify a Single Sender** — enter your own
   email, confirm the verification email SendGrid sends you.
3. **Settings → API Keys → Create API Key** (Full Access is fine) — copy the key,
   shown only once.
4. On Render, add these environment variables:
   - `EMAIL_PROVIDER` = `sendgrid`
   - `SENDGRID_API_KEY` = the key from step 3
   - `SENDGRID_FROM_EMAIL` = the email you verified in step 2
5. You can remove `SMTP_*` variables on Render — they're unused once `EMAIL_PROVIDER=sendgrid`.

Locally, or in GitHub Actions, plain SMTP works fine as-is (`EMAIL_PROVIDER=smtp`,
the default) — outbound SMTP isn't blocked there.


## Any Excel layout works automatically

You do **not** need to rename your columns or edit any config. `src/process.py`
scans your file's headers and matches them by keyword — e.g. a column called
`Sale Date`, `Amount Received`, `Executive Name`, or `Branch` is automatically
recognized as `Date`, `Collection`, `Salesperson`, or `Region` respectively.
Only `Date` and `Collection` (or anything that looks like a collection/revenue
column) are required — everything else is optional, and the pipeline just skips
the charts/cards it can't compute if a column is missing.

**If auto-detection ever guesses wrong** for one column, you can force it — open
`src/process.py` and add one line to `COLUMN_MAP`:

```python
COLUMN_MAP = {
    "Collection": "Net Amount",   # <- only add a line for the field that needs fixing
}
```

## Architecture

```
 Excel/CSV file        Processing            Reporting            Delivery
┌────────────┐      ┌──────────────┐      ┌───────────────┐    ┌─────────────┐
│ Selected in │ ---> │ Pandas clean │ ---> │ Matplotlib     │--> │ SMTP email  │
│ desktop app │      │ + KPIs       │      │ charts + PPTX  │    │ attachment  │
│ (or scheduled)     │              │      │ (python-pptx)  │    │             │
└────────────┘      └──────────────┘      └───────────────┘    └─────────────┘
```

## What it computes

Total collection, target achievement %, profit/loss, average order value, and
period-on-period growth — plus three charts: regional performance, daily trend,
and a salesperson leaderboard (any of these are skipped automatically if the
underlying column isn't in your file).

## Project structure

```
├── app.py                          # web app -- browser upload, no CLI needed
├── templates/                      # web app pages (index, login)
├── Procfile                        # for Render/Railway deployment
├── watch_folder.py                 # fully automatic mode -- watches incoming/
├── Start Auto-Watch.bat            # Windows launcher for watch_folder.py
├── desktop_app.py                  # one-click GUI tool (Excel -> PPT -> email)
├── Generate Report.bat             # Windows launcher for desktop_app.py
├── incoming/                       # drop your daily Excel/CSV file here
├── data/
│   └── sample_sales_data.csv       # demo dataset
├── src/
│   ├── ingest.py                   # data source connectors (CSV/Excel/Sheets/SQL/API)
│   ├── process.py                  # auto column detection + cleaning + KPI calculations
│   ├── charts.py                   # matplotlib chart generation
│   ├── report_builder.py           # python-pptx report generation
│   ├── email_sender.py             # SMTP delivery
│   └── main.py                     # scheduled pipeline entry point (GitHub Actions)
├── .github/workflows/
│   └── weekly_report.yml           # optional cloud-scheduled automation
├── output/                         # generated charts + reports (gitignored)
├── .env.example
└── requirements.txt
```

## Testing

A test suite covers the parts of this project most likely to break silently --
file parsing, KPI math, and (critically) the exact kind of signature mismatch
between `app.py` and `report_builder.py` that caused a real production bug
once already.

```bash
pip install -r requirements-dev.txt
pytest
```

Runs automatically on every push/PR via `.github/workflows/tests.yml`. If this
fails, don't deploy that commit — it means something inside the pipeline no
longer matches what another part of the code expects from it.

## Tech stack

Python · Flask · Pandas · Matplotlib · python-pptx · Tkinter · SMTP · GitHub Actions

## Possible extensions

- PDF output via ReportLab/WeasyPrint as an alternative to PPTX
- Slack delivery instead of / in addition to email
- Anomaly detection on period-on-period metrics (flag unusual dips automatically)
- A shareable web upload page (Flask) for teammates, instead of the local desktop tool
