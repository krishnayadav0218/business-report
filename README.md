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
- **Recipients** — add/remove emails one at a time, paste/import a whole list at once, or
  copy the current list out again — no `.env` editing needed
- **Automatic Daily Sending** — turn on a daily time to auto-send using whichever file
  you uploaded most recently, or trigger a one-off run manually any time
- **Data preview** — as soon as your file is read, see the detected columns and row count
  before the full deck finishes building
- **Past Reports** — every generated report stays available to re-download for a while
  (default 24h) without regenerating it
- **Login protection** with rate-limiting (locks out after repeated wrong passwords) when
  `APP_PASSWORD` is set

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

### Making Recipients/Schedule settings survive a redeploy

By default, settings are stored in a local JSON file, which resets whenever Render's
free tier redeploys (its disk isn't persistent). To keep them permanently:

1. Create a free Postgres database at **supabase.com** (no card needed) — or use
   Render's own managed Postgres if you're on a paid plan.
2. Copy the connection string (Supabase: **Project Settings → Database → Connection
   string → URI**).
3. Add it on Render as `DATABASE_URL`.

That's it — the app auto-detects `DATABASE_URL` and switches to it, creating the table
it needs on first run. Leave it unset to keep using the local file (fine for local/
desktop use, or if you don't mind reconfiguring after a redeploy).

### Getting notified when something breaks

Set `SENTRY_DSN` (free tier at **sentry.io**, no card needed) to get an email/alert the
moment a report generation fails in production, instead of finding out when someone
tells you it didn't work.

### Other tunables

- `FILE_RETENTION_HOURS` (default `24`) — how long generated reports/uploads stay on
  disk before automatic hourly cleanup removes them.
- `MAX_UPLOAD_MB` (default `20`) — rejects files larger than this before they reach
  the pipeline.

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

**Reports keep landing in spam.** This happens because `SENDGRID_FROM_EMAIL` is a
`@gmail.com` (or Yahoo/Outlook) address, but the mail is actually sent via SendGrid's
servers, not Google's. Gmail's DMARC policy specifically flags this mismatch as likely
spoofing — this is industry-wide since 2024 (Gmail/Yahoo's bulk-sender rules), not a bug
in this project.

- **Real fix:** authenticate a domain you own in SendGrid (**Settings → Sender
  Authentication → Authenticate Your Domain**, add the DNS records it gives you), then
  send from an address on that domain (e.g. `reports@yourdomain.com`). A domain costs
  roughly ₹100–500/year from any registrar. This is the only way to fully stop the
  spam-folder issue for *new* recipients automatically.
- **Free workaround (per recipient, one-time):** ask each recipient to open the email
  in Spam, click **"Report not spam"**, and add a filter (Settings → Filters →
  From: your sender address → *Never send to Spam*). After that, mail from that sender
  reaches their inbox normally going forward.
- `src/email_sender.py` also sends a proper multipart plain-text + HTML message with a
  real body and a Reply-To header (instead of a bare one-line automated-looking email),
  which helps marginally — but doesn't replace domain authentication as the real fix.



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

## Not just tables — form-style and free-text documents work too

Most Excel/CSV/PDF/Word files are proper tables, and those are handled as described
above. But some files list data as a vertical form instead:

```
Region: West
Date: 01/06/2026
Collection: 48000
```

This is detected automatically and turned into a normal row of data. If a sheet has
several such blocks separated by blank rows (e.g. one block per transaction), each
block becomes its own row. PDFs and Word documents that don't contain a real table at
all get the same "Label: Value" line-matching applied to their text as a last resort,
instead of failing outright.

**One real limitation:** a single free-floating document (one invoice, one letter) can
only ever produce **one row** of data this way, so charts that need multiple rows
(trends, regional comparison) won't have much to show — the summary numbers will still
come through, just without the multi-row charts. For a proper multi-row report, an
Excel/CSV table (or a PDF/Word file with several blocks) works best.

**Scanned/image PDFs** (a photo of a document, no selectable text) aren't supported —
that needs OCR, which isn't included since it requires system-level dependencies that
don't install cleanly on Render's standard Python deploy (would need a Dockerfile-based
deploy instead). If this is a common case for you, ask and this can be added.

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
│   ├── ingest.py                   # multi-format ingestion (CSV/Excel/PDF/Word/Sheets/SQL)
│   ├── process.py                  # auto column detection + cleaning + KPIs + insights
│   ├── charts.py                   # matplotlib chart generation
│   ├── report_builder.py           # python-pptx report generation + speaker notes
│   ├── email_sender.py             # SMTP + SendGrid delivery
│   └── main.py                     # scheduled pipeline entry point (GitHub Actions)
├── tests/                          # pytest suite -- see "Testing" below
├── pytest.ini
├── requirements-dev.txt
├── .github/workflows/
│   ├── weekly_report.yml           # optional cloud-scheduled automation
│   └── tests.yml                   # runs the test suite on every push/PR
├── output/                         # generated charts + reports (gitignored, auto-cleaned)
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

Python · Flask · Pandas · Matplotlib · python-pptx · pdfplumber · python-docx ·
SQLAlchemy (optional) · APScheduler · Tkinter · SMTP/SendGrid · GitHub Actions

## Architecture decisions

A few choices worth knowing about, in case they come up (they make decent
interview talking points too):

- **SMTP → SendGrid HTTP API on Render.** Render's free tier blocks outbound
  SMTP ports to prevent spam abuse, so plain SMTP silently times out there no
  matter how it's configured. `email_sender.py` supports both, switchable via
  `EMAIL_PROVIDER` — SMTP for local/GitHub Actions where it's unrestricted,
  SendGrid's HTTPS API for Render where SMTP can't work at all. Same interface,
  different transport, chosen per environment.
- **Vercel → Render.** The pipeline was first pointed at Vercel, which is built
  for static sites/serverless functions with short execution limits — not a
  fit for a Python service that runs for a while generating charts and a PPTX.
  Render (or Railway/Fly.io) fits a long-running Flask service properly.
- **Header-row auto-detection instead of a fixed `header=0`.** Real company
  spreadsheets almost never start with the header on row 1 (title rows, blank
  spacer rows, merged 2-row headers are all common). `ingest.py` scores each of
  the first ~20 rows by "how much does this look like a row of text labels"
  rather than assuming a fixed structure, and combines two-row merged headers
  automatically.
- **Keyword-based column mapping instead of exact-name matching.** Column
  auto-detection in `process.py` matches by substring keywords ("date", "target",
  "collection", etc.) rather than expecting exact internal names, so files from
  different companies/departments work without configuration.
- **A regression test for the exact bug that happened in production.**
  `test_report_builder.py::test_build_report_accepts_the_kwargs_app_py_actually_passes`
  exists specifically because a signature mismatch between `app.py` and
  `report_builder.py` once caused a live "unexpected keyword argument" error.
  The fix wasn't just patching that bug — it was adding a test that makes that
  whole class of bug impossible to ship again.
- **Optional persistence via `DATABASE_URL` instead of a hard dependency.**
  Settings default to a local JSON file (zero setup, works everywhere) and
  transparently switch to Postgres if `DATABASE_URL` is set, rather than forcing
  every user to stand up a database just to try the app.

## Possible extensions

- Anomaly detection on period-on-period metrics (flag unusual dips automatically)
- Slack delivery instead of / in addition to email
- Company branding customization (logo, colors) from the portal itself
- Direct API integration with accounting tools (Tally, Zoho Books, QuickBooks)
  so reports can generate without a manual upload at all
