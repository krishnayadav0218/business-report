"""
ingest.py
Data Ingestion Layer -- pulls raw data from whichever source is configured.

File types handled by from_any_file() (used by the web app / desktop app / watch folder):
    .csv           -> plain CSV
    .xlsx / .xls   -> Excel
    .pdf           -> PDF containing a real data table (extracted via pdfplumber)
    .docx          -> Word document containing a real table (extracted via python-docx)

Other sources (set DATA_SOURCE in .env, used by src/main.py for scheduled runs):
    gsheet  -> Google Sheets (via gspread)                (needs service_account.json)
    sql     -> PostgreSQL / MySQL                         (needs DB creds in .env)

The rest of the pipeline only cares about getting back a pandas DataFrame,
so swapping sources never touches process.py, charts.py, or report_builder.py.

HEADER ROW AUTO-DETECTION:
Real-world company files often have a company name, report title, or blank
rows/lines above the actual column headers -- e.g.:

    Row 0: "ABC Company - Monthly Sales Report"
    Row 1: (blank)
    Row 2: Date | Region | Salesperson | ... | Collection
    Row 3: 01/06/2026 | West | Rohit | ... | 98500

Reading this naively makes pandas treat row 0 as the header, producing
"Unnamed: 0", "Unnamed: 1", etc. Every loader below scans the first few rows
and picks whichever one actually looks like a header row.
"""

import os
import pandas as pd

MAX_HEADER_SCAN_ROWS = 15


def _detect_header_row(raw_df: pd.DataFrame) -> int:
    """
    Scan the first few rows and return the index of the one most likely to be
    the real header row -- the row with the most non-empty cells (title rows
    and blank spacer rows have far fewer filled cells than a header row).
    """
    scan_limit = min(MAX_HEADER_SCAN_ROWS, len(raw_df))
    best_idx, best_score = 0, -1
    for i in range(scan_limit):
        non_null = raw_df.iloc[i].notna().sum()
        if non_null > best_score:
            best_score = non_null
            best_idx = i
    return best_idx


def _clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Strip whitespace/newlines from headers and drop fully-empty columns."""
    df = df.dropna(axis=1, how="all")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def from_csv(path: str) -> pd.DataFrame:
    """Read raw data from a local CSV (or any URL pandas can read), auto-detecting the header row."""
    raw = pd.read_csv(path, header=None)
    header_row = _detect_header_row(raw)
    df = pd.read_csv(path, header=header_row)
    return _clean_columns(df)


def from_excel(path: str, sheet_name=0) -> pd.DataFrame:
    """Read raw data from a local .xlsx/.xls file, auto-detecting the header row."""
    engine = "openpyxl" if str(path).lower().endswith("x") else None
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None, engine=engine)
    header_row = _detect_header_row(raw)
    df = pd.read_excel(path, sheet_name=sheet_name, header=header_row, engine=engine)
    return _clean_columns(df)


def _rows_to_dataframe(rows: list) -> pd.DataFrame:
    """Turn a list-of-lists (as extracted from a PDF/Word table) into a DataFrame,
    auto-detecting which row is the header the same way Excel/CSV files are handled."""
    if not rows:
        raise ValueError("No table data found")
    raw = pd.DataFrame(rows)
    header_row = _detect_header_row(raw)
    header = raw.iloc[header_row]
    df = raw.iloc[header_row + 1:].copy()
    df.columns = header
    return _clean_columns(df)


def from_pdf(path: str) -> pd.DataFrame:
    """
    Extract tabular data from a PDF. Scans every page for tables and picks the
    largest one -- or, if several pages share the same header row (a table that
    continues across pages), stitches them together into a single DataFrame.
    Works for text-based PDFs with real table structure; scanned/image-only
    PDFs or free-form text reports won't have extractable tables.
    """
    import pdfplumber

    all_tables = []  # list of (page_num, rows)
    with pdfplumber.open(path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            for table in page.extract_tables():
                if table and len(table) >= 2:  # needs at least a header + 1 data row
                    all_tables.append((page_num, table))

    if not all_tables:
        raise ValueError(
            "No tables could be found in this PDF. This works best with PDFs that contain "
            "a real data table (not a scanned image or a plain-text report). "
            "Try exporting the data as Excel/CSV instead."
        )

    # If multiple tables share an identical header row, they're likely the same
    # table continuing across pages -- stitch them together.
    first_header = tuple(str(c) for c in all_tables[0][1][0])
    matching = [rows for _, rows in all_tables if tuple(str(c) for c in rows[0]) == first_header]
    if len(matching) > 1:
        combined = [matching[0][0]]  # one header row
        for rows in matching:
            combined.extend(rows[1:])  # skip repeated header on each page
        return _rows_to_dataframe(combined)

    # Otherwise just use the single largest table found anywhere in the document.
    _, biggest = max(all_tables, key=lambda t: len(t[1]))
    return _rows_to_dataframe(biggest)


def from_docx(path: str) -> pd.DataFrame:
    """
    Extract tabular data from a Word document (.docx). Picks the largest table
    in the document. .doc (old binary Word format) isn't supported -- save as
    .docx first.
    """
    import docx

    document = docx.Document(path)
    if not document.tables:
        raise ValueError(
            "No tables could be found in this Word document. Make sure your data is in a "
            "real Word table (Insert > Table), not just typed/tabbed text."
        )

    biggest_table = max(document.tables, key=lambda t: len(t.rows))
    rows = [[cell.text.strip() for cell in row.cells] for row in biggest_table.rows]
    return _rows_to_dataframe(rows)


def from_any_file(path: str) -> pd.DataFrame:
    """Auto-detect file type by extension -- used by the drag-and-drop tool and web app."""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xls"):
        return from_excel(path)
    elif ext == ".csv":
        return from_csv(path)
    elif ext == ".pdf":
        return from_pdf(path)
    elif ext == ".docx":
        return from_docx(path)
    elif ext == ".doc":
        raise ValueError("Old .doc format isn't supported -- please save/export as .docx and try again.")
    else:
        raise ValueError(f"Unsupported file type: {ext} (use .xlsx, .xls, .csv, .pdf, or .docx)")


def from_google_sheets(sheet_url: str, worksheet_name: str = "Sheet1") -> pd.DataFrame:
    """
    Read a live Google Sheet.
    Needs: pip install gspread oauth2client
           a service_account.json (share the sheet with its client_email)
    """
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials

    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds_path = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "service_account.json")
    creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, scope)
    client = gspread.authorize(creds)

    sheet = client.open_by_url(sheet_url).worksheet(worksheet_name)
    records = sheet.get_all_records()
    return pd.DataFrame(records)


def from_sql(query: str) -> pd.DataFrame:
    """
    Read from PostgreSQL or MySQL depending on DB_ENGINE in .env.
    Needs: pip install sqlalchemy psycopg2-binary   (postgres)
                       pymysql                       (mysql)
    """
    from sqlalchemy import create_engine

    engine_type = os.getenv("DB_ENGINE", "postgresql")  # or "mysql+pymysql"
    user = os.getenv("DB_USER")
    password = os.getenv("DB_PASSWORD")
    host = os.getenv("DB_HOST")
    port = os.getenv("DB_PORT")
    dbname = os.getenv("DB_NAME")

    conn_str = f"{engine_type}://{user}:{password}@{host}:{port}/{dbname}"
    engine = create_engine(conn_str)
    return pd.read_sql(query, engine)


def from_api(url: str, headers: dict | None = None) -> pd.DataFrame:
    """Pull JSON data from a third-party API (e.g. Stripe, GA) and flatten to a DataFrame."""
    import requests

    resp = requests.get(url, headers=headers or {}, timeout=30)
    resp.raise_for_status()
    return pd.json_normalize(resp.json())


def load_data() -> pd.DataFrame:
    """
    Single entry point used by main.py.
    Reads DATA_SOURCE from the environment and dispatches to the right loader.
    Defaults to the bundled sample CSV so the pipeline runs with zero config.
    """
    source = os.getenv("DATA_SOURCE", "csv")

    if source == "csv":
        path = os.getenv("CSV_PATH", "data/sample_sales_data.csv")
        return from_csv(path)
    elif source == "excel":
        path = os.getenv("EXCEL_PATH", "data/sample_sales_data.xlsx")
        return from_excel(path)
    elif source == "gsheet":
        return from_google_sheets(os.getenv("GOOGLE_SHEET_URL"))
    elif source == "sql":
        return from_sql(os.getenv("SQL_QUERY", "SELECT * FROM sales;"))
    elif source == "api":
        return from_api(os.getenv("API_URL"))
    else:
        raise ValueError(f"Unknown DATA_SOURCE: {source}")
