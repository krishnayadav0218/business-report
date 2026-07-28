"""
ingest.py
Data Ingestion Layer -- pulls raw data from whichever source is configured.

Supported sources (set DATA_SOURCE in .env):
    csv     -> local/remote CSV file                     (works out of the box)
    excel   -> local .xlsx/.xls file                      (works out of the box)
    gsheet  -> Google Sheets (via gspread)                (needs service_account.json)
    sql     -> PostgreSQL / MySQL                         (needs DB creds in .env)

The rest of the pipeline only cares about getting back a pandas DataFrame,
so swapping sources never touches process.py, charts.py, or report_builder.py.

HEADER ROW AUTO-DETECTION:
Real-world company Excel files often have a company name, report title, or
blank rows above the actual column headers -- e.g.:

    Row 0: "ABC Company - Monthly Sales Report"
    Row 1: (blank)
    Row 2: Date | Region | Salesperson | ... | Collection
    Row 3: 01/06/2026 | West | Rohit | ... | 98500

Reading this naively makes pandas treat row 0 as the header, producing
"Unnamed: 0", "Unnamed: 1", etc. from_csv/from_excel below scan the first
few rows and pick whichever one actually looks like a header row.
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


def from_any_file(path: str) -> pd.DataFrame:
    """Auto-detect CSV vs Excel by extension -- used by the drag-and-drop tool."""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xls"):
        return from_excel(path)
    elif ext == ".csv":
        return from_csv(path)
    else:
        raise ValueError(f"Unsupported file type: {ext} (use .xlsx, .xls, or .csv)")


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
