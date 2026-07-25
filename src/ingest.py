"""
ingest.py
Data Ingestion Layer -- pulls raw data from whichever source is configured.

Supported sources (set DATA_SOURCE in .env):
    csv     -> local/remote CSV file                     (works out of the box)
    gsheet  -> Google Sheets (via gspread)                (needs service_account.json)
    sql     -> PostgreSQL / MySQL                         (needs DB creds in .env)

The rest of the pipeline only cares about getting back a pandas DataFrame,
so swapping sources never touches process.py, charts.py, or report_builder.py.
"""

import os
import pandas as pd


def from_csv(path: str) -> pd.DataFrame:
    """Read raw data from a local CSV (or any URL pandas can read)."""
    return pd.read_csv(path)


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
    elif source == "gsheet":
        return from_google_sheets(os.getenv("GOOGLE_SHEET_URL"))
    elif source == "sql":
        return from_sql(os.getenv("SQL_QUERY", "SELECT * FROM sales;"))
    elif source == "api":
        return from_api(os.getenv("API_URL"))
    else:
        raise ValueError(f"Unknown DATA_SOURCE: {source}")
