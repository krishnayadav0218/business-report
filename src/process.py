"""
process.py
Cleans raw data and computes the metrics that go into the report.
Handles Indian-formatted numbers (₹1,20,000), messy dates, and duplicates.

AUTOMATIC COLUMN DETECTION:
You don't need to edit anything for most Excel files. clean_data() scans your
column headers and matches them by keyword (e.g. a column called "Sale Amount"
or "Revenue" or "Amount Received" is auto-detected as Collection).

If your file uses something the keyword list doesn't catch, add an override in
COLUMN_MAP below (left side = internal name, right side = your exact header) --
overrides always win over auto-detection.
"""

import re
import pandas as pd

# Optional manual overrides. Leave as-is to rely on auto-detection.
# Only fill in a line if auto-detection guesses wrong for your file, e.g.:
#     "Collection": "Net Amount",
COLUMN_MAP = {
    "Date": "Date",
    "Region": "Region",
    "Salesperson": "Salesperson",
    "Product": "Product",
    "Target": "Target",
    "Collection": "Collection",
    "Orders": "Orders",
    "Cost": "Cost",
}

# Keyword bank used for auto-detection. Order within a list doesn't matter;
# order of FIELD_PRIORITY below does (checked first = claims a column first).
FIELD_KEYWORDS = {
    "Date": ["date", "dt", "period", "day"],
    "Target": ["target", "goal", "quota", "budgeted"],
    "Cost": ["cost", "expense", "expenditure", "cogs"],
    "Orders": ["order", "qty", "quantity", "units", "no. of", "count"],
    "Region": ["region", "zone", "area", "location", "branch", "state", "city"],
    "Salesperson": ["salesperson", "sales person", "executive", "agent", "employee",
                     "rep", "staff", "owner", "assigned to"],
    "Product": ["product", "item", "sku", "category", "service"],
    "Collection": ["collection", "collected", "revenue", "received", "sales amount",
                    "sales value", "total sales", "payment", "paid", "amount", "net sales", "turnover"],
}

# Checked in this order so specific fields (Target/Cost/Orders) claim their column
# before the broad "amount"-style keywords in Collection grab it by mistake.
FIELD_PRIORITY = ["Date", "Target", "Cost", "Orders", "Region", "Salesperson", "Product", "Collection"]


def _normalize(header: str) -> str:
    return re.sub(r"[^a-z0-9. ]", "", str(header).lower()).strip()


def _auto_detect_field(field: str, candidate_columns: list) -> str:
    keywords = FIELD_KEYWORDS.get(field, [])
    for col in candidate_columns:
        norm = _normalize(col)
        if any(kw in norm for kw in keywords):
            return col
    return None


def apply_column_map(df: pd.DataFrame) -> pd.DataFrame:
    """
    Figure out which of the user's actual columns correspond to the fields the
    pipeline needs, using (in order): manual COLUMN_MAP overrides -> exact name
    match -> keyword auto-detection.
    """
    original_columns = list(df.columns)
    resolved = {}
    remaining = list(df.columns)

    # 1. Manual overrides (only count if the user actually changed them AND the column exists)
    for field, header in COLUMN_MAP.items():
        if header != field and header in df.columns and header in remaining:
            resolved[field] = header
            remaining.remove(header)

    # 2. Exact name match (covers files that already use our internal names)
    for field in FIELD_PRIORITY:
        if field in resolved:
            continue
        if field in df.columns and field in remaining:
            resolved[field] = field
            remaining.remove(field)

    # 3. Keyword auto-detection on whatever's left
    for field in FIELD_PRIORITY:
        if field in resolved:
            continue
        match = _auto_detect_field(field, remaining)
        if match:
            resolved[field] = match
            remaining.remove(match)

    reverse_map = {v: k for k, v in resolved.items()}
    df = df.rename(columns=reverse_map)

    missing = [f for f in ["Date", "Collection"] if f not in df.columns]
    if missing:
        raise ValueError(
            f"Could not auto-detect required column(s): {missing}. "
            f"Your file's columns are: {original_columns}. "
            f"Add an override in COLUMN_MAP in src/process.py, e.g. \"Collection\": \"<your exact header>\"."
        )
    return df


def _clean_currency(value) -> float:
    """Turn '₹1,20,000' or '1,20,000' or 95200 into a plain float."""
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    cleaned = re.sub(r"[₹,\s]", "", str(value))
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """Standardise dtypes, drop duplicates, coerce currency and date columns."""
    df = apply_column_map(df)
    df = df.copy()
    df = df.drop_duplicates()

    currency_cols = ["Target", "Collection", "Cost"]
    for col in currency_cols:
        if col in df.columns:
            df[col] = df[col].apply(_clean_currency)

    if "Orders" in df.columns:
        df["Orders"] = pd.to_numeric(df["Orders"], errors="coerce").fillna(0).astype(int)

    if "Date" in df.columns:
        # dayfirst=True because Indian data is usually DD/MM/YYYY
        df["Date"] = pd.to_datetime(df["Date"], dayfirst=True, errors="coerce")

    df = df.dropna(subset=[c for c in ["Date"] if c in df.columns])
    return df.sort_values("Date") if "Date" in df.columns else df


def compute_kpis(df: pd.DataFrame) -> dict:
    """Return the headline numbers used on the report's summary slide."""
    total_target = df["Target"].sum() if "Target" in df.columns else 0
    total_collection = df["Collection"].sum()
    total_cost = df["Cost"].sum() if "Cost" in df.columns else 0
    total_orders = df["Orders"].sum() if "Orders" in df.columns else 0

    profit = total_collection - total_cost
    achievement_pct = (total_collection / total_target * 100) if total_target else 0
    aov = (total_collection / total_orders) if total_orders else 0

    wow_growth = None
    if "Date" in df.columns and df["Date"].nunique() > 1:
        midpoint = df["Date"].min() + (df["Date"].max() - df["Date"].min()) / 2
        first_half = df[df["Date"] <= midpoint]["Collection"].sum()
        second_half = df[df["Date"] > midpoint]["Collection"].sum()
        if first_half:
            wow_growth = (second_half - first_half) / first_half * 100

    return {
        "total_target": total_target,
        "total_collection": total_collection,
        "total_cost": total_cost,
        "total_orders": int(total_orders),
        "profit": profit,
        "achievement_pct": achievement_pct,
        "aov": aov,
        "wow_growth": wow_growth,
    }


def region_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Collection and target grouped by region -- feeds the regional bar chart."""
    if "Region" not in df.columns:
        return pd.DataFrame()
    agg = {"Collection": ("Collection", "sum")}
    if "Target" in df.columns:
        agg["Target"] = ("Target", "sum")
    result = df.groupby("Region", as_index=False).agg(**agg)
    return result.sort_values("Collection", ascending=False)


def trend_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Daily collection total -- feeds the trend line chart."""
    if "Date" not in df.columns:
        return pd.DataFrame()
    return df.groupby("Date", as_index=False).agg(Collection=("Collection", "sum"))


def salesperson_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Collection per salesperson -- feeds the leaderboard chart."""
    if "Salesperson" not in df.columns:
        return pd.DataFrame()
    agg = {"Collection": ("Collection", "sum")}
    if "Orders" in df.columns:
        agg["Orders"] = ("Orders", "sum")
    result = df.groupby("Salesperson", as_index=False).agg(**agg)
    return result.sort_values("Collection", ascending=False)
