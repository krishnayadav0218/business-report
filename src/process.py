"""
process.py
Cleans raw data and computes the metrics that go into the report.
Handles Indian-formatted numbers (₹1,20,000), messy dates, and duplicates.
"""

import re
import pandas as pd


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
    total_target = df["Target"].sum()
    total_collection = df["Collection"].sum()
    total_cost = df["Cost"].sum()
    total_orders = df["Orders"].sum()

    profit = total_collection - total_cost
    achievement_pct = (total_collection / total_target * 100) if total_target else 0
    aov = (total_collection / total_orders) if total_orders else 0

    # Week-on-week growth: split the date range in half and compare
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
    return (
        df.groupby("Region", as_index=False)
        .agg(Target=("Target", "sum"), Collection=("Collection", "sum"))
        .sort_values("Collection", ascending=False)
    )


def trend_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Daily collection total -- feeds the trend line chart."""
    if "Date" not in df.columns:
        return pd.DataFrame()
    return df.groupby("Date", as_index=False).agg(Collection=("Collection", "sum"))


def salesperson_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Collection per salesperson -- feeds the leaderboard chart."""
    if "Salesperson" not in df.columns:
        return pd.DataFrame()
    return (
        df.groupby("Salesperson", as_index=False)
        .agg(Collection=("Collection", "sum"), Orders=("Orders", "sum"))
        .sort_values("Collection", ascending=False)
    )
