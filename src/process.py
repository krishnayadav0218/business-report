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


def _parse_dates_robust(series: pd.Series) -> pd.Series:
    """
    Tries several date conventions and keeps whichever parses the most values --
    real-world files mix DD/MM/YYYY, MM/DD/YYYY, "Jun 2026", Excel serial numbers,
    inconsistent formats row-to-row, etc. Falls back gracefully instead of
    silently dropping everything.

    dayfirst=True is applied wherever possible (including the "mixed" attempt)
    since this pipeline's audience mostly uses DD/MM/YYYY -- e.g. "01/06/2026"
    is treated as 1 June, not 6 January, unless that's genuinely unparseable.
    """
    candidates = []

    # Handles rows that are consistently formatted (single format works for the whole column)
    candidates.append(pd.to_datetime(series, dayfirst=True, errors="coerce"))
    # Handles rows with inconsistent/mixed formatting -- still day-first for ambiguous numeric dates
    try:
        candidates.append(pd.to_datetime(series, format="mixed", dayfirst=True, errors="coerce"))
    except (TypeError, ValueError):
        pass
    # US convention, only as a last resort if the above genuinely can't parse something
    candidates.append(pd.to_datetime(series, dayfirst=False, errors="coerce"))
    try:
        candidates.append(pd.to_datetime(series, format="mixed", dayfirst=False, errors="coerce"))
    except (TypeError, ValueError):
        pass
    # Excel serial date numbers stored as plain numbers/text (e.g. 45810)
    try:
        numeric = pd.to_numeric(series, errors="coerce")
        if numeric.notna().sum() > 0:
            candidates.append(pd.to_datetime(numeric, unit="D", origin="1899-12-30", errors="coerce"))
    except (TypeError, ValueError):
        pass

    best = max(candidates, key=lambda s: s.notna().sum())
    return best


def _clean_currency(value) -> float:
    """Turn '₹1,20,000' or '1,20,000' or 95200 into a plain float."""
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    cleaned = re.sub(r"[₹$€,\s]", "", str(value))
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
        raw_dates = df["Date"]
        df["Date"] = _parse_dates_robust(raw_dates)
        if df["Date"].notna().sum() == 0:
            sample = [str(v) for v in raw_dates.dropna().head(5).tolist()]
            raise ValueError(
                f"Found a Date column, but none of its values could be parsed as dates. "
                f"Example values seen: {sample}. Try a standard format like DD/MM/YYYY."
            )

    df = df.dropna(subset=[c for c in ["Date"] if c in df.columns])
    if df.empty:
        raise ValueError(
            "After cleaning, no valid data rows remained. Check that the file has at least "
            "one row with both a valid date and a collection/amount value."
        )
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
    period_start = period_end = None
    if "Date" in df.columns and df["Date"].notna().any():
        period_start = df["Date"].min()
        period_end = df["Date"].max()
        if df["Date"].nunique() > 1:
            midpoint = period_start + (period_end - period_start) / 2
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
        "period_start": period_start,
        "period_end": period_end,
        "record_count": len(df),
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


def generate_insights(kpis: dict, region_df: pd.DataFrame, trend_df: pd.DataFrame,
                       sp_df: pd.DataFrame) -> dict:
    """
    Turns the raw numbers into plain-English bullet points for the report --
    the narrative content that makes a deck meeting-ready instead of just charts.
    Every insight is derived directly from the data; nothing is invented.
    """
    achievement = kpis["achievement_pct"]

    if achievement >= 100:
        headline = f"Target exceeded — {achievement:.1f}% of target achieved"
    elif achievement >= 90:
        headline = f"Strong performance — {achievement:.1f}% of target achieved"
    elif achievement >= 75:
        headline = f"On track — {achievement:.1f}% of target achieved"
    else:
        headline = f"Below target — {achievement:.1f}% of target achieved"

    # --- Key highlights (executive summary bullets) ---
    highlights = []
    if kpis["total_target"]:
        highlights.append(
            f"Total collection reached ₹{kpis['total_collection']:,.0f} against a target of "
            f"₹{kpis['total_target']:,.0f} ({achievement:.1f}%)."
        )
    else:
        highlights.append(f"Total collection for the period was ₹{kpis['total_collection']:,.0f}.")

    if kpis.get("total_cost"):
        if kpis["profit"] >= 0:
            highlights.append(f"The business remained profitable with a net profit of ₹{kpis['profit']:,.0f}.")
        else:
            highlights.append(f"Costs exceeded collections, resulting in a shortfall of ₹{abs(kpis['profit']):,.0f}.")

    if kpis.get("wow_growth") is not None:
        direction = "increased" if kpis["wow_growth"] >= 0 else "declined"
        highlights.append(f"Collection {direction} by {abs(kpis['wow_growth']):.1f}% compared to the earlier half of the period.")

    if not region_df.empty:
        top_region = region_df.iloc[0]
        highlights.append(f"{top_region['Region']} was the top-performing region, contributing ₹{top_region['Collection']:,.0f}.")

    if not sp_df.empty:
        top_sp = sp_df.iloc[0]
        highlights.append(f"{top_sp['Salesperson']} was the top individual performer, contributing ₹{top_sp['Collection']:,.0f}.")

    if kpis.get("total_orders"):
        highlights.append(f"A total of {kpis['total_orders']:,} orders were processed, averaging ₹{kpis['aov']:,.0f} per order.")

    # --- Region-chart-slide insights ---
    region_insights = []
    if not region_df.empty:
        top = region_df.iloc[0]
        region_insights.append(f"{top['Region']} leads with ₹{top['Collection']:,.0f} in collections.")
        if len(region_df) > 1:
            bottom = region_df.iloc[-1]
            region_insights.append(f"{bottom['Region']} recorded the lowest collections at ₹{bottom['Collection']:,.0f}.")
        if "Target" in region_df.columns:
            achieved = region_df.assign(pct=region_df["Collection"] / region_df["Target"] * 100)
            below = achieved[achieved["pct"] < 100]
            if not below.empty:
                names = ", ".join(below["Region"].tolist())
                region_insights.append(f"{names} fell short of target and may need additional focus.")
            else:
                region_insights.append("Every region met or exceeded its target this period.")

    # --- Trend-chart-slide insights ---
    trend_insights = []
    if not trend_df.empty and len(trend_df) > 1:
        first_val = trend_df.iloc[0]["Collection"]
        last_val = trend_df.iloc[-1]["Collection"]
        change = ((last_val - first_val) / first_val * 100) if first_val else 0
        direction = "risen" if change >= 0 else "fallen"
        trend_insights.append(f"Daily collection has {direction} {abs(change):.1f}% from the start to the end of the period.")
        peak_row = trend_df.loc[trend_df["Collection"].idxmax()]
        peak_date = peak_row["Date"].strftime("%d %b") if hasattr(peak_row["Date"], "strftime") else str(peak_row["Date"])
        trend_insights.append(f"Peak collection was on {peak_date}, at ₹{peak_row['Collection']:,.0f}.")
        avg_daily = trend_df["Collection"].mean()
        trend_insights.append(f"Average daily collection across the period was ₹{avg_daily:,.0f}.")

    # --- Salesperson-chart-slide insights ---
    sp_insights = []
    if not sp_df.empty:
        top = sp_df.iloc[0]
        sp_insights.append(f"{top['Salesperson']} led the team with ₹{top['Collection']:,.0f} in collections.")
        if len(sp_df) > 1:
            avg = sp_df["Collection"].mean()
            below_avg = sp_df[sp_df["Collection"] < avg]
            if not below_avg.empty:
                sp_insights.append(f"{len(below_avg)} of {len(sp_df)} team members are below the team average of ₹{avg:,.0f}.")
            gap = top["Collection"] - sp_df.iloc[-1]["Collection"]
            sp_insights.append(f"The gap between the top and lowest performer is ₹{gap:,.0f}.")

    # --- Recommendations ---
    # Each angle below contributes one recommendation whenever it applies, so a
    # typical report ends up with 4-6 substantive items -- not just a single
    # line -- regardless of whether the period was strong or weak.
    recommendations = []

    # 1. Target achievement
    if kpis.get("total_target"):
        if achievement < 90:
            recommendations.append(
                f"We're at {achievement:.1f}% of target — follow up closely on pending collections "
                f"across every region this week to close the remaining gap."
            )
        elif achievement < 100:
            recommendations.append(
                f"We're close to target at {achievement:.1f}% — a final push on outstanding collections "
                f"should close the gap before period close."
            )
        else:
            recommendations.append(
                f"Target has been met ({achievement:.1f}%) — use this momentum to set a slightly higher "
                f"target next period and lock in whatever changed to get here."
            )
    else:
        recommendations.append("Set a formal collection target for the next period so progress is easy to track and celebrate.")

    # 2. Profitability
    if kpis.get("total_cost"):
        if kpis["profit"] < 0:
            recommendations.append(
                f"Costs currently exceed collections by ₹{abs(kpis['profit']):,.0f} — review the biggest cost "
                f"drivers this month before they compound into next period."
            )
        else:
            recommendations.append(
                f"The period closed profitable at ₹{kpis['profit']:,.0f} — maintain the current cost discipline "
                f"and consider reinvesting part of this margin into growth."
            )
    else:
        recommendations.append("Start tracking costs alongside collections so profitability can be measured going forward.")

    # 3. Trend / growth
    if kpis.get("wow_growth") is not None:
        if kpis["wow_growth"] < 0:
            recommendations.append(
                f"Collection declined {abs(kpis['wow_growth']):.1f}% across the period — a targeted promotion "
                f"or renewed outreach in the second half would help reverse this."
            )
        else:
            recommendations.append(
                f"Collection grew {kpis['wow_growth']:.1f}% across the period — identify what specifically "
                f"drove this (a campaign, a region, a team member) and repeat it next period."
            )

    # 4. Regional focus
    if not region_df.empty:
        if "Target" in region_df.columns:
            achieved = region_df.assign(pct=region_df["Collection"] / region_df["Target"] * 100)
            below = achieved[achieved["pct"] < 100]
            if not below.empty:
                recommendations.append(
                    f"Prioritize support for underperforming regions this period: {', '.join(below['Region'].tolist())} "
                    f"— a short review with the regional leads should surface the blockers."
                )
            else:
                recommendations.append("Every region met its target this period — consider raising targets modestly for the strongest region(s) next time.")
        else:
            recommendations.append("Track regional collection trends period over period to spot which markets are gaining or losing momentum.")

    # 5. Team performance
    if not sp_df.empty and len(sp_df) > 1:
        top_name = sp_df.iloc[0]["Salesperson"]
        recommendations.append(
            f"Recognize {top_name} publicly for leading the team this period, and set up a short session where "
            f"they share their approach with the rest of the team."
        )

    # 6. Always end with a forward-looking action
    recommendations.append("Set next period's targets using this period's actual numbers as the baseline, so goals stay realistic and motivating.")

    return {
        "headline": headline,
        "highlights": highlights,
        "region": region_insights,
        "trend": trend_insights,
        "salesperson": sp_insights,
        "recommendations": recommendations,
    }
