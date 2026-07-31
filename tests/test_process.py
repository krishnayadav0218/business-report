"""
test_process.py
Covers src/process.py: automatic column detection, KPI calculations, and
insight/recommendation generation. Several of these tests exist specifically
because they would have caught real bugs seen in production (e.g. the
recommendations list collapsing to a single generic line).
"""

import pandas as pd
import pytest
from src import process


def test_clean_data_maps_and_types_correctly(sample_csv_path):
    from src import ingest
    raw_df = ingest.from_any_file(sample_csv_path)
    df = process.clean_data(raw_df)
    assert "Date" in df.columns
    assert pd.api.types.is_datetime64_any_dtype(df["Date"])
    assert df["Collection"].dtype.kind in "if"  # int or float


def test_auto_detects_differently_named_columns():
    """Real company files rarely use our exact internal column names."""
    df = pd.DataFrame({
        "Sale Date": ["01-06-2026", "02-06-2026"],
        "Branch": ["Mumbai", "Delhi"],
        "Executive Name": ["Rahul", "Sneha"],
        "Amount Received": [48000, 61200],
    })
    cleaned = process.clean_data(df)
    assert set(["Date", "Region", "Salesperson", "Collection"]).issubset(cleaned.columns)


def test_missing_required_columns_raises_clear_error():
    df = pd.DataFrame({"Foo": [1, 2], "Bar": [3, 4]})
    with pytest.raises(ValueError, match="Could not auto-detect"):
        process.clean_data(df)


def test_compute_kpis_basic_math(sample_df):
    kpis = process.compute_kpis(sample_df)
    assert kpis["total_collection"] == sample_df["Collection"].sum()
    assert kpis["total_orders"] == int(sample_df["Orders"].sum())
    assert kpis["profit"] == kpis["total_collection"] - kpis["total_cost"]
    assert kpis["period_start"] is not None
    assert kpis["period_end"] is not None
    assert kpis["record_count"] == len(sample_df)


def test_mixed_date_formats_parsed_dayfirst():
    """DD/MM/YYYY and named-month dates mixed together should both parse,
    with ambiguous numeric dates read as day-first (not US month-first)."""
    df = pd.DataFrame({
        "Date": ["01/06/2026", "15 Jun 2026", "03/06/2026"],
        "Region": ["West", "North", "East"],
        "Collection": [50000, 60000, 45000],
    })
    cleaned = process.clean_data(df)
    assert len(cleaned) == 3
    # 01/06/2026 read day-first should be June 1st, not January 6th
    first_row_date = cleaned.sort_values("Date").iloc[0]["Date"]
    assert first_row_date.month == 6
    assert first_row_date.day == 1


def test_recommendations_are_never_just_one_line(sample_df):
    """Regression test for a real production bug: recommendations used to
    collapse to a single generic line unless a specific condition was met."""
    kpis = process.compute_kpis(sample_df)
    region_df = process.region_summary(sample_df)
    trend_df = process.trend_summary(sample_df)
    sp_df = process.salesperson_summary(sample_df)
    insights = process.generate_insights(kpis, region_df, trend_df, sp_df)
    assert len(insights["recommendations"]) >= 3


def test_recommendations_rich_even_for_perfect_performance():
    """Even a period where every target is exceeded should yield multiple
    recommendations, not just a generic 'all good' fallback."""
    df = pd.DataFrame({
        "Date": ["01/06/2026", "02/06/2026", "03/06/2026"],
        "Region": ["West", "North", "East"],
        "Salesperson": ["A", "B", "C"],
        "Target": [10000, 10000, 10000],
        "Collection": [15000, 16000, 17000],
        "Orders": [10, 12, 14],
        "Cost": [5000, 5000, 5000],
    })
    cleaned = process.clean_data(df)
    kpis = process.compute_kpis(cleaned)
    region_df = process.region_summary(cleaned)
    trend_df = process.trend_summary(cleaned)
    sp_df = process.salesperson_summary(cleaned)
    insights = process.generate_insights(kpis, region_df, trend_df, sp_df)
    assert len(insights["recommendations"]) >= 3


def test_generate_insights_has_all_expected_keys(sample_df):
    kpis = process.compute_kpis(sample_df)
    region_df = process.region_summary(sample_df)
    trend_df = process.trend_summary(sample_df)
    sp_df = process.salesperson_summary(sample_df)
    insights = process.generate_insights(kpis, region_df, trend_df, sp_df)
    for key in ["headline", "highlights", "region", "trend", "salesperson", "recommendations"]:
        assert key in insights
