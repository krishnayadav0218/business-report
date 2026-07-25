"""
charts.py
Generates chart PNGs from processed data and saves them to output/charts/.
Kept as static images (not native pptx charts) so report_builder.py can just
paste them onto slides -- simplest, most reliable path for an automated pipeline.
"""

import os
import matplotlib
matplotlib.use("Agg")  # no display needed on a server / cron job
import matplotlib.pyplot as plt

# Brand palette -- change these to match your company colors
NAVY = "#1B2A4A"
TEAL = "#2E9E8F"
CORAL = "#E4622C"
GREY = "#8A94A6"

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "axes.edgecolor": GREY,
    "axes.labelcolor": NAVY,
    "text.color": NAVY,
    "xtick.color": NAVY,
    "ytick.color": NAVY,
    "axes.spines.top": False,
    "axes.spines.right": False,
})

CHART_DIR = "output/charts"


def _ensure_dir():
    os.makedirs(CHART_DIR, exist_ok=True)


def region_bar_chart(region_df, filename="region_target_vs_collection.png"):
    """Grouped bar chart: Target vs Collection by region."""
    _ensure_dir()
    fig, ax = plt.subplots(figsize=(9, 5), dpi=150)
    x = range(len(region_df))
    width = 0.35

    ax.bar([i - width / 2 for i in x], region_df["Target"], width, label="Target", color=GREY)
    ax.bar([i + width / 2 for i in x], region_df["Collection"], width, label="Collection", color=TEAL)

    ax.set_xticks(list(x))
    ax.set_xticklabels(region_df["Region"])
    ax.set_ylabel("Amount (₹)")
    ax.set_title("Target vs Collection by Region", fontsize=14, fontweight="bold", loc="left")
    ax.legend(frameon=False)
    ax.yaxis.grid(True, color="#E5E8EC", linewidth=0.8)
    ax.set_axisbelow(True)

    path = os.path.join(CHART_DIR, filename)
    fig.tight_layout()
    fig.savefig(path, transparent=True)
    plt.close(fig)
    return path


def trend_line_chart(trend_df, filename="collection_trend.png"):
    """Line chart of daily collection over time."""
    _ensure_dir()
    fig, ax = plt.subplots(figsize=(9, 5), dpi=150)
    ax.plot(trend_df["Date"], trend_df["Collection"], color=CORAL, linewidth=2.5, marker="o", markersize=4)
    ax.fill_between(trend_df["Date"], trend_df["Collection"], color=CORAL, alpha=0.08)

    ax.set_ylabel("Collection (₹)")
    ax.set_title("Daily Collection Trend", fontsize=14, fontweight="bold", loc="left")
    ax.yaxis.grid(True, color="#E5E8EC", linewidth=0.8)
    ax.set_axisbelow(True)
    fig.autofmt_xdate(rotation=30)

    path = os.path.join(CHART_DIR, filename)
    fig.tight_layout()
    fig.savefig(path, transparent=True)
    plt.close(fig)
    return path


def salesperson_leaderboard_chart(sp_df, filename="salesperson_leaderboard.png"):
    """Horizontal bar chart ranking salespeople by collection."""
    _ensure_dir()
    sp_df = sp_df.sort_values("Collection")
    fig, ax = plt.subplots(figsize=(9, 5), dpi=150)
    colors = [TEAL if i == len(sp_df) - 1 else NAVY for i in range(len(sp_df))]
    ax.barh(sp_df["Salesperson"], sp_df["Collection"], color=colors)

    ax.set_xlabel("Collection (₹)")
    ax.set_title("Salesperson Leaderboard", fontsize=14, fontweight="bold", loc="left")
    ax.xaxis.grid(True, color="#E5E8EC", linewidth=0.8)
    ax.set_axisbelow(True)

    path = os.path.join(CHART_DIR, filename)
    fig.tight_layout()
    fig.savefig(path, transparent=True)
    plt.close(fig)
    return path
