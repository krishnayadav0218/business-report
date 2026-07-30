"""
report_builder.py
Builds a meeting-ready PPTX report: title, executive summary, key highlights,
chart slides with plain-English commentary next to each chart, a data table,
and a recommendations slide -- not just a stack of charts.

Design note: instead of hand-editing a fixed .pptx template's placeholders
(brittle -- breaks the moment a shape gets renamed), this builds each slide
programmatically with python-pptx using fixed style helper functions (colors,
fonts, spacing). Same end result -- a consistent branded deck every run --
without depending on a specific template file surviving edits in PowerPoint.
"""

import os
from datetime import date
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

NAVY = RGBColor(0x1B, 0x2A, 0x4A)
TEAL = RGBColor(0x2E, 0x9E, 0x8F)
CORAL = RGBColor(0xE4, 0x62, 0x2C)
GREY = RGBColor(0x8A, 0x94, 0xA6)
LIGHT_BG = RGBColor(0xF7, 0xF8, 0xFA)
LIGHT_LINE = RGBColor(0xE4, 0xE7, 0xEC)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
MARGIN = Inches(0.6)


# ---------------------------------------------------------------------------
# Low-level helpers
# ---------------------------------------------------------------------------

def _blank_slide(prs, bg=WHITE):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # fully blank layout
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = bg
    return slide


def _textbox(slide, left, top, width, height, text, size=18, bold=False,
             color=NAVY, align=PP_ALIGN.LEFT, font="Calibri", italic=False):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = font
    return box


def _bullet_list(slide, left, top, width, height, items, size=14, color=NAVY,
                  bullet_color=TEAL, line_spacing=1.35, gap_after=10):
    """A clean bullet list -- used for highlights, insights, and recommendations."""
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = line_spacing
        p.space_after = Pt(gap_after)
        r1 = p.add_run()
        r1.text = "\u25cf  "
        r1.font.size = Pt(size)
        r1.font.color.rgb = bullet_color
        r1.font.name = "Calibri"
        r2 = p.add_run()
        r2.text = item
        r2.font.size = Pt(size)
        r2.font.color.rgb = color
        r2.font.name = "Calibri"
    return box


def _slide_header(slide, title, subtitle=None):
    _textbox(slide, MARGIN, Inches(0.45), Inches(11.5), Inches(0.6),
              title, size=26, bold=True, color=NAVY)
    if subtitle:
        _textbox(slide, MARGIN, Inches(1.02), Inches(11.5), Inches(0.4),
                  subtitle, size=13, color=GREY, italic=True)
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, MARGIN, Inches(1.42), Inches(12.13), Pt(1.5))
    line.fill.solid()
    line.fill.fore_color.rgb = LIGHT_LINE
    line.line.fill.background()
    line.shadow.inherit = False


def _kpi_card(slide, left, top, width, height, label, value, accent=TEAL):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.adjustments[0] = 0.06
    card.fill.solid()
    card.fill.fore_color.rgb = LIGHT_BG
    card.line.fill.background()
    card.shadow.inherit = False

    _textbox(slide, left + Inches(0.25), top + Inches(0.2), width - Inches(0.5), Inches(0.4),
              label, size=13, color=GREY)
    _textbox(slide, left + Inches(0.25), top + Inches(0.6), width - Inches(0.5), Inches(0.7),
              value, size=26, bold=True, color=accent)


def _insight_panel(slide, left, top, width, height, panel_title, items):
    """A light card holding a short heading + bullet insights -- sits beside charts."""
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.adjustments[0] = 0.03
    card.fill.solid()
    card.fill.fore_color.rgb = LIGHT_BG
    card.line.fill.background()
    card.shadow.inherit = False

    _textbox(slide, left + Inches(0.3), top + Inches(0.22), width - Inches(0.6), Inches(0.4),
              panel_title, size=14, bold=True, color=NAVY)
    if items:
        _bullet_list(slide, left + Inches(0.3), top + Inches(0.75), width - Inches(0.6),
                     height - Inches(1.0), items, size=12.5, line_spacing=1.3, gap_after=8)
    else:
        _textbox(slide, left + Inches(0.3), top + Inches(0.75), width - Inches(0.6), Inches(0.5),
                  "No notable observations for this section.", size=12, color=GREY, italic=True)


def _footer(slide, page_label):
    _textbox(slide, MARGIN, Inches(7.12), Inches(6), Inches(0.3),
              "Confidential \u2014 for internal use", size=9, color=GREY)
    _textbox(slide, Inches(11.5), Inches(7.12), Inches(1.2), Inches(0.3),
              page_label, size=9, color=GREY, align=PP_ALIGN.RIGHT)


def _chart_with_insights_slide(prs, title, chart_path, insight_title, insights, page_label, subtitle=None):
    slide = _blank_slide(prs)
    _slide_header(slide, title, subtitle)

    chart_left = MARGIN
    chart_top = Inches(1.65)
    chart_width = Inches(7.6)
    slide.shapes.add_picture(chart_path, chart_left, chart_top, width=chart_width)

    panel_left = Inches(8.45)
    panel_width = Inches(4.28)
    _insight_panel(slide, panel_left, chart_top, panel_width, Inches(5.2), insight_title, insights)

    _footer(slide, page_label)
    return slide


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def build_report(kpis: dict, chart_paths: dict, output_path="output/Weekly_Business_Report.pptx",
                  report_title="Business Report", period_label=None,
                  region_df=None, sp_df=None, insights=None):
    """
    kpis         -- dict from process.compute_kpis()
    chart_paths  -- dict with any of: region, trend, leaderboard (paths to PNGs)
    region_df    -- optional DataFrame from process.region_summary() (for the data table)
    sp_df        -- optional DataFrame from process.salesperson_summary() (unused directly,
                    kept for symmetry / future extension)
    insights     -- optional dict from process.generate_insights() -- if omitted, chart
                    slides fall back to full-width charts with no commentary panel
    """
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    period_label = period_label or date.today().strftime("%d %b %Y")
    insights = insights or {}
    page_num = 1

    def next_page():
        nonlocal page_num
        page_num += 1
        return f"{page_num - 1}"

    # ---- Slide 1: Title ----
    s1 = _blank_slide(prs, bg=NAVY)
    _textbox(s1, Inches(0.8), Inches(2.5), Inches(11), Inches(1.2),
              report_title, size=40, bold=True, color=WHITE)
    headline = insights.get("headline")
    if headline:
        _textbox(s1, Inches(0.8), Inches(3.35), Inches(11), Inches(0.5),
                  headline, size=17, color=TEAL, bold=True)
    _textbox(s1, Inches(0.8), Inches(3.95), Inches(11), Inches(0.5),
              f"Reporting period ending {period_label}", size=14, color=RGBColor(0xB8, 0xC2, 0xD6))

    # ---- Slide 2: Executive Summary (KPI cards) ----
    s2 = _blank_slide(prs)
    _slide_header(s2, "Executive Summary", headline)

    card_w, card_h = Inches(2.9), Inches(1.5)
    gap = Inches(0.3)
    top = Inches(1.75)
    cards = [
        ("Total Collection", f"\u20b9{kpis['total_collection']:,.0f}", TEAL),
        ("Target Achievement", f"{kpis['achievement_pct']:.1f}%", NAVY),
        ("Total Orders", f"{kpis['total_orders']:,}", CORAL),
        ("Avg Order Value", f"\u20b9{kpis['aov']:,.0f}", TEAL),
    ]
    left = MARGIN
    for label, value, accent in cards:
        _kpi_card(s2, left, top, card_w, card_h, label, value, accent)
        left = Emu(left + card_w + gap)

    profit_color = TEAL if kpis["profit"] >= 0 else CORAL
    profit_label = "Profit" if kpis["profit"] >= 0 else "Loss"
    _kpi_card(s2, MARGIN, Inches(3.55), card_w, card_h, profit_label, f"\u20b9{abs(kpis['profit']):,.0f}", profit_color)

    if kpis.get("wow_growth") is not None:
        growth_color = TEAL if kpis["wow_growth"] >= 0 else CORAL
        arrow = "\u25b2" if kpis["wow_growth"] >= 0 else "\u25bc"
        _kpi_card(s2, Emu(MARGIN + card_w + gap), Inches(3.55), card_w, card_h,
                  "Period-on-Period Growth", f"{arrow} {abs(kpis['wow_growth']):.1f}%", growth_color)
    _footer(s2, next_page())

    # ---- Slide 3: Key Highlights ----
    if insights.get("highlights"):
        s3 = _blank_slide(prs)
        _slide_header(s3, "Key Highlights", "What stands out in this period's numbers")
        _bullet_list(s3, MARGIN, Inches(1.9), Inches(12.1), Inches(4.8),
                     insights["highlights"], size=16, line_spacing=1.5, gap_after=16)
        _footer(s3, next_page())

    # ---- Slide 4: Regional Performance ----
    if chart_paths.get("region"):
        _chart_with_insights_slide(
            prs, "Regional Performance", chart_paths["region"],
            "What this shows", insights.get("region", []), next_page(),
            subtitle="Target vs. collection across regions"
        )

    # ---- Slide 5: Collection Trend ----
    if chart_paths.get("trend"):
        _chart_with_insights_slide(
            prs, "Collection Trend", chart_paths["trend"],
            "What this shows", insights.get("trend", []), next_page(),
            subtitle="Daily collection over the reporting period"
        )

    # ---- Slide 6: Salesperson Leaderboard ----
    if chart_paths.get("leaderboard"):
        _chart_with_insights_slide(
            prs, "Salesperson Leaderboard", chart_paths["leaderboard"],
            "What this shows", insights.get("salesperson", []), next_page(),
            subtitle="Individual contribution to total collection"
        )

    # ---- Slide 7: Detailed Breakdown Table ----
    if region_df is not None and not region_df.empty:
        s7 = _blank_slide(prs)
        _slide_header(s7, "Regional Breakdown", "Detailed figures by region")

        has_target = "Target" in region_df.columns
        headers = ["Region", "Collection"] + (["Target", "Achievement"] if has_target else [])
        rows = len(region_df) + 1
        cols = len(headers)
        table_shape = s7.shapes.add_table(rows, cols, MARGIN, Inches(1.75), Inches(12.1), Inches(0.5) * rows)
        table = table_shape.table

        for c, h in enumerate(headers):
            cell = table.cell(0, c)
            cell.text = h
            cell.fill.solid()
            cell.fill.fore_color.rgb = NAVY
            para = cell.text_frame.paragraphs[0]
            para.font.size = Pt(13)
            para.font.bold = True
            para.font.color.rgb = WHITE

        for r, (_, row) in enumerate(region_df.iterrows(), start=1):
            values = [str(row["Region"]), f"\u20b9{row['Collection']:,.0f}"]
            if has_target:
                achievement = (row["Collection"] / row["Target"] * 100) if row["Target"] else 0
                values += [f"\u20b9{row['Target']:,.0f}", f"{achievement:.1f}%"]
            for c, val in enumerate(values):
                cell = table.cell(r, c)
                cell.text = val
                cell.fill.solid()
                cell.fill.fore_color.rgb = LIGHT_BG if r % 2 == 0 else WHITE
                para = cell.text_frame.paragraphs[0]
                para.font.size = Pt(12)
                para.font.color.rgb = NAVY
        _footer(s7, next_page())

    # ---- Slide 8: Recommendations ----
    if insights.get("recommendations"):
        s8 = _blank_slide(prs)
        _slide_header(s8, "Recommendations", "Suggested next steps based on this period's data")
        _bullet_list(s8, MARGIN, Inches(1.9), Inches(12.1), Inches(4.8),
                     insights["recommendations"], size=16, line_spacing=1.6, gap_after=18,
                     bullet_color=CORAL)
        _footer(s8, next_page())

    # ---- Slide 9: Closing ----
    s9 = _blank_slide(prs, bg=NAVY)
    _textbox(s9, Inches(0.8), Inches(3.2), Inches(11), Inches(1),
              "Thank You", size=36, bold=True, color=WHITE)
    _textbox(s9, Inches(0.8), Inches(4.0), Inches(11), Inches(0.5),
              "Questions and discussion welcome.", size=15, color=RGBColor(0xB8, 0xC2, 0xD6))

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    prs.save(output_path)
    return output_path
