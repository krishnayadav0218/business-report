"""
report_builder.py
Builds the final PPTX report from KPIs + chart images.

Design note: instead of hand-editing a fixed .pptx template's placeholders
(brittle -- breaks the moment a shape gets renamed), this builds each slide
programmatically with python-pptx using a fixed style function (colors,
fonts, spacing). Same end result -- consistent branded deck every run --
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
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)


def _blank_slide(prs, bg=WHITE):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # fully blank layout
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = bg
    return slide


def _textbox(slide, left, top, width, height, text, size=18, bold=False,
             color=NAVY, align=PP_ALIGN.LEFT, font="Calibri"):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font
    return box


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


def build_report(kpis: dict, chart_paths: dict, output_path="output/Weekly_Business_Report.pptx",
                  report_title="Weekly Business Report", period_label=None):
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    period_label = period_label or date.today().strftime("%d %b %Y")

    # ---- Slide 1: Title ----
    s1 = _blank_slide(prs, bg=NAVY)
    _textbox(s1, Inches(0.8), Inches(2.8), Inches(11), Inches(1.2),
              report_title, size=40, bold=True, color=WHITE)
    _textbox(s1, Inches(0.8), Inches(3.8), Inches(11), Inches(0.6),
              f"Auto-generated on {period_label}", size=16, color=RGBColor(0xB8, 0xC2, 0xD6))

    # ---- Slide 2: KPI Summary ----
    s2 = _blank_slide(prs)
    _textbox(s2, Inches(0.6), Inches(0.4), Inches(10), Inches(0.6),
              "Summary", size=28, bold=True, color=NAVY)

    card_w, card_h = Inches(2.9), Inches(1.5)
    gap = Inches(0.3)
    top = Inches(1.4)
    cards = [
        ("Total Collection", f"₹{kpis['total_collection']:,.0f}", TEAL),
        ("Target Achievement", f"{kpis['achievement_pct']:.1f}%", NAVY),
        ("Total Orders", f"{kpis['total_orders']:,}", CORAL),
        ("Avg Order Value", f"₹{kpis['aov']:,.0f}", TEAL),
    ]
    left = Inches(0.6)
    for label, value, accent in cards:
        _kpi_card(s2, left, top, card_w, card_h, label, value, accent)
        left = Emu(left + card_w + gap)

    profit_color = TEAL if kpis["profit"] >= 0 else CORAL
    profit_label = "Profit" if kpis["profit"] >= 0 else "Loss"
    _kpi_card(s2, Inches(0.6), Inches(3.2), Inches(2.9), Inches(1.5),
              profit_label, f"₹{abs(kpis['profit']):,.0f}", profit_color)

    if kpis.get("wow_growth") is not None:
        growth_color = TEAL if kpis["wow_growth"] >= 0 else CORAL
        arrow = "▲" if kpis["wow_growth"] >= 0 else "▼"
        _kpi_card(s2, Inches(3.8), Inches(3.2), Inches(2.9), Inches(1.5),
                  "Period-on-Period Growth", f"{arrow} {abs(kpis['wow_growth']):.1f}%", growth_color)

    # ---- Slide 3: Region chart ----
    if chart_paths.get("region"):
        s3 = _blank_slide(prs)
        _textbox(s3, Inches(0.6), Inches(0.4), Inches(10), Inches(0.6),
                  "Regional Performance", size=28, bold=True, color=NAVY)
        s3.shapes.add_picture(chart_paths["region"], Inches(0.6), Inches(1.3), width=Inches(12.1))

    # ---- Slide 4: Trend chart ----
    if chart_paths.get("trend"):
        s4 = _blank_slide(prs)
        _textbox(s4, Inches(0.6), Inches(0.4), Inches(10), Inches(0.6),
                  "Collection Trend", size=28, bold=True, color=NAVY)
        s4.shapes.add_picture(chart_paths["trend"], Inches(0.6), Inches(1.3), width=Inches(12.1))

    # ---- Slide 5: Leaderboard chart ----
    if chart_paths.get("leaderboard"):
        s5 = _blank_slide(prs)
        _textbox(s5, Inches(0.6), Inches(0.4), Inches(10), Inches(0.6),
                  "Salesperson Leaderboard", size=28, bold=True, color=NAVY)
        s5.shapes.add_picture(chart_paths["leaderboard"], Inches(0.6), Inches(1.3), width=Inches(12.1))

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    prs.save(output_path)
    return output_path
