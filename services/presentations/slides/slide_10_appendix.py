"""
Slide 10: Appendix, Governance & Methodology Slide.
"""
from pptx.util import Inches, Pt
from datetime import datetime
from services.finance.canonical_schema import ReportContext
from services.commentary.base import PresentationCommentary
from services.presentations.tokens import BrandConfig, DEFAULT_BRAND
from services.presentations.layouts import apply_slide_header, apply_slide_footer, create_card_container


def render_appendix_slide(prs, context: ReportContext, commentary: PresentationCommentary, current_page: int, total_pages: int, brand: BrandConfig = DEFAULT_BRAND):
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)

    apply_slide_header(slide, "Appendix: Methodology, Definitions & Data Provenance", "Governance & Audit", context.period.label, brand)
    apply_slide_footer(slide, current_page, total_pages, context.organisation.name, context.period.label, brand)

    col_w = Inches(5.95)
    gap = Inches(0.233)
    card_h = Inches(2.65)
    top1 = Inches(1.2)
    top2 = Inches(4.1)

    # 1. Card Top-Left: Accounting Policies & Basis of Preparation
    create_card_container(slide, brand.margin_left, top1, col_w, card_h, "1. ACCOUNTING POLICIES & BASIS OF PREPARATION", brand)
    tb1 = slide.shapes.add_textbox(brand.margin_left + Inches(0.2), top1 + Inches(0.42), col_w - Inches(0.4), card_h - Inches(0.5))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    tf1.margin_left = tf1.margin_top = tf1.margin_right = tf1.margin_bottom = 0

    p1 = tf1.paragraphs[0]
    p1.text = f"• Reporting Entity: {context.organisation.name}"
    p1.font.name = brand.font_body
    p1.font.size = Pt(8.5)
    p1.font.color.rgb = brand.primary_color
    p1.space_after = Pt(2)

    p2 = tf1.add_paragraph()
    p2.text = f"• Functional & Reporting Currency: {context.organisation.currency_code} ({context.organisation.currency_symbol})"
    p2.font.name = brand.font_body
    p2.font.size = Pt(8.5)
    p2.font.color.rgb = brand.primary_color
    p2.space_after = Pt(2)

    p3 = tf1.add_paragraph()
    p3.text = "• Basis of Preparation: Accrual accounting principles aligned with internal management accounts framework."
    p3.font.name = brand.font_body
    p3.font.size = Pt(8.5)
    p3.font.color.rgb = brand.primary_color
    p3.space_after = Pt(2)

    # 2. Card Top-Right: Materiality & Variance Framework
    create_card_container(slide, brand.margin_left + col_w + gap, top1, col_w, card_h, "2. VARIANCE THRESHOLDS & MATERIALITY", brand)
    tb2 = slide.shapes.add_textbox(brand.margin_left + col_w + gap + Inches(0.2), top1 + Inches(0.42), col_w - Inches(0.4), card_h - Inches(0.5))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    tf2.margin_left = tf2.margin_top = tf2.margin_right = tf2.margin_bottom = 0

    p4 = tf2.paragraphs[0]
    p4.text = "• Materiality Threshold: Financial variances exceeding ±5.0% and >$25,000 against approved budget."
    p4.font.name = brand.font_body
    p4.font.size = Pt(8.5)
    p4.font.color.rgb = brand.primary_color
    p4.space_after = Pt(2)

    p5 = tf2.add_paragraph()
    p5.text = "• Favorable / Adverse Criteria: Revenue increases are classified as favorable; operating expense overruns are adverse."
    p5.font.name = brand.font_body
    p5.font.size = Pt(8.5)
    p5.font.color.rgb = brand.primary_color
    p5.space_after = Pt(2)

    # 3. Card Bottom-Left: Metric Definitions
    create_card_container(slide, brand.margin_left, top2, col_w, card_h, "3. KEY FINANCIAL METRICS & DEFINITIONS", brand)
    tb3 = slide.shapes.add_textbox(brand.margin_left + Inches(0.2), top2 + Inches(0.42), col_w - Inches(0.4), card_h - Inches(0.5))
    tf3 = tb3.text_frame
    tf3.word_wrap = True
    tf3.margin_left = tf3.margin_top = tf3.margin_right = tf3.margin_bottom = 0

    p6 = tf3.paragraphs[0]
    p6.text = "• Operating Profit (EBIT): Earnings before interest and corporate taxation."
    p6.font.name = brand.font_body
    p6.font.size = Pt(8.5)
    p6.font.color.rgb = brand.primary_color
    p6.space_after = Pt(2)

    p7 = tf3.add_paragraph()
    p7.text = "• Net Working Capital: Total current assets less total current liabilities."
    p7.font.name = brand.font_body
    p7.font.size = Pt(8.5)
    p7.font.color.rgb = brand.primary_color
    p7.space_after = Pt(2)

    p8 = tf3.add_paragraph()
    p8.text = "• Cash Runway: Projected runway based on 3-month trailing operating cash conversion."
    p8.font.name = brand.font_body
    p8.font.size = Pt(8.5)
    p8.font.color.rgb = brand.primary_color

    # 4. Card Bottom-Right: Data Provenance & Snapshot Hash
    create_card_container(slide, brand.margin_left + col_w + gap, top2, col_w, card_h, "4. DATA SOURCES & AUDIT INTEGRITY", brand)
    tb4 = slide.shapes.add_textbox(brand.margin_left + col_w + gap + Inches(0.2), top2 + Inches(0.42), col_w - Inches(0.4), card_h - Inches(0.5))
    tf4 = tb4.text_frame
    tf4.word_wrap = True
    tf4.margin_left = tf4.margin_top = tf4.margin_right = tf4.margin_bottom = 0

    p9 = tf4.paragraphs[0]
    p9.text = f"• Source Snapshot: Canonical ERP & Financial Ledger sync."
    p9.font.name = brand.font_body
    p9.font.size = Pt(8.5)
    p9.font.color.rgb = brand.primary_color
    p9.space_after = Pt(2)

    p10 = tf4.add_paragraph()
    p10.text = f"• Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S UTC')}"
    p10.font.name = brand.font_body
    p10.font.size = Pt(8.5)
    p10.font.color.rgb = brand.primary_color
    p10.space_after = Pt(2)

    p11 = tf4.add_paragraph()
    p11.text = f"• Audit Reference: SEC-AUDIT-{context.period.label.replace(' ', '')}-CANONICAL"
    p11.font.name = brand.font_body
    p11.font.size = Pt(8.5)
    p11.font.color.rgb = brand.primary_color

    return slide
