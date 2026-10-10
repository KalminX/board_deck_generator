"""
Slide 4: Profit & Loss Statement Slide.
"""
from pptx.util import Inches, Pt
from services.finance.canonical_schema import ReportContext, FinancialLineItem
from services.commentary.base import PresentationCommentary
from services.presentations.tokens import BrandConfig, DEFAULT_BRAND
from services.presentations.layouts import apply_slide_header, apply_slide_footer, create_commentary_box
from services.presentations.tables import create_financial_statement_table


def render_pnl_slide(prs, context: ReportContext, commentary: PresentationCommentary, current_page: int, total_pages: int, brand: BrandConfig = DEFAULT_BRAND):
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)

    apply_slide_header(slide, "Profit & Loss Statement", "Financial Statements", context.period.label, brand)
    apply_slide_footer(slide, current_page, total_pages, context.organisation.name, context.period.label, brand)

    pnl = context.income_statement

    # Build line item list for table
    table_items = []

    # Revenue
    for r in pnl.revenue_items:
        table_items.append(r)
    if pnl.total_revenue:
        table_items.append(pnl.total_revenue)

    # COGS
    for c in pnl.cogs_items:
        table_items.append(c)
    if pnl.total_cogs:
        table_items.append(pnl.total_cogs)

    # Gross Profit
    if pnl.gross_profit:
        table_items.append(pnl.gross_profit)

    # Opex
    for o in pnl.opex_items:
        table_items.append(o)
    if pnl.total_opex:
        table_items.append(pnl.total_opex)

    # Operating Profit
    if pnl.operating_profit:
        table_items.append(pnl.operating_profit)

    # Limit to top 11 items to ensure no overflow
    display_items = table_items[:12]

    table_w = Inches(8.3)
    table_h = Inches(5.6)

    create_financial_statement_table(
        slide,
        brand.margin_left, Inches(1.2), table_w, table_h,
        display_items,
        show_prior=True,
        show_ytd=True,
        currency_symbol=context.organisation.currency_symbol,
        brand=brand
    )

    # Right side commentary
    right_left = brand.margin_left + table_w + Inches(0.25)
    right_w = Inches(3.58)

    sec = commentary.pnl_commentary
    create_commentary_box(slide, right_left, Inches(1.2), right_w, table_h, sec, "P&L VARIANCE NOTES", brand)

    return slide
