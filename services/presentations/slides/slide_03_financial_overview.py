"""
Slide 3: Financial Performance Overview.
"""
from pptx.util import Inches, Pt
from decimal import Decimal
from services.finance.canonical_schema import ReportContext
from services.commentary.base import PresentationCommentary
from services.presentations.tokens import BrandConfig, DEFAULT_BRAND
from services.presentations.layouts import apply_slide_header, apply_slide_footer, create_card_container, create_commentary_box
from services.presentations.charts import create_actual_vs_budget_chart, create_monthly_trend_chart


def render_financial_overview_slide(prs, context: ReportContext, commentary: PresentationCommentary, current_page: int, total_pages: int, brand: BrandConfig = DEFAULT_BRAND):
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)

    apply_slide_header(slide, "Financial Performance Overview", "Financial Analysis", context.period.label, brand)
    apply_slide_footer(slide, current_page, total_pages, context.organisation.name, context.period.label, brand)

    pnl = context.income_statement

    # Left Side: 2 Native PowerPoint Charts
    left_w = Inches(6.8)
    chart_h = Inches(2.7)

    # 1. Actual vs Budget Clustered Chart
    categories = []
    actuals = []
    budgets = []

    if pnl.total_revenue:
        categories.append("Revenue")
        actuals.append(float(pnl.total_revenue.actual or 0) / 1000.0)
        budgets.append(float(pnl.total_revenue.budget or 0) / 1000.0)

    if pnl.gross_profit:
        categories.append("Gross Profit")
        actuals.append(float(pnl.gross_profit.actual or 0) / 1000.0)
        budgets.append(float(pnl.gross_profit.budget or 0) / 1000.0)

    if pnl.total_opex:
        categories.append("Total Opex")
        actuals.append(float(pnl.total_opex.actual or 0) / 1000.0)
        budgets.append(float(pnl.total_opex.budget or 0) / 1000.0)

    if pnl.operating_profit:
        categories.append("Operating Profit")
        actuals.append(float(pnl.operating_profit.actual or 0) / 1000.0)
        budgets.append(float(pnl.operating_profit.budget or 0) / 1000.0)

    if categories:
        create_actual_vs_budget_chart(
            slide,
            brand.margin_left, Inches(1.2), left_w, chart_h,
            categories, actuals, budgets,
            title="Core P&L Actual vs Budget ($k)",
            brand=brand
        )

    # 2. Historical Trend or Revenue Stream breakdown
    if pnl.monthly_trend:
        create_monthly_trend_chart(
            slide,
            brand.margin_left, Inches(4.1), left_w, chart_h,
            pnl.monthly_trend,
            title="Multi-Period Revenue & Profit Trend ($k)",
            brand=brand
        )
    elif pnl.revenue_items:
        rev_cats = [item.label for item in pnl.revenue_items[:4]]
        rev_acts = [float(item.actual or 0) / 1000.0 for item in pnl.revenue_items[:4]]
        rev_buds = [float(item.budget or 0) / 1000.0 for item in pnl.revenue_items[:4]]
        create_actual_vs_budget_chart(
            slide,
            brand.margin_left, Inches(4.1), left_w, chart_h,
            rev_cats, rev_acts, rev_buds,
            title="Revenue Lines Actual vs Budget ($k)",
            brand=brand
        )

    # Right Side: Commentary & Margin Breakdown
    right_left = brand.margin_left + left_w + Inches(0.25)
    right_w = Inches(5.08)

    sec = commentary.performance_overview
    create_commentary_box(slide, right_left, Inches(1.2), right_w, Inches(5.6), sec, "PERFORMANCE COMMENTARY & MARGINS", brand)

    return slide
