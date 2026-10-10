"""
Unit Tests for Presentation Builder & Slide Renderers.
"""
import pytest
from pptx import Presentation
from decimal import Decimal
from services.finance.canonical_schema import (
    OrganisationData, ReportingPeriodData, FinancialLineItem,
    IncomeStatementData, BalanceSheetData, CashFlowData, KPIData, ReportContext,
    RiskItem, ActionItem
)
from services.presentations.builder import PresentationBuilder, SLIDE_REGISTRY
from services.presentations.tokens import DEFAULT_BRAND, MODERN_NAVY_BRAND


def sample_full_context():
    pnl = IncomeStatementData(
        revenue_items=[FinancialLineItem("sub", "Subscriptions", actual=Decimal("2000000"), budget=Decimal("1900000"))],
        cogs_items=[FinancialLineItem("host", "Cloud Hosting", actual=Decimal("400000"), budget=Decimal("380000"))],
        opex_items=[FinancialLineItem("sales", "Sales & Marketing", actual=Decimal("800000"), budget=Decimal("750000"))],
        monthly_trend=[
            {"period": "Jul 2026", "revenue": 650000, "opex": 400000, "profit": 120000},
            {"period": "Aug 2026", "revenue": 670000, "opex": 400000, "profit": 130000},
            {"period": "Sep 2026", "revenue": 680000, "opex": 400000, "profit": 150000}
        ]
    )
    bs = BalanceSheetData(
        current_assets=[FinancialLineItem("cash", "Cash", actual=Decimal("5000000"), prior=Decimal("4500000"))],
        current_liabilities=[FinancialLineItem("ap", "Payables", actual=Decimal("600000"), prior=Decimal("550000"))],
        equity_items=[FinancialLineItem("eq", "Share Capital", actual=Decimal("4400000"), prior=Decimal("4400000"))]
    )
    cf = CashFlowData(
        operating_activities=[FinancialLineItem("cf_op", "Operating Cash", actual=Decimal("600000"))],
        investing_activities=[FinancialLineItem("cf_inv", "Capex", actual=Decimal("-100000"))],
        cash_runway_months=Decimal("32.0")
    )
    kpis = [
        KPIData(key="arr", name="ARR ($M)", current_value=Decimal("12.5"), target_value=Decimal("12.0"), unit="$M",
                monthly_history=[{"period": "Jul", "value": 11.5}, {"period": "Aug", "value": 12.0}, {"period": "Sep", "value": 12.5}])
    ]
    risks = [
        RiskItem(id="r1", title="Regulatory Compliance Review", category="Compliance", severity="Medium", likelihood="Low", description="Annual GDPR audit scheduled.", mitigation_plan="Compliance team assigned.")
    ]
    actions = [
        ActionItem(id="a1", title="Complete Series C Financial Audit", reason="Preparation for upcoming expansion financing.", priority="High", owner="CFO", due_date="2026-11-30", status="In Progress")
    ]

    return ReportContext(
        organisation=OrganisationData(name="Apex Tech"),
        period=ReportingPeriodData(label="Q3 2026"),
        income_statement=pnl,
        balance_sheet=bs,
        cash_flow=cf,
        kpis=kpis,
        risks=risks,
        actions=actions
    )


def test_build_all_10_slides():
    context = sample_full_context()
    builder = PresentationBuilder(brand=DEFAULT_BRAND, narrative_style="standard")
    prs, manifest, val_out = builder.build_presentation(context)

    assert len(prs.slides) == 10
    assert manifest["slide_count"] == 10
    assert val_out.is_valid is True

    # Check Slide 4 (PnL) has native table
    pnl_slide = prs.slides[3]
    has_table = any(s.has_table for s in pnl_slide.shapes)
    assert has_table is True

    # Check Slide 3 (Financial Overview) has native chart
    overview_slide = prs.slides[2]
    has_chart = any(s.has_chart for s in overview_slide.shapes)
    assert has_chart is True


def test_build_subset_of_slides():
    context = sample_full_context()
    builder = PresentationBuilder(brand=MODERN_NAVY_BRAND)
    selected = ["cover", "executive_summary", "pnl"]
    prs, manifest, _ = builder.build_presentation(context, selected_slides=selected)

    assert len(prs.slides) == 3
    assert manifest["slide_count"] == 3


def test_build_with_custom_commentary_overrides():
    context = sample_full_context()
    builder = PresentationBuilder()
    overrides = {
        "executive_summary": {
            "lead_summary": "Customized CEO Summary for Board Meeting.",
            "bullet_points": ["Custom point 1", "Custom point 2"]
        }
    }
    prs, manifest, _ = builder.build_presentation(context, custom_commentary=overrides)
    assert len(prs.slides) == 10
