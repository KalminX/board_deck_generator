"""
Unit Tests for Commentary Generation Provider.
"""
from decimal import Decimal
from services.finance.canonical_schema import (
    OrganisationData, ReportingPeriodData, FinancialLineItem,
    IncomeStatementData, BalanceSheetData, CashFlowData, KPIData, ReportContext
)
from services.finance.metrics import calculate_financial_metrics
from services.commentary.factory import get_commentary_provider
from services.commentary.templates import TemplateCommentaryProvider


def build_sample_context():
    pnl = IncomeStatementData(
        revenue_items=[FinancialLineItem("rev", "Software Revenue", actual=Decimal("1500000"), budget=Decimal("1400000"), prior=Decimal("1200000"))],
        cogs_items=[FinancialLineItem("cogs", "Direct Costs", actual=Decimal("300000"), budget=Decimal("280000"))],
        opex_items=[FinancialLineItem("opex", "Operating Costs", actual=Decimal("700000"), budget=Decimal("680000"))]
    )
    bs = BalanceSheetData(
        current_assets=[FinancialLineItem("cash", "Cash Reserves", actual=Decimal("4500000"), prior=Decimal("4000000"))]
    )
    cf = CashFlowData(
        operating_activities=[FinancialLineItem("cf_op", "Operating Cash", actual=Decimal("500000"))],
        cash_runway_months=Decimal("24.0")
    )
    kpis = [
        KPIData(key="arr", name="ARR", current_value=Decimal("15.0"), target_value=Decimal("14.0"), unit="$M")
    ]

    context = ReportContext(
        organisation=OrganisationData(name="Apex Tech", currency_symbol="$"),
        period=ReportingPeriodData(label="Q3 2026"),
        income_statement=pnl,
        balance_sheet=bs,
        cash_flow=cf,
        kpis=kpis
    )
    return calculate_financial_metrics(context)


def test_template_commentary_generates_grounded_facts():
    context = build_sample_context()
    provider = TemplateCommentaryProvider()
    commentary = provider.generate_commentary(context, style="standard")

    assert commentary.executive_summary is not None
    assert "$1.5M" in commentary.executive_summary.lead_summary or "$1.50M" in commentary.executive_summary.lead_summary
    assert len(commentary.executive_summary.fact_references) >= 1
    assert any(ref.fact_id == "fact:pnl:total_revenue:actual" for ref in commentary.executive_summary.fact_references)

    assert commentary.pnl_commentary is not None
    assert commentary.balance_sheet_commentary is not None
    assert commentary.cash_flow_commentary is not None
    assert commentary.kpi_commentary is not None


def test_narrative_style_lengths():
    context = build_sample_context()
    provider = TemplateCommentaryProvider()

    concise = provider.generate_commentary(context, style="concise")
    standard = provider.generate_commentary(context, style="standard")

    assert len(concise.executive_summary.bullet_points) <= 3
    assert len(standard.executive_summary.bullet_points) <= 4


def test_commentary_factory():
    provider = get_commentary_provider("template")
    assert isinstance(provider, TemplateCommentaryProvider)

    # Unknown provider defaults safely to template
    fallback = get_commentary_provider("unknown_provider")
    assert isinstance(fallback, TemplateCommentaryProvider)
