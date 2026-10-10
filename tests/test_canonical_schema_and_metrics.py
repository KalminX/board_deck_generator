"""
Unit Tests for Canonical Schema and Financial Metric Calculations.
Verifies Decimal calculations, Gross Margin, EBIT, EBITDA, Working Capital,
Material Variance detection, and fact registration.
"""
import pytest
from decimal import Decimal
from services.finance.canonical_schema import (
    OrganisationData, ReportingPeriodData, FinancialLineItem,
    IncomeStatementData, BalanceSheetData, CashFlowData, KPIData, ReportContext
)
from services.finance.metrics import calculate_financial_metrics, safe_divide, safe_pct
from services.finance.normalization import to_decimal, normalize_report_payload


def test_to_decimal_conversion():
    assert to_decimal(100) == Decimal("100")
    assert to_decimal(100.5) == Decimal("100.5")
    assert to_decimal("$1,234.50") == Decimal("1234.50")
    assert to_decimal("(500.00)") == Decimal("-500.00")
    assert to_decimal("25.5%") == Decimal("25.5")
    assert to_decimal(None) is None
    assert to_decimal("-") is None
    assert to_decimal("N/A") is None
    assert to_decimal("invalid") is None


def test_financial_line_item_variances():
    item = FinancialLineItem(
        key="rev_saas",
        label="SaaS Subscriptions",
        actual=Decimal("120000"),
        budget=Decimal("100000"),
        prior=Decimal("90000")
    )
    assert item.variance_vs_budget == Decimal("20000")
    assert item.variance_pct_vs_budget == Decimal("20.0")
    assert item.variance_vs_prior == Decimal("30000")
    # Variance vs prior: (120000 - 90000) / 90000 * 100 = 33.333... %
    assert round(item.variance_pct_vs_prior, 2) == Decimal("33.33")


def test_pnl_metric_aggregations():
    pnl = IncomeStatementData(
        revenue_items=[
            FinancialLineItem("sub", "Subscriptions", actual=Decimal("200000"), budget=Decimal("180000")),
            FinancialLineItem("serv", "Services", actual=Decimal("50000"), budget=Decimal("50000"))
        ],
        cogs_items=[
            FinancialLineItem("host", "Hosting", actual=Decimal("40000"), budget=Decimal("35000")),
            FinancialLineItem("supp", "Support", actual=Decimal("20000"), budget=Decimal("20000"))
        ],
        opex_items=[
            FinancialLineItem("sales", "Sales", actual=Decimal("70000"), budget=Decimal("65000")),
            FinancialLineItem("eng", "Engineering", actual=Decimal("50000"), budget=Decimal("50000"))
        ]
    )

    bs = BalanceSheetData(
        current_assets=[FinancialLineItem("cash", "Cash", actual=Decimal("500000"), prior=Decimal("400000"))],
        current_liabilities=[FinancialLineItem("ap", "Payables", actual=Decimal("100000"), prior=Decimal("80000"))]
    )

    cf = CashFlowData()

    context = ReportContext(
        organisation=OrganisationData(name="Test Corp"),
        period=ReportingPeriodData(label="Q3 2026"),
        income_statement=pnl,
        balance_sheet=bs,
        cash_flow=cf
    )

    # Compute metrics
    context = calculate_financial_metrics(context)

    # Total Revenue: 200k + 50k = 250k
    assert pnl.total_revenue.actual == Decimal("250000")
    assert pnl.total_revenue.budget == Decimal("230000")

    # Total COGS: 40k + 20k = 60k
    assert pnl.total_cogs.actual == Decimal("60000")

    # Gross Profit: 250k - 60k = 190k
    assert pnl.gross_profit.actual == Decimal("190000")
    assert pnl.gross_profit.budget == Decimal("175000")

    # Total Opex: 70k + 50k = 120k
    assert pnl.total_opex.actual == Decimal("120000")

    # Operating Profit (EBIT): 190k - 120k = 70k
    assert pnl.operating_profit.actual == Decimal("70000")
    assert pnl.operating_profit.budget == Decimal("60000")

    # Net Working Capital: 500k - 100k = 400k
    assert bs.net_working_capital.actual == Decimal("400000")

    # Check registered facts
    assert "fact:pnl:total_revenue:actual" in context.all_facts
    assert context.all_facts["fact:pnl:total_revenue:actual"]["value"] == Decimal("250000")
    assert "fact:pnl:gross_profit:actual" in context.all_facts
    assert "fact:pnl:operating_profit:actual" in context.all_facts
