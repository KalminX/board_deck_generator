"""
Tests for Seed Demo Command and Explicit Missing Data Handling.
"""
from decimal import Decimal
from django.core.management import call_command
from django.test import TestCase
from board_deck.models import PresentationDeck, UploadedWorkbook, NormalizedSnapshot
from services.finance.canonical_schema import (
    OrganisationData, ReportingPeriodData, FinancialLineItem,
    IncomeStatementData, BalanceSheetData, CashFlowData, ReportContext
)
from services.finance.metrics import calculate_financial_metrics
from services.presentations.builder import PresentationBuilder


class SeedDemoTestCase(TestCase):
    def test_seed_board_demo_command(self):
        call_command("seed_board_demo")
        deck = PresentationDeck.objects.first()
        self.assertIsNotNone(deck)
        self.assertEqual(deck.organisation.name, "Apex Technologies Holdings Ltd")
        self.assertEqual(deck.reporting_period.label, "Q3 2026")
        self.assertEqual(len(deck.selected_slides), 10)
        self.assertTrue(bool(deck.pptx_file))


def test_missing_data_never_silently_zeros():
    # Verify that missing budget or prior values remain None rather than 0
    item = FinancialLineItem(
        key="custom_rev",
        label="Custom Stream",
        actual=Decimal("100000"),
        budget=None,
        prior=None
    )
    assert item.variance_vs_budget is None
    assert item.variance_pct_vs_budget is None
    assert item.variance_vs_prior is None
    assert item.variance_pct_vs_prior is None

    pnl = IncomeStatementData(revenue_items=[item])
    bs = BalanceSheetData()
    cf = CashFlowData()

    context = ReportContext(
        organisation=OrganisationData(name="Incomplete Data Co"),
        period=ReportingPeriodData(label="Q1 2026"),
        income_statement=pnl,
        balance_sheet=bs,
        cash_flow=cf
    )

    context = calculate_financial_metrics(context)
    builder = PresentationBuilder()
    prs, manifest, val_out = builder.build_presentation(context, selected_slides=["cover", "pnl"])

    assert len(prs.slides) == 2
    assert val_out.is_valid is True
