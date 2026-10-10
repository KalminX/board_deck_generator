"""
Unit Tests for PDF Export Service & High-Fidelity PDF Generation.
"""
import os
import tempfile
from decimal import Decimal
from services.exports.pdf import PDFExportService
from services.finance.canonical_schema import (
    OrganisationData, ReportingPeriodData, FinancialLineItem,
    IncomeStatementData, BalanceSheetData, CashFlowData, KPIData, ReportContext,
    RiskItem, ActionItem
)
from services.finance.metrics import calculate_financial_metrics
from services.commentary.factory import get_commentary_provider


def test_pdf_service_non_existent_file():
    success, msg = PDFExportService.convert_pptx_to_pdf("non_existent_path.pptx")
    assert success is False
    assert "not found" in msg.lower()


def test_pdf_service_availability_check():
    is_avail = PDFExportService.is_pdf_conversion_available()
    assert is_avail is True


def test_build_reportlab_pdf_output():
    pnl = IncomeStatementData(
        revenue_items=[FinancialLineItem("sub", "Subscriptions", actual=Decimal("2000000"), budget=Decimal("1900000"))],
        cogs_items=[FinancialLineItem("host", "Cloud Hosting", actual=Decimal("400000"), budget=Decimal("380000"))],
        opex_items=[FinancialLineItem("sales", "Sales & Marketing", actual=Decimal("800000"), budget=Decimal("750000"))]
    )
    bs = BalanceSheetData(
        current_assets=[FinancialLineItem("cash", "Cash", actual=Decimal("5000000"), prior=Decimal("4500000"))],
        current_liabilities=[FinancialLineItem("ap", "Payables", actual=Decimal("600000"), prior=Decimal("550000"))]
    )
    cf = CashFlowData(
        operating_activities=[FinancialLineItem("cf_op", "Operating Cash", actual=Decimal("600000"))],
        cash_runway_months=Decimal("32.0")
    )
    kpis = [
        KPIData(key="arr", name="ARR ($M)", current_value=Decimal("12.5"), target_value=Decimal("12.0"), unit="$M")
    ]
    risks = [
        RiskItem(id="r1", title="Compliance Audit", category="Compliance", severity="Medium", likelihood="Low", description="Annual review.", mitigation_plan="Team assigned.")
    ]
    actions = [
        ActionItem(id="a1", title="Close Series C", reason="Expansion funding.", priority="High", owner="CFO", due_date="2026-11-30", status="Open")
    ]

    context = ReportContext(
        organisation=OrganisationData(name="Apex Tech"),
        period=ReportingPeriodData(label="Q3 2026"),
        income_statement=pnl,
        balance_sheet=bs,
        cash_flow=cf,
        kpis=kpis,
        risks=risks,
        actions=actions
    )
    context = calculate_financial_metrics(context)
    comm_provider = get_commentary_provider("template")
    commentary = comm_provider.generate_commentary(context)

    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp_pdf:
        tmp_path = tmp_pdf.name

    try:
        PDFExportService._build_reportlab_pdf(
            output_pdf_path=tmp_path,
            context=context,
            commentary=commentary,
            selected_slides=["cover", "executive_summary", "financial_overview", "pnl", "balance_sheet", "cash_flow", "kpis", "risks", "actions", "appendix"],
            title="Board of Directors Review"
        )
        assert os.path.exists(tmp_path)
        assert os.path.getsize(tmp_path) > 1000  # Valid non-empty PDF
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
