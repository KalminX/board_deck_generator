"""
Tests for Excel Ingestion and Financial Validation.
"""
import os
import openpyxl
from io import BytesIO
from services.imports.excel import ExcelImportService
from services.imports.validation import FinancialDataValidator


def test_excel_parse_canonical_workbook():
    sample_path = "sample_data/sample_board_deck_data.xlsx"
    assert os.path.exists(sample_path)

    with open(sample_path, "rb") as f:
        file_bytes = f.read()

    sheet_info = ExcelImportService.inspect_workbook(file_bytes)
    assert "P&L" in sheet_info
    assert "Balance Sheet" in sheet_info

    payload = ExcelImportService.parse_workbook(file_bytes)
    assert payload["organisation"]["name"] == "Apex Technologies Holdings Ltd"
    assert len(payload["income_statement"]["revenue_items"]) == 3
    assert len(payload["balance_sheet"]["current_assets"]) >= 1
    assert len(payload["kpis"]) >= 1

    validation = FinancialDataValidator.validate_raw_payload(payload)
    assert validation.is_valid is True
    assert len(validation.errors) == 0


def test_validation_detects_missing_revenue():
    invalid_payload = {
        "organisation": {"name": "Empty Corp"},
        "period": {"label": "Q1 2026"},
        "income_statement": {
            "revenue_items": [],
            "cogs_items": [],
            "opex_items": []
        }
    }
    validation = FinancialDataValidator.validate_raw_payload(invalid_payload)
    assert validation.is_valid is False
    assert any("Income Statement is missing Revenue" in e.message for e in validation.errors)


def test_excel_inspect_custom_workbook():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Overview"
    ws.append(["Company", "Beta Solutions Inc"])
    ws.append(["Currency", "GBP"])
    ws.append(["Period", "FY 2026"])

    ws_pnl = wb.create_sheet(title="Income Statement")
    ws_pnl.append(["Line Item", "Actual", "Budget"])
    ws_pnl.append(["Consulting Revenue", 500000, 450000])

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    data = bio.getvalue()

    sheet_info = ExcelImportService.inspect_workbook(data)
    assert "Overview" in sheet_info
    assert "Income Statement" in sheet_info

    payload = ExcelImportService.parse_workbook(data)
    assert payload["organisation"]["name"] == "Beta Solutions Inc"
    assert payload["organisation"]["currency_code"] == "GBP"
    assert len(payload["income_statement"]["revenue_items"]) >= 1
