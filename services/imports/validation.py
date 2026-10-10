"""
Financial Data Validation Engine.
Performs row-level and statement-level validations on incoming financial data.
Ensures missing data is flagged explicitly and calculations follow strict accounting principles.
"""
from dataclasses import dataclass, field
from decimal import Decimal
from typing import List, Dict, Any, Optional
from services.finance.canonical_schema import ReportContext


@dataclass
class ValidationIssue:
    severity: str  # 'ERROR', 'WARNING', 'INFO'
    sheet: str
    row: Optional[int]
    column: Optional[str]
    message: str
    field_key: Optional[str] = None


@dataclass
class ValidationResult:
    is_valid: bool
    errors: List[ValidationIssue] = field(default_factory=list)
    warnings: List[ValidationIssue] = field(default_factory=list)
    infos: List[ValidationIssue] = field(default_factory=list)

    @property
    def has_errors(self) -> bool:
        return len(self.errors) > 0

    @property
    def total_issues(self) -> int:
        return len(self.errors) + len(self.warnings) + len(self.infos)


class FinancialDataValidator:
    """
    Validates parsed raw dictionary or ReportContext before generation.
    """

    @classmethod
    def validate_raw_payload(cls, payload: Dict[str, Any]) -> ValidationResult:
        result = ValidationResult(is_valid=True)

        pnl = payload.get("income_statement", {})
        rev_items = pnl.get("revenue_items", [])
        cogs_items = pnl.get("cogs_items", [])
        opex_items = pnl.get("opex_items", [])

        # 1. Check for minimum required financial data
        if not rev_items and not pnl.get("total_revenue"):
            result.errors.append(ValidationIssue(
                severity="ERROR",
                sheet="P&L",
                row=None,
                column="Revenue",
                message="Income Statement is missing Revenue line items."
            ))
            result.is_valid = False

        if not cogs_items and not opex_items:
            result.warnings.append(ValidationIssue(
                severity="WARNING",
                sheet="P&L",
                row=None,
                column="Expenses",
                message="No Cost of Sales or Operating Expenses detected."
            ))

        # 2. Check for missing actuals in revenue
        for idx, item in enumerate(rev_items, start=1):
            if item.get("actual") is None:
                result.warnings.append(ValidationIssue(
                    severity="WARNING",
                    sheet="P&L",
                    row=idx,
                    column="Actual",
                    message=f"Revenue item '{item.get('label')}' is missing Actual value."
                ))

        # 3. Check Balance Sheet if present
        bs = payload.get("balance_sheet", {})
        curr_assets = bs.get("current_assets", [])
        curr_liab = bs.get("current_liabilities", [])
        equity = bs.get("equity_items", [])

        if not curr_assets and not bs.get("non_current_assets"):
            result.warnings.append(ValidationIssue(
                severity="WARNING",
                sheet="Balance Sheet",
                row=None,
                column="Assets",
                message="Balance Sheet data is incomplete or empty."
            ))

        # 4. Check KPIs
        kpis = payload.get("kpis", [])
        if not kpis:
            result.infos.append(ValidationIssue(
                severity="INFO",
                sheet="KPIs",
                row=None,
                column=None,
                message="No KPIs provided. KPI analysis slide will show general performance indicators."
            ))

        return result

    @classmethod
    def validate_context(cls, context: ReportContext) -> ValidationResult:
        """
        Statement-level accounting consistency validation on typed ReportContext.
        """
        result = ValidationResult(is_valid=True)
        pnl = context.income_statement
        bs = context.balance_sheet
        cf = context.cash_flow

        # P&L Revenue check
        if not pnl.revenue_items and not pnl.total_revenue:
            result.errors.append(ValidationIssue(
                severity="ERROR",
                sheet="P&L",
                row=None,
                column="Revenue",
                message="No revenue line items present in report."
            ))
            result.is_valid = False

        # Balance Sheet equality check (Assets = Liabilities + Equity)
        if bs.total_assets and bs.total_assets.actual is not None:
            total_liab = bs.total_liabilities.actual if bs.total_liabilities and bs.total_liabilities.actual is not None else Decimal("0")
            total_eq = bs.total_equity.actual if bs.total_equity and bs.total_equity.actual is not None else Decimal("0")
            bs_diff = abs(bs.total_assets.actual - (total_liab + total_eq))
            if bs_diff > Decimal("1.00"):
                result.warnings.append(ValidationIssue(
                    severity="WARNING",
                    sheet="Balance Sheet",
                    row=None,
                    column="Balance",
                    message=f"Balance sheet out of balance by {bs_diff:,.2f} (Assets: {bs.total_assets.actual:,.2f} vs Liab+Equity: {(total_liab + total_eq):,.2f})."
                ))

        # Cash Flow reconciliation check
        if cf.ending_cash and cf.beginning_cash and cf.net_change_in_cash:
            if cf.ending_cash.actual is not None and cf.beginning_cash.actual is not None and cf.net_change_in_cash.actual is not None:
                calc_ending = cf.beginning_cash.actual + cf.net_change_in_cash.actual
                if abs(cf.ending_cash.actual - calc_ending) > Decimal("1.00"):
                    result.warnings.append(ValidationIssue(
                        severity="WARNING",
                        sheet="Cash Flow",
                        row=None,
                        column="Ending Cash",
                        message=f"Cash flow ending cash ({cf.ending_cash.actual:,.2f}) does not match beginning + change ({calc_ending:,.2f})."
                    ))

        return result
