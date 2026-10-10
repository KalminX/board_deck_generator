"""
Normalization Service.
Transforms raw structured data (dictionaries, Excel rows, JSON) into Canonical Schema instances.
Safely casts values to Decimal, preserves None for missing figures, and attaches provenance.
"""
from decimal import Decimal, InvalidOperation
from typing import Dict, Any, List, Optional
from datetime import date, datetime
from services.finance.canonical_schema import (
    OrganisationData, ReportingPeriodData, FinancialLineItem,
    IncomeStatementData, BalanceSheetData, CashFlowData, KPIData,
    RiskItem, ActionItem, ReportContext
)


def to_decimal(val: Any) -> Optional[Decimal]:
    """Safely convert value to Decimal. Returns None for None/empty strings/invalid values."""
    if val is None:
        return None
    if isinstance(val, Decimal):
        return val
    if isinstance(val, (int, float)):
        return Decimal(str(val))
    if isinstance(val, str):
        val_str = val.strip().replace("$", "").replace(",", "").replace("%", "")
        if val_str == "" or val_str.lower() in ("-", "n/a", "none", "null"):
            return None
        # Handle accounting parentheses for negative numbers: (1,234.50) -> -1234.50
        if val_str.startswith("(") and val_str.endswith(")"):
            val_str = "-" + val_str[1:-1].strip()
        try:
            return Decimal(val_str)
        except InvalidOperation:
            return None
    return None


def normalize_line_item(raw: Dict[str, Any], default_category: str = "general") -> FinancialLineItem:
    return FinancialLineItem(
        key=str(raw.get("key", raw.get("label", "item"))).lower().replace(" ", "_").replace("-", "_"),
        label=str(raw.get("label", raw.get("name", "Line Item"))),
        actual=to_decimal(raw.get("actual")),
        budget=to_decimal(raw.get("budget")),
        prior=to_decimal(raw.get("prior")),
        ytd_actual=to_decimal(raw.get("ytd_actual", raw.get("ytd"))),
        ytd_budget=to_decimal(raw.get("ytd_budget")),
        category=raw.get("category", default_category),
        is_subtotal=bool(raw.get("is_subtotal", False)),
        is_total=bool(raw.get("is_total", False)),
        notes=raw.get("notes"),
        source_ref=raw.get("source_ref"),
        fact_id=raw.get("fact_id")
    )


def normalize_kpi(raw: Dict[str, Any]) -> KPIData:
    return KPIData(
        key=str(raw.get("key", raw.get("name", "kpi"))).lower().replace(" ", "_"),
        name=str(raw.get("name", raw.get("label", "KPI"))),
        category=str(raw.get("category", "financial")),
        current_value=to_decimal(raw.get("current_value", raw.get("actual"))),
        target_value=to_decimal(raw.get("target_value", raw.get("target", raw.get("budget")))),
        prior_value=to_decimal(raw.get("prior_value", raw.get("prior"))),
        unit=str(raw.get("unit", "$")),
        definition=str(raw.get("definition", "")),
        higher_is_better=bool(raw.get("higher_is_better", True)),
        notes=raw.get("notes"),
        fact_id=raw.get("fact_id", f"fact:kpi:{raw.get('key', 'kpi')}"),
        monthly_history=raw.get("monthly_history", [])
    )


def normalize_risk(raw: Dict[str, Any]) -> RiskItem:
    return RiskItem(
        id=str(raw.get("id", f"risk_{raw.get('title', 'risk')}")),
        title=str(raw.get("title", raw.get("name", "Risk"))),
        category=str(raw.get("category", "Financial")),
        severity=str(raw.get("severity", "Medium")),
        likelihood=str(raw.get("likelihood", "Medium")),
        description=str(raw.get("description", "")),
        mitigation_plan=str(raw.get("mitigation_plan", raw.get("mitigation", ""))),
        owner=raw.get("owner"),
        source_finding_id=raw.get("source_finding_id")
    )


def normalize_action(raw: Dict[str, Any]) -> ActionItem:
    return ActionItem(
        id=str(raw.get("id", f"action_{raw.get('title', 'action')}")),
        title=str(raw.get("title", raw.get("action", "Action"))),
        reason=str(raw.get("reason", raw.get("rationale", ""))),
        priority=str(raw.get("priority", "Medium")),
        owner=raw.get("owner"),
        due_date=str(raw.get("due_date", raw.get("date", ""))) if raw.get("due_date") or raw.get("date") else None,
        status=str(raw.get("status", "Open")),
        source_ref=raw.get("source_ref")
    )


def normalize_report_payload(payload: Dict[str, Any]) -> ReportContext:
    """
    Constructs a complete ReportContext from normalized dictionary payload.
    """
    org_raw = payload.get("organisation", {})
    org = OrganisationData(
        id=org_raw.get("id"),
        name=org_raw.get("name", "Apex Technologies Holdings Ltd"),
        legal_entity_number=org_raw.get("legal_entity_number"),
        industry=org_raw.get("industry", "Technology & Software"),
        currency_code=org_raw.get("currency_code", "USD"),
        currency_symbol=org_raw.get("currency_symbol", "$"),
        fiscal_year_end_month=int(org_raw.get("fiscal_year_end_month", 12))
    )

    per_raw = payload.get("period", {})
    period = ReportingPeriodData(
        id=per_raw.get("id"),
        label=per_raw.get("label", "Q3 2026"),
        period_type=per_raw.get("period_type", "quarter"),
        start_date=per_raw.get("start_date"),
        end_date=per_raw.get("end_date"),
        prior_period_label=per_raw.get("prior_period_label", "Q2 2026"),
        budget_period_label=per_raw.get("budget_period_label", "Q3 2026 Budget"),
        is_closed=bool(per_raw.get("is_closed", True))
    )

    # P&L
    pnl_raw = payload.get("income_statement", {})
    pnl = IncomeStatementData(
        revenue_items=[normalize_line_item(item, "revenue") for item in pnl_raw.get("revenue_items", [])],
        cogs_items=[normalize_line_item(item, "cogs") for item in pnl_raw.get("cogs_items", [])],
        opex_items=[normalize_line_item(item, "opex") for item in pnl_raw.get("opex_items", [])],
        other_items=[normalize_line_item(item, "other") for item in pnl_raw.get("other_items", [])],
        monthly_trend=pnl_raw.get("monthly_trend", [])
    )
    if "total_revenue" in pnl_raw and pnl_raw["total_revenue"]:
        pnl.total_revenue = normalize_line_item(pnl_raw["total_revenue"], "revenue")
    if "gross_profit" in pnl_raw and pnl_raw["gross_profit"]:
        pnl.gross_profit = normalize_line_item(pnl_raw["gross_profit"], "profit")
    if "total_opex" in pnl_raw and pnl_raw["total_opex"]:
        pnl.total_opex = normalize_line_item(pnl_raw["total_opex"], "opex")
    if "operating_profit" in pnl_raw and pnl_raw["operating_profit"]:
        pnl.operating_profit = normalize_line_item(pnl_raw["operating_profit"], "profit")

    # Balance Sheet
    bs_raw = payload.get("balance_sheet", {})
    bs = BalanceSheetData(
        current_assets=[normalize_line_item(item, "current_asset") for item in bs_raw.get("current_assets", [])],
        non_current_assets=[normalize_line_item(item, "non_current_asset") for item in bs_raw.get("non_current_assets", [])],
        current_liabilities=[normalize_line_item(item, "current_liability") for item in bs_raw.get("current_liabilities", [])],
        non_current_liabilities=[normalize_line_item(item, "non_current_liability") for item in bs_raw.get("non_current_liabilities", [])],
        equity_items=[normalize_line_item(item, "equity") for item in bs_raw.get("equity_items", [])]
    )
    if "cash_and_equivalents" in bs_raw and bs_raw["cash_and_equivalents"]:
        bs.cash_and_equivalents = normalize_line_item(bs_raw["cash_and_equivalents"], "current_asset")
    elif bs.current_assets:
        for ca in bs.current_assets:
            if "cash" in ca.key or "cash" in ca.label.lower():
                bs.cash_and_equivalents = ca
                break

    # Cash Flow
    cf_raw = payload.get("cash_flow", {})
    cf = CashFlowData(
        operating_activities=[normalize_line_item(item, "operating") for item in cf_raw.get("operating_activities", [])],
        investing_activities=[normalize_line_item(item, "investing") for item in cf_raw.get("investing_activities", [])],
        financing_activities=[normalize_line_item(item, "financing") for item in cf_raw.get("financing_activities", [])],
        cash_burn_rate=to_decimal(cf_raw.get("cash_burn_rate")),
        cash_runway_months=to_decimal(cf_raw.get("cash_runway_months"))
    )

    # KPIs
    kpis = [normalize_kpi(item) for item in payload.get("kpis", [])]

    # Risks & Actions
    risks = [normalize_risk(item) for item in payload.get("risks", [])]
    actions = [normalize_action(item) for item in payload.get("actions", [])]

    context = ReportContext(
        organisation=org,
        period=period,
        income_statement=pnl,
        balance_sheet=bs,
        cash_flow=cf,
        kpis=kpis,
        risks=risks,
        actions=actions,
        metadata=payload.get("metadata", {})
    )

    return context
