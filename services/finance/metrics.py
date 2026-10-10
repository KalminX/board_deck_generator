"""
Authoritative financial metrics and variance analysis engine.
Performs deterministic calculations using Decimal precision.
Identifies material variances, anomalies, and registers facts.
"""
from decimal import Decimal, ROUND_HALF_UP
from typing import List, Optional, Tuple, Dict, Any
from services.finance.canonical_schema import (
    IncomeStatementData, BalanceSheetData, CashFlowData, FinancialLineItem,
    KPIData, VarianceFinding, ReportContext, RiskItem, ActionItem
)


def safe_divide(num: Optional[Decimal], denom: Optional[Decimal]) -> Optional[Decimal]:
    if num is None or denom is None or denom == Decimal("0"):
        return None
    return num / denom


def safe_pct(num: Optional[Decimal], denom: Optional[Decimal]) -> Optional[Decimal]:
    val = safe_divide(num, denom)
    if val is not None:
        return val * Decimal("100")
    return None


def calculate_financial_metrics(context: ReportContext) -> ReportContext:
    """
    Computes summary lines for Income Statement, Balance Sheet, Cash Flow,
    populates KPIs, and detects material variances and risk indicators.
    """
    pnl = context.income_statement
    bs = context.balance_sheet
    cf = context.cash_flow

    # 1. P&L Aggregations
    # Total Revenue
    if pnl.revenue_items and pnl.total_revenue is None:
        actual_rev = sum((item.actual for item in pnl.revenue_items if item.actual is not None), Decimal("0"))
        budget_rev = sum((item.budget for item in pnl.revenue_items if item.budget is not None), Decimal("0"))
        prior_rev = sum((item.prior for item in pnl.revenue_items if item.prior is not None), Decimal("0"))
        ytd_rev = sum((item.ytd_actual for item in pnl.revenue_items if item.ytd_actual is not None), Decimal("0"))
        pnl.total_revenue = FinancialLineItem(
            key="total_revenue", label="Total Revenue",
            actual=actual_rev, budget=budget_rev, prior=prior_rev, ytd_actual=ytd_rev,
            is_total=True, fact_id="fact:pnl:total_revenue"
        )
    elif pnl.total_revenue and not pnl.total_revenue.fact_id:
        pnl.total_revenue.fact_id = "fact:pnl:total_revenue"

    # Total COGS
    if pnl.cogs_items and pnl.total_cogs is None:
        actual_cogs = sum((item.actual for item in pnl.cogs_items if item.actual is not None), Decimal("0"))
        budget_cogs = sum((item.budget for item in pnl.cogs_items if item.budget is not None), Decimal("0"))
        prior_cogs = sum((item.prior for item in pnl.cogs_items if item.prior is not None), Decimal("0"))
        ytd_cogs = sum((item.ytd_actual for item in pnl.cogs_items if item.ytd_actual is not None), Decimal("0"))
        pnl.total_cogs = FinancialLineItem(
            key="total_cogs", label="Total Cost of Sales",
            actual=actual_cogs, budget=budget_cogs, prior=prior_cogs, ytd_actual=ytd_cogs,
            is_subtotal=True, fact_id="fact:pnl:total_cogs"
        )
    elif pnl.total_cogs and not pnl.total_cogs.fact_id:
        pnl.total_cogs.fact_id = "fact:pnl:total_cogs"

    # Gross Profit = Revenue - COGS
    if pnl.total_revenue and pnl.total_cogs and pnl.gross_profit is None:
        gp_actual = pnl.total_revenue.actual - pnl.total_cogs.actual if pnl.total_revenue.actual is not None and pnl.total_cogs.actual is not None else None
        gp_budget = pnl.total_revenue.budget - pnl.total_cogs.budget if pnl.total_revenue.budget is not None and pnl.total_cogs.budget is not None else None
        gp_prior = pnl.total_revenue.prior - pnl.total_cogs.prior if pnl.total_revenue.prior is not None and pnl.total_cogs.prior is not None else None
        gp_ytd = pnl.total_revenue.ytd_actual - pnl.total_cogs.ytd_actual if pnl.total_revenue.ytd_actual is not None and pnl.total_cogs.ytd_actual is not None else None
        pnl.gross_profit = FinancialLineItem(
            key="gross_profit", label="Gross Profit",
            actual=gp_actual, budget=gp_budget, prior=gp_prior, ytd_actual=gp_ytd,
            is_total=True, fact_id="fact:pnl:gross_profit"
        )
    elif pnl.gross_profit and not pnl.gross_profit.fact_id:
        pnl.gross_profit.fact_id = "fact:pnl:gross_profit"

    # Total Operating Expenses
    if pnl.opex_items and pnl.total_opex is None:
        actual_opex = sum((item.actual for item in pnl.opex_items if item.actual is not None), Decimal("0"))
        budget_opex = sum((item.budget for item in pnl.opex_items if item.budget is not None), Decimal("0"))
        prior_opex = sum((item.prior for item in pnl.opex_items if item.prior is not None), Decimal("0"))
        ytd_opex = sum((item.ytd_actual for item in pnl.opex_items if item.ytd_actual is not None), Decimal("0"))
        pnl.total_opex = FinancialLineItem(
            key="total_opex", label="Total Operating Expenses",
            actual=actual_opex, budget=budget_opex, prior=prior_opex, ytd_actual=ytd_opex,
            is_subtotal=True, fact_id="fact:pnl:total_opex"
        )
    elif pnl.total_opex and not pnl.total_opex.fact_id:
        pnl.total_opex.fact_id = "fact:pnl:total_opex"

    # Operating Profit (EBIT) = Gross Profit - Total Opex
    if pnl.gross_profit and pnl.total_opex and pnl.operating_profit is None:
        ebit_actual = pnl.gross_profit.actual - pnl.total_opex.actual if pnl.gross_profit.actual is not None and pnl.total_opex.actual is not None else None
        ebit_budget = pnl.gross_profit.budget - pnl.total_opex.budget if pnl.gross_profit.budget is not None and pnl.total_opex.budget is not None else None
        ebit_prior = pnl.gross_profit.prior - pnl.total_opex.prior if pnl.gross_profit.prior is not None and pnl.total_opex.prior is not None else None
        ebit_ytd = pnl.gross_profit.ytd_actual - pnl.total_opex.ytd_actual if pnl.gross_profit.ytd_actual is not None and pnl.total_opex.ytd_actual is not None else None
        pnl.operating_profit = FinancialLineItem(
            key="operating_profit", label="Operating Profit (EBIT)",
            actual=ebit_actual, budget=ebit_budget, prior=ebit_prior, ytd_actual=ebit_ytd,
            is_total=True, fact_id="fact:pnl:operating_profit"
        )
    elif pnl.operating_profit and not pnl.operating_profit.fact_id:
        pnl.operating_profit.fact_id = "fact:pnl:operating_profit"

    # EBITDA
    if pnl.operating_profit and pnl.ebitda is None:
        pnl.ebitda = FinancialLineItem(
            key="ebitda", label="EBITDA",
            actual=pnl.operating_profit.actual,
            budget=pnl.operating_profit.budget,
            prior=pnl.operating_profit.prior,
            ytd_actual=pnl.operating_profit.ytd_actual,
            is_total=True, fact_id="fact:pnl:ebitda"
        )

    # Net Profit
    if pnl.operating_profit and pnl.net_profit is None:
        pnl.net_profit = FinancialLineItem(
            key="net_profit", label="Net Profit",
            actual=pnl.operating_profit.actual,
            budget=pnl.operating_profit.budget,
            prior=pnl.operating_profit.prior,
            ytd_actual=pnl.operating_profit.ytd_actual,
            is_total=True, fact_id="fact:pnl:net_profit"
        )

    # 2. Balance Sheet Calculations
    if bs.current_assets and bs.total_current_assets is None:
        bs.total_current_assets = FinancialLineItem(
            key="total_current_assets", label="Total Current Assets",
            actual=sum((i.actual for i in bs.current_assets if i.actual is not None), Decimal("0")),
            prior=sum((i.prior for i in bs.current_assets if i.prior is not None), Decimal("0")),
            is_subtotal=True, fact_id="fact:bs:total_current_assets"
        )

    if bs.non_current_assets and bs.total_non_current_assets is None:
        bs.total_non_current_assets = FinancialLineItem(
            key="total_non_current_assets", label="Total Non-Current Assets",
            actual=sum((i.actual for i in bs.non_current_assets if i.actual is not None), Decimal("0")),
            prior=sum((i.prior for i in bs.non_current_assets if i.prior is not None), Decimal("0")),
            is_subtotal=True, fact_id="fact:bs:total_non_current_assets"
        )

    if bs.total_current_assets and bs.total_non_current_assets and bs.total_assets is None:
        bs.total_assets = FinancialLineItem(
            key="total_assets", label="Total Assets",
            actual=bs.total_current_assets.actual + bs.total_non_current_assets.actual,
            prior=(bs.total_current_assets.prior or Decimal("0")) + (bs.total_non_current_assets.prior or Decimal("0")),
            is_total=True, fact_id="fact:bs:total_assets"
        )

    if bs.current_liabilities and bs.total_current_liabilities is None:
        bs.total_current_liabilities = FinancialLineItem(
            key="total_current_liabilities", label="Total Current Liabilities",
            actual=sum((i.actual for i in bs.current_liabilities if i.actual is not None), Decimal("0")),
            prior=sum((i.prior for i in bs.current_liabilities if i.prior is not None), Decimal("0")),
            is_subtotal=True, fact_id="fact:bs:total_current_liabilities"
        )

    if bs.non_current_liabilities and bs.total_non_current_liabilities is None:
        bs.total_non_current_liabilities = FinancialLineItem(
            key="total_non_current_liabilities", label="Total Non-Current Liabilities",
            actual=sum((i.actual for i in bs.non_current_liabilities if i.actual is not None), Decimal("0")),
            prior=sum((i.prior for i in bs.non_current_liabilities if i.prior is not None), Decimal("0")),
            is_subtotal=True, fact_id="fact:bs:total_non_current_liabilities"
        )

    if bs.total_current_liabilities and bs.total_liabilities is None:
        non_curr = bs.total_non_current_liabilities.actual if bs.total_non_current_liabilities else Decimal("0")
        bs.total_liabilities = FinancialLineItem(
            key="total_liabilities", label="Total Liabilities",
            actual=bs.total_current_liabilities.actual + non_curr,
            prior=(bs.total_current_liabilities.prior or Decimal("0")) + ((bs.total_non_current_liabilities.prior or Decimal("0")) if bs.total_non_current_liabilities else Decimal("0")),
            is_subtotal=True, fact_id="fact:bs:total_liabilities"
        )

    if bs.equity_items and bs.total_equity is None:
        bs.total_equity = FinancialLineItem(
            key="total_equity", label="Total Equity",
            actual=sum((i.actual for i in bs.equity_items if i.actual is not None), Decimal("0")),
            prior=sum((i.prior for i in bs.equity_items if i.prior is not None), Decimal("0")),
            is_total=True, fact_id="fact:bs:total_equity"
        )

    if bs.total_current_assets and bs.total_current_liabilities and bs.net_working_capital is None:
        bs.net_working_capital = FinancialLineItem(
            key="net_working_capital", label="Net Working Capital",
            actual=bs.total_current_assets.actual - bs.total_current_liabilities.actual,
            prior=(bs.total_current_assets.prior or Decimal("0")) - (bs.total_current_liabilities.prior or Decimal("0")),
            is_total=True, fact_id="fact:bs:net_working_capital"
        )

    # 3. Cash Flow Calculations
    if cf.operating_activities and cf.net_cash_operating is None:
        cf.net_cash_operating = FinancialLineItem(
            key="net_cash_operating", label="Cash Flow from Operating Activities",
            actual=sum((i.actual for i in cf.operating_activities if i.actual is not None), Decimal("0")),
            budget=sum((i.budget for i in cf.operating_activities if i.budget is not None), Decimal("0")),
            prior=sum((i.prior for i in cf.operating_activities if i.prior is not None), Decimal("0")),
            is_subtotal=True, fact_id="fact:cf:net_cash_operating"
        )

    if cf.investing_activities and cf.net_cash_investing is None:
        cf.net_cash_investing = FinancialLineItem(
            key="net_cash_investing", label="Cash Flow from Investing Activities",
            actual=sum((i.actual for i in cf.investing_activities if i.actual is not None), Decimal("0")),
            budget=sum((i.budget for i in cf.investing_activities if i.budget is not None), Decimal("0")),
            prior=sum((i.prior for i in cf.investing_activities if i.prior is not None), Decimal("0")),
            is_subtotal=True, fact_id="fact:cf:net_cash_investing"
        )

    if cf.financing_activities and cf.net_cash_financing is None:
        cf.net_cash_financing = FinancialLineItem(
            key="net_cash_financing", label="Cash Flow from Financing Activities",
            actual=sum((i.actual for i in cf.financing_activities if i.actual is not None), Decimal("0")),
            budget=sum((i.budget for i in cf.financing_activities if i.budget is not None), Decimal("0")),
            prior=sum((i.prior for i in cf.financing_activities if i.prior is not None), Decimal("0")),
            is_subtotal=True, fact_id="fact:cf:net_cash_financing"
        )

    if cf.net_change_in_cash is None and cf.net_cash_operating:
        op = cf.net_cash_operating.actual or Decimal("0")
        inv = cf.net_cash_investing.actual if cf.net_cash_investing else Decimal("0")
        fin = cf.net_cash_financing.actual if cf.net_cash_financing else Decimal("0")
        cf.net_change_in_cash = FinancialLineItem(
            key="net_change_in_cash", label="Net Change in Cash",
            actual=op + inv + fin,
            is_total=True, fact_id="fact:cf:net_change_in_cash"
        )

    # 4. Register All Standard Facts for Provenance
    _register_all_facts(context)

    # 5. Detect Variances and Findings
    context.findings = _detect_variances(context)

    return context


def _register_all_facts(context: ReportContext):
    pnl = context.income_statement
    bs = context.balance_sheet
    cf = context.cash_flow

    if pnl.total_revenue and pnl.total_revenue.actual is not None:
        context.register_fact("fact:pnl:total_revenue:actual", pnl.total_revenue.actual, "Total Revenue Actual")
    if pnl.total_revenue and pnl.total_revenue.budget is not None:
        context.register_fact("fact:pnl:total_revenue:budget", pnl.total_revenue.budget, "Total Revenue Budget")
    if pnl.total_revenue and pnl.total_revenue.prior is not None:
        context.register_fact("fact:pnl:total_revenue:prior", pnl.total_revenue.prior, "Total Revenue Prior")

    if pnl.gross_profit and pnl.gross_profit.actual is not None:
        context.register_fact("fact:pnl:gross_profit:actual", pnl.gross_profit.actual, "Gross Profit Actual")
        gm = safe_pct(pnl.gross_profit.actual, pnl.total_revenue.actual if pnl.total_revenue else None)
        if gm is not None:
            context.register_fact("fact:pnl:gross_margin_pct:actual", gm, "Gross Margin % Actual")

    if pnl.operating_profit and pnl.operating_profit.actual is not None:
        context.register_fact("fact:pnl:operating_profit:actual", pnl.operating_profit.actual, "Operating Profit Actual")
        op_margin = safe_pct(pnl.operating_profit.actual, pnl.total_revenue.actual if pnl.total_revenue else None)
        if op_margin is not None:
            context.register_fact("fact:pnl:operating_margin_pct:actual", op_margin, "Operating Margin % Actual")

    if bs.cash_and_equivalents and bs.cash_and_equivalents.actual is not None:
        context.register_fact("fact:bs:cash_and_equivalents:actual", bs.cash_and_equivalents.actual, "Cash Balance Actual")

    for kpi in context.kpis:
        if kpi.fact_id and kpi.current_value is not None:
            context.register_fact(kpi.fact_id, kpi.current_value, f"KPI: {kpi.name}")


def _detect_variances(context: ReportContext) -> List[VarianceFinding]:
    findings: List[VarianceFinding] = []
    pnl = context.income_statement

    # Check Revenue lines
    all_lines = (
        [(item, "revenue", True) for item in pnl.revenue_items] +
        [(pnl.total_revenue, "revenue", True)] if pnl.total_revenue else []
    ) + (
        [(item, "cogs", False) for item in pnl.cogs_items] +
        [(pnl.total_cogs, "cogs", False)] if pnl.total_cogs else []
    ) + (
        [(item, "opex", False) for item in pnl.opex_items] +
        [(pnl.total_opex, "opex", False)] if pnl.total_opex else []
    ) + (
        [(pnl.operating_profit, "profit", True)] if pnl.operating_profit else []
    )

    for item, category, higher_is_better in all_lines:
        if not item or item.actual is None:
            continue

        # Check vs Budget
        if item.budget is not None and item.budget != Decimal("0"):
            diff = item.actual - item.budget
            diff_pct = (diff / abs(item.budget)) * Decimal("100")
            is_fav = diff >= Decimal("0") if higher_is_better else diff <= Decimal("0")
            is_mat = abs(diff_pct) >= Decimal("5.0") and abs(diff) >= Decimal("10000")

            sym = "+" if diff > 0 else ""
            findings.append(VarianceFinding(
                id=f"finding:var:budget:{item.key}",
                line_item_key=item.key,
                line_item_name=item.label,
                category=category,
                comparison="vs_budget",
                actual_value=item.actual,
                baseline_value=item.budget,
                variance_amount=diff,
                variance_pct=diff_pct,
                is_favorable=is_fav,
                is_material=is_mat,
                headline=f"{item.label} vs Budget: {sym}{diff_pct:.1f}%",
                description=f"{item.label} was {item.actual:,.0f} compared to budget {item.budget:,.0f} ({sym}{diff:,.0f} or {sym}{diff_pct:.1f}%).",
                requires_investigation=not is_fav and is_mat
            ))

        # Check vs Prior Period
        if item.prior is not None and item.prior != Decimal("0"):
            diff = item.actual - item.prior
            diff_pct = (diff / abs(item.prior)) * Decimal("100")
            is_fav = diff >= Decimal("0") if higher_is_better else diff <= Decimal("0")
            is_mat = abs(diff_pct) >= Decimal("5.0") and abs(diff) >= Decimal("10000")

            sym = "+" if diff > 0 else ""
            findings.append(VarianceFinding(
                id=f"finding:var:prior:{item.key}",
                line_item_key=item.key,
                line_item_name=item.label,
                category=category,
                comparison="vs_prior",
                actual_value=item.actual,
                baseline_value=item.prior,
                variance_amount=diff,
                variance_pct=diff_pct,
                is_favorable=is_fav,
                is_material=is_mat,
                headline=f"{item.label} vs Prior: {sym}{diff_pct:.1f}%",
                description=f"{item.label} was {item.actual:,.0f} compared to prior period {item.prior:,.0f} ({sym}{diff:,.0f} or {sym}{diff_pct:.1f}%).",
                requires_investigation=not is_fav and is_mat
            ))

    return findings
