"""
Canonical Reporting Schema for Financial Data.
Typed data structures for financial facts, statements, KPIs, risks, and actions.
All monetary values use Decimal. Missing values are represented as None (never silently 0).
Every material figure tracks its source provenance and fact ID.
"""
from dataclasses import dataclass, field
from decimal import Decimal
from typing import Optional, List, Dict, Any
from datetime import date, datetime


@dataclass
class OrganisationData:
    id: Optional[int] = None
    name: str = "Apex Technologies Holdings Ltd"
    legal_entity_number: Optional[str] = None
    industry: str = "Technology & Software"
    currency_code: str = "USD"
    currency_symbol: str = "$"
    fiscal_year_end_month: int = 12


@dataclass
class ReportingPeriodData:
    id: Optional[int] = None
    label: str = "Q3 2026"
    period_type: str = "quarter"  # 'month', 'quarter', 'year'
    start_date: Optional[date] = None
    end_date: Optional[date] = None
    prior_period_label: Optional[str] = "Q2 2026"
    budget_period_label: Optional[str] = "Q3 2026 Budget"
    is_closed: bool = True


@dataclass
class FinancialLineItem:
    key: str
    label: str
    actual: Optional[Decimal] = None
    budget: Optional[Decimal] = None
    prior: Optional[Decimal] = None
    ytd_actual: Optional[Decimal] = None
    ytd_budget: Optional[Decimal] = None
    category: str = "general"
    is_subtotal: bool = False
    is_total: bool = False
    notes: Optional[str] = None
    source_ref: Optional[str] = None
    fact_id: Optional[str] = None

    @property
    def variance_vs_budget(self) -> Optional[Decimal]:
        if self.actual is not None and self.budget is not None:
            return self.actual - self.budget
        return None

    @property
    def variance_pct_vs_budget(self) -> Optional[Decimal]:
        if self.actual is not None and self.budget is not None and self.budget != Decimal("0"):
            return ((self.actual - self.budget) / abs(self.budget)) * Decimal("100")
        return None

    @property
    def variance_vs_prior(self) -> Optional[Decimal]:
        if self.actual is not None and self.prior is not None:
            return self.actual - self.prior
        return None

    @property
    def variance_pct_vs_prior(self) -> Optional[Decimal]:
        if self.actual is not None and self.prior is not None and self.prior != Decimal("0"):
            return ((self.actual - self.prior) / abs(self.prior)) * Decimal("100")
        return None


@dataclass
class IncomeStatementData:
    revenue_items: List[FinancialLineItem] = field(default_factory=list)
    cogs_items: List[FinancialLineItem] = field(default_factory=list)
    opex_items: List[FinancialLineItem] = field(default_factory=list)
    other_items: List[FinancialLineItem] = field(default_factory=list)

    # Calculated / Authoritative Totals
    total_revenue: Optional[FinancialLineItem] = None
    total_cogs: Optional[FinancialLineItem] = None
    gross_profit: Optional[FinancialLineItem] = None
    total_opex: Optional[FinancialLineItem] = None
    operating_profit: Optional[FinancialLineItem] = None  # EBIT
    ebitda: Optional[FinancialLineItem] = None
    net_profit: Optional[FinancialLineItem] = None

    # Historical monthly/quarterly series for charting
    monthly_trend: List[Dict[str, Any]] = field(default_factory=list)


@dataclass
class BalanceSheetData:
    current_assets: List[FinancialLineItem] = field(default_factory=list)
    non_current_assets: List[FinancialLineItem] = field(default_factory=list)
    current_liabilities: List[FinancialLineItem] = field(default_factory=list)
    non_current_liabilities: List[FinancialLineItem] = field(default_factory=list)
    equity_items: List[FinancialLineItem] = field(default_factory=list)

    total_current_assets: Optional[FinancialLineItem] = None
    total_non_current_assets: Optional[FinancialLineItem] = None
    total_assets: Optional[FinancialLineItem] = None
    total_current_liabilities: Optional[FinancialLineItem] = None
    total_non_current_liabilities: Optional[FinancialLineItem] = None
    total_liabilities: Optional[FinancialLineItem] = None
    total_equity: Optional[FinancialLineItem] = None
    total_liabilities_and_equity: Optional[FinancialLineItem] = None
    cash_and_equivalents: Optional[FinancialLineItem] = None
    net_working_capital: Optional[FinancialLineItem] = None


@dataclass
class CashFlowData:
    operating_activities: List[FinancialLineItem] = field(default_factory=list)
    investing_activities: List[FinancialLineItem] = field(default_factory=list)
    financing_activities: List[FinancialLineItem] = field(default_factory=list)

    net_cash_operating: Optional[FinancialLineItem] = None
    net_cash_investing: Optional[FinancialLineItem] = None
    net_cash_financing: Optional[FinancialLineItem] = None
    net_change_in_cash: Optional[FinancialLineItem] = None
    beginning_cash: Optional[FinancialLineItem] = None
    ending_cash: Optional[FinancialLineItem] = None
    cash_burn_rate: Optional[Decimal] = None
    cash_runway_months: Optional[Decimal] = None


@dataclass
class KPIData:
    key: str
    name: str
    category: str = "financial"  # e.g., 'financial', 'operational', 'customer', 'efficiency'
    current_value: Optional[Decimal] = None
    target_value: Optional[Decimal] = None
    prior_value: Optional[Decimal] = None
    unit: str = "$"  # '$', '%', 'x', 'count', 'days'
    definition: str = ""
    higher_is_better: bool = True
    notes: Optional[str] = None
    fact_id: Optional[str] = None
    monthly_history: List[Dict[str, Any]] = field(default_factory=list)

    @property
    def variance_vs_target(self) -> Optional[Decimal]:
        if self.current_value is not None and self.target_value is not None:
            return self.current_value - self.target_value
        return None

    @property
    def status(self) -> str:
        """Returns 'green', 'amber', or 'red' based on target variance."""
        if self.current_value is None or self.target_value is None:
            return "neutral"
        diff = self.current_value - self.target_value
        if not self.higher_is_better:
            diff = -diff
        
        pct_diff = (diff / abs(self.target_value)) * Decimal("100") if self.target_value != Decimal("0") else Decimal("0")
        if pct_diff >= Decimal("0"):
            return "green"
        elif pct_diff >= Decimal("-5.0"):
            return "amber"
        else:
            return "red"


@dataclass
class VarianceFinding:
    id: str
    line_item_key: str
    line_item_name: str
    category: str  # 'revenue', 'cogs', 'opex', 'kpi', 'cash'
    comparison: str  # 'vs_budget', 'vs_prior'
    actual_value: Optional[Decimal]
    baseline_value: Optional[Decimal]
    variance_amount: Optional[Decimal]
    variance_pct: Optional[Decimal]
    is_favorable: bool
    is_material: bool  # Exceeds threshold (e.g., >5% and >$25k)
    headline: str
    description: str
    requires_investigation: bool = False


@dataclass
class RiskItem:
    id: str
    title: str
    category: str  # 'Financial', 'Operational', 'Compliance', 'Market'
    severity: str  # 'High', 'Medium', 'Low'
    likelihood: str  # 'High', 'Medium', 'Low'
    description: str
    mitigation_plan: str
    owner: Optional[str] = None
    source_finding_id: Optional[str] = None


@dataclass
class ActionItem:
    id: str
    title: str
    reason: str
    priority: str  # 'Urgent', 'High', 'Medium', 'Low'
    owner: Optional[str] = None
    due_date: Optional[str] = None
    status: str = "Open"  # 'Open', 'In Progress', 'Completed'
    source_ref: Optional[str] = None


@dataclass
class ReportContext:
    organisation: OrganisationData
    period: ReportingPeriodData
    income_statement: IncomeStatementData
    balance_sheet: BalanceSheetData
    cash_flow: CashFlowData
    kpis: List[KPIData] = field(default_factory=list)
    findings: List[VarianceFinding] = field(default_factory=list)
    risks: List[RiskItem] = field(default_factory=list)
    actions: List[ActionItem] = field(default_factory=list)
    metadata: Dict[str, Any] = field(default_factory=dict)
    all_facts: Dict[str, Any] = field(default_factory=dict)

    def register_fact(self, fact_id: str, value: Any, label: str, source: str = "calculated"):
        self.all_facts[fact_id] = {
            "fact_id": fact_id,
            "value": value,
            "label": label,
            "source": source
        }
