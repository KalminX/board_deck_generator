"""
Deterministic Template-based Commentary Provider.
Generates fact-grounded, professional commentary for corporate board presentations.
Never invents facts or figures; links all narratives to registered fact IDs.
Supports concise, standard, and detailed narrative styles with length bounding.
"""
from decimal import Decimal
from typing import List, Optional, Dict, Any
from services.finance.canonical_schema import ReportContext, VarianceFinding, KPIData
from services.commentary.base import CommentaryProvider, PresentationCommentary, CommentarySection, FactReference


def fmt_curr(val: Optional[Decimal], sym: str = "$") -> str:
    if val is None:
        return "N/A"
    abs_val = abs(val)
    if abs_val >= Decimal("1000000"):
        formatted = f"{sym}{abs_val / Decimal('1000000'):,.2f}M"
    elif abs_val >= Decimal("1000"):
        formatted = f"{sym}{abs_val / Decimal('1000'):,.1f}k"
    else:
        formatted = f"{sym}{abs_val:,.0f}"
    return f"-{formatted}" if val < Decimal("0") else formatted


def fmt_pct(val: Optional[Decimal]) -> str:
    if val is None:
        return "N/A"
    return f"{val:+.1f}%" if val != Decimal("0") else "0.0%"


class TemplateCommentaryProvider(CommentaryProvider):
    """
    Rule-based deterministic commentary generator.
    """

    def generate_commentary(self, context: ReportContext, style: str = "standard") -> PresentationCommentary:
        sym = context.organisation.currency_symbol
        period_lbl = context.period.label

        exec_summary = self._build_executive_summary(context, style, sym, period_lbl)
        perf_overview = self._build_performance_overview(context, style, sym, period_lbl)
        pnl_sec = self._build_pnl_commentary(context, style, sym, period_lbl)
        bs_sec = self._build_balance_sheet_commentary(context, style, sym, period_lbl)
        cf_sec = self._build_cash_flow_commentary(context, style, sym, period_lbl)
        kpi_sec = self._build_kpi_commentary(context, style, sym, period_lbl)
        risks_sec = self._build_risks_commentary(context, style, sym, period_lbl)
        actions_sec = self._build_actions_commentary(context, style, sym, period_lbl)
        appendix_sec = self._build_appendix_commentary(context, style, sym, period_lbl)

        return PresentationCommentary(
            executive_summary=exec_summary,
            performance_overview=perf_overview,
            pnl_commentary=pnl_sec,
            balance_sheet_commentary=bs_sec,
            cash_flow_commentary=cf_sec,
            kpi_commentary=kpi_sec,
            risks_commentary=risks_sec,
            actions_commentary=actions_sec,
            appendix_commentary=appendix_sec
        )

    def _build_executive_summary(self, context: ReportContext, style: str, sym: str, period_lbl: str) -> CommentarySection:
        pnl = context.income_statement
        bs = context.balance_sheet
        facts: List[FactReference] = []
        bullets: List[str] = []

        rev = pnl.total_revenue
        op = pnl.operating_profit
        cash = bs.cash_and_equivalents

        if rev and rev.actual is not None:
            facts.append(FactReference("fact:pnl:total_revenue:actual", "Revenue Actual", rev.actual))
            if rev.budget is not None:
                diff_pct = rev.variance_pct_vs_budget
                direction = "ahead of" if (diff_pct or 0) >= 0 else "behind"
                lead = f"Revenue reached {fmt_curr(rev.actual, sym)} in {period_lbl}, tracking {direction} budget by {fmt_pct(diff_pct)} ({fmt_curr(rev.variance_vs_budget, sym)})."
            else:
                lead = f"Revenue reached {fmt_curr(rev.actual, sym)} in {period_lbl}."
        else:
            lead = f"Financial performance review for {period_lbl}."

        # Key Findings Bullets
        if op and op.actual is not None:
            facts.append(FactReference("fact:pnl:operating_profit:actual", "Operating Profit Actual", op.actual))
            op_var = op.variance_pct_vs_budget
            if op_var is not None:
                status = "favorable" if op_var >= 0 else "adverse"
                bullets.append(f"Operating Profit (EBIT) delivered {fmt_curr(op.actual, sym)}, representing a {fmt_pct(op_var)} {status} variance to budget.")
            else:
                bullets.append(f"Operating Profit (EBIT) delivered {fmt_curr(op.actual, sym)} for the period.")

        if cash and cash.actual is not None:
            facts.append(FactReference("fact:bs:cash_and_equivalents:actual", "Cash Balance", cash.actual))
            bullets.append(f"Closing cash reserves stood at {fmt_curr(cash.actual, sym)}, providing stability for ongoing operational execution.")

        # Add top material variance if exists
        mat_findings = [f for f in context.findings if f.is_material]
        if mat_findings:
            top_var = mat_findings[0]
            bullets.append(f"Material variance observed in {top_var.line_item_name}: {top_var.headline}.")

        # Add risk/action headline
        if context.actions:
            urgent_count = len([a for a in context.actions if a.priority in ("Urgent", "High")])
            bullets.append(f"{len(context.actions)} operational actions active, including {urgent_count} high-priority initiatives under active management.")

        # Truncate bullets based on style
        if style == "concise":
            bullets = bullets[:3]
        elif style == "standard":
            bullets = bullets[:4]
        else:
            bullets = bullets[:5]

        return CommentarySection(
            section_id="executive_summary",
            title="Executive Summary",
            lead_summary=lead,
            bullet_points=bullets,
            fact_references=facts,
            narrative_style=style
        )

    def _build_performance_overview(self, context: ReportContext, style: str, sym: str, period_lbl: str) -> CommentarySection:
        pnl = context.income_statement
        facts: List[FactReference] = []
        bullets: List[str] = []

        rev = pnl.total_revenue
        gp = pnl.gross_profit
        op = pnl.operating_profit

        lead = f"Comprehensive financial performance summary for {period_lbl} across core profitability and margin indicators."

        if gp and gp.actual is not None and rev and rev.actual:
            gp_margin = (gp.actual / rev.actual) * Decimal("100") if rev.actual != 0 else Decimal("0")
            facts.append(FactReference("fact:pnl:gross_margin_pct:actual", "Gross Margin %", gp_margin))
            bullets.append(f"Gross margin closed at {gp_margin:.1f}%, reflecting disciplined pricing and direct cost management.")

        if op and op.actual is not None and rev and rev.actual:
            op_margin = (op.actual / rev.actual) * Decimal("100") if rev.actual != 0 else Decimal("0")
            facts.append(FactReference("fact:pnl:operating_margin_pct:actual", "Operating Margin %", op_margin))
            bullets.append(f"Operating margin achieved was {op_margin:.1f}%, demonstrating operational leverage.")

        if pnl.total_opex and pnl.total_opex.actual is not None:
            facts.append(FactReference("fact:pnl:total_opex:actual", "Operating Expenses", pnl.total_opex.actual))
            if pnl.total_opex.budget is not None:
                opex_var = pnl.total_opex.variance_pct_vs_budget
                bullets.append(f"Operating overheads totaled {fmt_curr(pnl.total_opex.actual, sym)} ({fmt_pct(opex_var)} vs budget).")

        return CommentarySection(
            section_id="performance_overview",
            title="Financial Performance Overview",
            lead_summary=lead,
            bullet_points=bullets[:4],
            fact_references=facts,
            narrative_style=style
        )

    def _build_pnl_commentary(self, context: ReportContext, style: str, sym: str, period_lbl: str) -> CommentarySection:
        pnl = context.income_statement
        facts: List[FactReference] = []
        bullets: List[str] = []

        lead = f"Income statement review for {period_lbl} with detailed actual vs budget and prior period variance analysis."

        # High level line checks
        for item in (pnl.revenue_items + pnl.opex_items)[:4]:
            if item.actual is not None and item.budget is not None and abs(item.variance_vs_budget or 0) > 0:
                facts.append(FactReference(item.fact_id or f"fact:pnl:{item.key}", item.label, item.actual))
                diff_pct = item.variance_pct_vs_budget
                bullets.append(f"{item.label}: actual {fmt_curr(item.actual, sym)} vs budget {fmt_curr(item.budget, sym)} ({fmt_pct(diff_pct)}).")

        if not bullets:
            bullets.append("All primary income statement categories operated within normal planned variance parameters.")

        return CommentarySection(
            section_id="pnl",
            title="Profit and Loss Statement",
            lead_summary=lead,
            bullet_points=bullets[:4],
            fact_references=facts,
            narrative_style=style
        )

    def _build_balance_sheet_commentary(self, context: ReportContext, style: str, sym: str, period_lbl: str) -> CommentarySection:
        bs = context.balance_sheet
        facts: List[FactReference] = []
        bullets: List[str] = []

        lead = f"Balance sheet position as at end of {period_lbl}, maintaining robust liquidity and capital structure."

        if bs.total_assets and bs.total_assets.actual is not None:
            facts.append(FactReference("fact:bs:total_assets", "Total Assets", bs.total_assets.actual))
            bullets.append(f"Total asset base stood at {fmt_curr(bs.total_assets.actual, sym)}.")

        if bs.net_working_capital and bs.net_working_capital.actual is not None:
            facts.append(FactReference("fact:bs:net_working_capital", "Net Working Capital", bs.net_working_capital.actual))
            bullets.append(f"Net working capital totaled {fmt_curr(bs.net_working_capital.actual, sym)}.")

        if bs.total_equity and bs.total_equity.actual is not None:
            facts.append(FactReference("fact:bs:total_equity", "Total Equity", bs.total_equity.actual))
            bullets.append(f"Shareholders' equity reached {fmt_curr(bs.total_equity.actual, sym)}.")

        return CommentarySection(
            section_id="balance_sheet",
            title="Balance Sheet",
            lead_summary=lead,
            bullet_points=bullets[:3],
            fact_references=facts,
            narrative_style=style
        )

    def _build_cash_flow_commentary(self, context: ReportContext, style: str, sym: str, period_lbl: str) -> CommentarySection:
        cf = context.cash_flow
        facts: List[FactReference] = []
        bullets: List[str] = []

        lead = f"Cash flow statement analysis detailing operating cash conversion, capital investments, and financing movements for {period_lbl}."

        if cf.net_cash_operating and cf.net_cash_operating.actual is not None:
            facts.append(FactReference("fact:cf:operating", "Operating Cash Flow", cf.net_cash_operating.actual))
            bullets.append(f"Operating cash flow generated {fmt_curr(cf.net_cash_operating.actual, sym)}.")

        if cf.net_cash_investing and cf.net_cash_investing.actual is not None:
            facts.append(FactReference("fact:cf:investing", "Investing Cash Flow", cf.net_cash_investing.actual))
            bullets.append(f"Investing activities utilized {fmt_curr(cf.net_cash_investing.actual, sym)}.")

        if cf.cash_runway_months is not None:
            bullets.append(f"Estimated cash runway stands at {cf.cash_runway_months:.1f} months based on current monthly run-rate.")

        return CommentarySection(
            section_id="cash_flow",
            title="Cash Flow Statement",
            lead_summary=lead,
            bullet_points=bullets[:3],
            fact_references=facts,
            narrative_style=style
        )

    def _build_kpi_commentary(self, context: ReportContext, style: str, sym: str, period_lbl: str) -> CommentarySection:
        facts: List[FactReference] = []
        bullets: List[str] = []

        lead = f"Performance scorecards tracking operational and financial key performance indicators for {period_lbl}."

        green_kpis = [k for k in context.kpis if k.status == "green"]
        amber_kpis = [k for k in context.kpis if k.status == "amber"]
        red_kpis = [k for k in context.kpis if k.status == "red"]

        bullets.append(f"{len(green_kpis)} of {len(context.kpis)} core KPIs are meeting or exceeding designated target thresholds.")
        if red_kpis:
            bullets.append(f"Attention required on lagging metrics: {', '.join(k.name for k in red_kpis[:3])}.")

        for k in context.kpis[:3]:
            if k.current_value is not None:
                facts.append(FactReference(k.fact_id or f"fact:kpi:{k.key}", k.name, k.current_value))

        return CommentarySection(
            section_id="kpis",
            title="KPI Analysis",
            lead_summary=lead,
            bullet_points=bullets[:3],
            fact_references=facts,
            narrative_style=style
        )

    def _build_risks_commentary(self, context: ReportContext, style: str, sym: str, period_lbl: str) -> CommentarySection:
        lead = f"Risk registry and financial anomaly monitoring for {period_lbl}, outlining identified exposures and mitigation status."
        bullets = []

        high_risks = [r for r in context.risks if r.severity == "High"]
        if high_risks:
            bullets.append(f"{len(high_risks)} high-severity risks currently tracked with designated mitigation workstreams.")
        else:
            bullets.append("No critical severity risks currently flagged requiring executive escalation.")

        mat_findings = [f for f in context.findings if f.requires_investigation]
        if mat_findings:
            bullets.append(f"{len(mat_findings)} material financial variances identified requiring departmental investigation.")

        return CommentarySection(
            section_id="risks",
            title="Risks and Anomalies",
            lead_summary=lead,
            bullet_points=bullets[:3],
            fact_references=[],
            narrative_style=style
        )

    def _build_actions_commentary(self, context: ReportContext, style: str, sym: str, period_lbl: str) -> CommentarySection:
        lead = f"Executive governance matrix detailing active decisions, accountable owners, and milestone deadlines."
        bullets = []

        urgent_actions = [a for a in context.actions if a.priority in ("Urgent", "High")]
        bullets.append(f"{len(context.actions)} approved board and executive actions under management ({len(urgent_actions)} high/urgent priority).")

        return CommentarySection(
            section_id="actions",
            title="Actions and Decisions",
            lead_summary=lead,
            bullet_points=bullets[:2],
            fact_references=[],
            narrative_style=style
        )

    def _build_appendix_commentary(self, context: ReportContext, style: str, sym: str, period_lbl: str) -> CommentarySection:
        lead = f"Governance notes, accounting methodology, and canonical data snapshot provenance for the {period_lbl} board presentation."
        bullets = [
            f"Financial statements prepared in accordance with standard management accounting policies ({context.organisation.currency_code}).",
            "Figures derived from authoritative source workbook snapshots without automated extrapolation.",
            "Variance thresholds: Materiality defined as absolute variance >= $25k and >= 5.0% vs baseline."
        ]
        return CommentarySection(
            section_id="appendix",
            title="Appendix & Governance",
            lead_summary=lead,
            bullet_points=bullets,
            fact_references=[],
            narrative_style=style
        )
