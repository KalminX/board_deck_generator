"""
Management Command to Seed Fictional Enterprise Board Demo Data.
Creates sample organisation, periods, snapshots, pre-generated presentation,
and outputs a canonical Excel file in sample_data/.
"""
import os
import openpyxl
from decimal import Decimal
from django.core.management.base import BaseCommand
from django.core.files.base import ContentFile
from django.conf import settings
from board_deck.models import (
    Organisation, ReportingPeriod, BrandConfiguration,
    NormalizedSnapshot, PresentationDeck, SlideReviewNarrative, AuditLog
)
from services.finance.canonical_schema import ReportContext
from services.finance.normalization import normalize_report_payload
from services.finance.metrics import calculate_financial_metrics
from services.presentations.builder import PresentationBuilder, SLIDE_REGISTRY
from services.presentations.tokens import DEFAULT_BRAND, MODERN_NAVY_BRAND


class Command(BaseCommand):
    help = "Seeds fictional enterprise demo dataset and generates sample board deck and Excel workbook."

    def handle(self, *args, **options):
        self.stdout.write("Seeding fictional demo dataset...")

        # 1. Create Organisation
        org, _ = Organisation.objects.get_or_create(
            name="Apex Technologies Holdings Ltd",
            defaults={
                "legal_entity_number": "SEC-8492041",
                "industry": "B2B SaaS & Cloud Infrastructure",
                "currency_code": "USD",
                "currency_symbol": "$",
                "fiscal_year_end_month": 12
            }
        )

        # 2. Create Reporting Period
        period, _ = ReportingPeriod.objects.get_or_create(
            organisation=org,
            label="Q3 2026",
            defaults={
                "period_type": "quarter",
                "prior_period_label": "Q2 2026",
                "budget_period_label": "Q3 2026 Budget",
                "is_closed": True
            }
        )

        # 3. Create Brand Configurations
        BrandConfiguration.objects.get_or_create(
            name="Corporate Slate",
            defaults={
                "primary_color_hex": "#0F172A",
                "secondary_color_hex": "#334155",
                "accent_color_hex": "#2563EB",
                "font_title": "Calibri",
                "font_body": "Calibri",
                "is_default": True
            }
        )

        # 4. Construct Comprehensive Canonical Financial Dataset
        demo_payload = {
            "organisation": {
                "name": "Apex Technologies Holdings Ltd",
                "legal_entity_number": "SEC-8492041",
                "industry": "B2B SaaS & Cloud Infrastructure",
                "currency_code": "USD",
                "currency_symbol": "$"
            },
            "period": {
                "label": "Q3 2026",
                "period_type": "quarter",
                "prior_period_label": "Q2 2026",
                "budget_period_label": "Q3 2026 Budget"
            },
            "income_statement": {
                "revenue_items": [
                    {"key": "rev_enterprise", "label": "Enterprise Subscriptions (SaaS)", "actual": "2450000", "budget": "2200000", "prior": "2050000", "ytd_actual": "6800000", "category": "revenue"},
                    {"key": "rev_midmarket", "label": "Mid-Market Platform Licenses", "actual": "950000", "budget": "1000000", "prior": "920000", "ytd_actual": "2800000", "category": "revenue"},
                    {"key": "rev_services", "label": "Professional Implementation Services", "actual": "450000", "budget": "400000", "prior": "380000", "ytd_actual": "1250000", "category": "revenue"}
                ],
                "cogs_items": [
                    {"key": "cogs_hosting", "label": "Cloud Hosting & Infrastructure (AWS/GCP)", "actual": "520000", "budget": "480000", "prior": "460000", "ytd_actual": "1480000", "category": "cogs"},
                    {"key": "cogs_support", "label": "Customer Support & Success Team", "actual": "260000", "budget": "250000", "prior": "240000", "ytd_actual": "740000", "category": "cogs"}
                ],
                "opex_items": [
                    {"key": "opex_sales", "label": "Sales & Marketing", "actual": "1150000", "budget": "1100000", "prior": "1020000", "ytd_actual": "3250000", "category": "opex"},
                    {"key": "opex_rd", "label": "Research & Engineering (R&D)", "actual": "890000", "budget": "820000", "prior": "780000", "ytd_actual": "2480000", "category": "opex"},
                    {"key": "opex_ga", "label": "General & Administrative (G&A)", "actual": "390000", "budget": "410000", "prior": "370000", "ytd_actual": "1160000", "category": "opex"}
                ],
                "monthly_trend": [
                    {"period": "Apr 2026", "revenue": "1100000", "opex": "780000", "profit": "160000"},
                    {"period": "May 2026", "revenue": "1120000", "opex": "790000", "profit": "175000"},
                    {"period": "Jun 2026", "revenue": "1130000", "opex": "810000", "profit": "185000"},
                    {"period": "Jul 2026", "revenue": "1240000", "opex": "800000", "profit": "210000"},
                    {"period": "Aug 2026", "revenue": "1280000", "opex": "815000", "profit": "230000"},
                    {"period": "Sep 2026", "revenue": "1330000", "opex": "815000", "profit": "250000"}
                ]
            },
            "balance_sheet": {
                "current_assets": [
                    {"key": "bs_cash", "label": "Cash & Cash Equivalents", "actual": "6420000", "prior": "5890000", "category": "current_asset"},
                    {"key": "bs_ar", "label": "Accounts Receivable (Trade)", "actual": "1450000", "prior": "1320000", "category": "current_asset"},
                    {"key": "bs_prepaid", "label": "Prepaid Expenses & Other", "actual": "280000", "prior": "250000", "category": "current_asset"}
                ],
                "non_current_assets": [
                    {"key": "bs_ppe", "label": "Property, Plant & Equipment", "actual": "650000", "prior": "680000", "category": "non_current_asset"},
                    {"key": "bs_intangibles", "label": "Capitalized Software & IP", "actual": "2100000", "prior": "1950000", "category": "non_current_asset"}
                ],
                "current_liabilities": [
                    {"key": "bs_ap", "label": "Accounts Payable", "actual": "540000", "prior": "490000", "category": "current_liability"},
                    {"key": "bs_deferred_rev", "label": "Deferred Subscription Revenue", "actual": "2850000", "prior": "2600000", "category": "current_liability"},
                    {"key": "bs_accruals", "label": "Accrued Payroll & Taxes", "actual": "420000", "prior": "390000", "category": "current_liability"}
                ],
                "non_current_liabilities": [
                    {"key": "bs_term_debt", "label": "Long-Term Venture Facility", "actual": "1500000", "prior": "1500000", "category": "non_current_liability"}
                ],
                "equity_items": [
                    {"key": "bs_share_capital", "label": "Share Capital (Series B)", "actual": "4500000", "prior": "4500000", "category": "equity"},
                    {"key": "bs_retained_earnings", "label": "Retained Earnings & Reserves", "actual": "1090000", "prior": "610000", "category": "equity"}
                ]
            },
            "cash_flow": {
                "operating_activities": [
                    {"key": "cf_net_income", "label": "Net Operating Profit", "actual": "690000", "budget": "550000", "prior": "520000"},
                    {"key": "cf_depreciation", "label": "Depreciation & Amortization", "actual": "95000", "budget": "90000", "prior": "85000"},
                    {"key": "cf_wc_change", "label": "Change in Net Working Capital", "actual": "125000", "budget": "80000", "prior": "65000"}
                ],
                "investing_activities": [
                    {"key": "cf_capex", "label": "Capitalized R&D & Equipment", "actual": "-220000", "budget": "-200000", "prior": "-180000"}
                ],
                "financing_activities": [
                    {"key": "cf_debt_service", "label": "Debt Service & Interest", "actual": "-30000", "budget": "-30000", "prior": "-30000"}
                ],
                "cash_burn_rate": "0",
                "cash_runway_months": "28.5"
            },
            "kpis": [
                {
                    "key": "arr", "name": "Annual Recurring Revenue (ARR)", "category": "financial",
                    "current_value": "15.4", "target_value": "14.8", "prior_value": "13.2", "unit": "$M",
                    "definition": "Contracted annual recurring subscription run-rate.", "higher_is_better": True,
                    "monthly_history": [
                        {"period": "Apr", "value": 13.2, "target": 13.0},
                        {"period": "May", "value": 13.7, "target": 13.4},
                        {"period": "Jun", "value": 14.1, "target": 13.8},
                        {"period": "Jul", "value": 14.6, "target": 14.2},
                        {"period": "Aug", "value": 15.0, "target": 14.5},
                        {"period": "Sep", "value": 15.4, "target": 14.8}
                    ]
                },
                {"key": "gross_margin", "name": "Gross Margin %", "category": "financial", "current_value": "79.7", "target_value": "78.5", "prior_value": "77.8", "unit": "%", "definition": "Direct revenue gross margin after hosting & support.", "higher_is_better": True},
                {"key": "nrr", "name": "Net Revenue Retention (NRR)", "category": "customer", "current_value": "118.5", "target_value": "115.0", "prior_value": "114.2", "unit": "%", "definition": "Trailing 12-month net expansion across cohort.", "higher_is_better": True},
                {"key": "cac_payback", "name": "CAC Payback Period", "category": "efficiency", "current_value": "13.2", "target_value": "14.0", "prior_value": "14.5", "unit": " mo", "definition": "Gross profit months required to recover sales CAC.", "higher_is_better": False},
                {"key": "rule_40", "name": "Rule of 40 Score", "category": "financial", "current_value": "44.2", "target_value": "40.0", "prior_value": "38.5", "unit": "%", "definition": "ARR Growth Rate (28%) + Free Cash Flow Margin (16.2%).", "higher_is_better": True},
                {"key": "headcount", "name": "Total Headcount (FTEs)", "category": "operational", "current_value": "74", "target_value": "76", "prior_value": "68", "unit": " FTE", "definition": "Full-time equivalent staff across all business units.", "higher_is_better": True}
            ],
            "risks": [
                {
                    "id": "risk_1",
                    "title": "Cloud Infrastructure Cost Inflation",
                    "category": "Financial",
                    "severity": "High",
                    "likelihood": "Medium",
                    "description": "GPU cluster demand for LLM customer features increased hosting burn +8.3% over budget in Q3.",
                    "mitigation_plan": "Engineering is implementing reserved instance commitments and token caching to reduce inference compute.",
                    "owner": "CTO / Head of Infrastructure"
                },
                {
                    "id": "risk_2",
                    "title": "Key Enterprise Deal Cycle Lengthening",
                    "category": "Operational",
                    "severity": "Medium",
                    "likelihood": "Medium",
                    "description": "Procurement security reviews expanded average enterprise closing cycle from 65 to 82 days.",
                    "mitigation_plan": "SOC 2 Type II refresh and automated compliance portal deployed for prospect infosec signoffs.",
                    "owner": "VP Sales"
                }
            ],
            "actions": [
                {
                    "id": "action_1",
                    "title": "Finalize AWS 3-Year Savings Plan Commitment",
                    "reason": "Lock in 38% compute cost reduction on base tier Kubernetes nodes.",
                    "priority": "Urgent",
                    "owner": "Director of DevOps",
                    "due_date": "2026-10-31",
                    "status": "In Progress"
                },
                {
                    "id": "action_2",
                    "title": "Establish European Subsidiary & Data Residency",
                    "reason": "Satisfy GDPR compliance mandates for expanding Tier 1 UK/EU banking pipeline.",
                    "priority": "High",
                    "owner": "General Counsel / CFO",
                    "due_date": "2026-11-15",
                    "status": "Open"
                },
                {
                    "id": "action_3",
                    "title": "Deploy Dynamic In-App Seat Expansion Automation",
                    "reason": "Automate self-serve license tier upgrades to drive NRR expansion.",
                    "priority": "Medium",
                    "owner": "Head of Product",
                    "due_date": "2026-12-15",
                    "status": "In Progress"
                }
            ]
        }

        # 5. Save Snapshot
        snapshot, _ = NormalizedSnapshot.objects.get_or_create(
            organisation=org,
            reporting_period=period,
            defaults={"data_payload": demo_payload}
        )
        snapshot.data_payload = demo_payload
        snapshot.save()

        # 6. Generate Sample Presentation Deck
        context = normalize_report_payload(demo_payload)
        builder = PresentationBuilder(brand=DEFAULT_BRAND, narrative_style="standard")
        prs, manifest, _ = builder.build_presentation(context)

        deck, _ = PresentationDeck.objects.get_or_create(
            organisation=org,
            reporting_period=period,
            normalized_snapshot=snapshot,
            defaults={
                "title": "Board of Directors Presentation - Q3 2026",
                "narrative_style": "standard",
                "commentary_provider": "template",
                "selected_slides": list(SLIDE_REGISTRY.keys()),
                "status": "approved",
                "version": 1,
                "generation_manifest": manifest,
                "validation_warnings": manifest.get("warnings", [])
            }
        )

        pptx_bytes = builder.save_to_bytes(prs)
        filename = f"Apex_Technologies_Holdings_Ltd_Q3_2026_Board_Deck_v1.pptx"
        deck.pptx_file.save(filename, ContentFile(pptx_bytes.getvalue()), save=True)

        # Seed Slide Reviews
        comm_provider = builder.commentary_provider
        commentary = comm_provider.generate_commentary(calculate_financial_metrics(context), style="standard")
        for slide_key in SLIDE_REGISTRY.keys():
            sec = commentary.get_section(slide_key)
            if sec:
                SlideReviewNarrative.objects.update_or_create(
                    presentation_deck=deck,
                    slide_key=slide_key,
                    defaults={
                        "title": sec.title,
                        "custom_lead_summary": sec.lead_summary,
                        "custom_bullets": sec.bullet_points,
                        "is_customized": False
                    }
                )

        # 7. Generate Canonical Sample Workbook
        sample_dir = os.path.join(settings.BASE_DIR, "sample_data")
        os.makedirs(sample_dir, exist_ok=True)
        sample_xlsx_path = os.path.join(sample_dir, "sample_board_deck_data.xlsx")
        self._generate_canonical_excel(sample_xlsx_path, demo_payload)

        AuditLog.objects.create(
            presentation_deck=deck,
            action="SEED_DEMO_DATA",
            details={"organisation": org.name, "period": period.label}
        )

        self.stdout.write(self.style.SUCCESS(f"Successfully seeded demo dataset! Presentation Deck ID: {deck.id}"))
        self.stdout.write(self.style.SUCCESS(f"Generated sample Excel workbook at: {sample_xlsx_path}"))

    def _generate_canonical_excel(self, filepath: str, payload: dict):
        wb = openpyxl.Workbook()
        # Sheet 1: Summary / Metadata
        ws_meta = wb.active
        ws_meta.title = "Summary"
        ws_meta.append(["Property", "Value"])
        ws_meta.append(["Company Name", payload["organisation"]["name"]])
        ws_meta.append(["Currency", payload["organisation"]["currency_code"]])
        ws_meta.append(["Reporting Period", payload["period"]["label"]])
        ws_meta.append(["Prior Period", payload["period"]["prior_period_label"]])
        ws_meta.append(["Budget Period", payload["period"]["budget_period_label"]])

        # Sheet 2: P&L
        ws_pnl = wb.create_sheet(title="P&L")
        ws_pnl.append(["Line Item", "Actual", "Budget", "Prior Period", "YTD Actual", "Category"])
        for item in payload["income_statement"]["revenue_items"]:
            ws_pnl.append([item["label"], float(item["actual"]), float(item["budget"]), float(item["prior"]), float(item["ytd_actual"]), "Revenue"])
        for item in payload["income_statement"]["cogs_items"]:
            ws_pnl.append([item["label"], float(item["actual"]), float(item["budget"]), float(item["prior"]), float(item["ytd_actual"]), "COGS"])
        for item in payload["income_statement"]["opex_items"]:
            ws_pnl.append([item["label"], float(item["actual"]), float(item["budget"]), float(item["prior"]), float(item["ytd_actual"]), "Opex"])

        # Sheet 3: Balance Sheet
        ws_bs = wb.create_sheet(title="Balance Sheet")
        ws_bs.append(["Line Item", "Actual", "Prior Period", "Category"])
        for item in payload["balance_sheet"]["current_assets"]:
            ws_bs.append([item["label"], float(item["actual"]), float(item["prior"]), "Current Assets"])
        for item in payload["balance_sheet"]["non_current_assets"]:
            ws_bs.append([item["label"], float(item["actual"]), float(item["prior"]), "Non-Current Assets"])
        for item in payload["balance_sheet"]["current_liabilities"]:
            ws_bs.append([item["label"], float(item["actual"]), float(item["prior"]), "Current Liabilities"])
        for item in payload["balance_sheet"]["non_current_liabilities"]:
            ws_bs.append([item["label"], float(item["actual"]), float(item["prior"]), "Non-Current Liabilities"])
        for item in payload["balance_sheet"]["equity_items"]:
            ws_bs.append([item["label"], float(item["actual"]), float(item["prior"]), "Equity"])

        # Sheet 4: Cash Flow
        ws_cf = wb.create_sheet(title="Cash Flow")
        ws_cf.append(["Line Item", "Actual", "Budget", "Prior Period"])
        for item in payload["cash_flow"]["operating_activities"]:
            ws_cf.append([item["label"], float(item["actual"]), float(item["budget"]), float(item["prior"])])
        for item in payload["cash_flow"]["investing_activities"]:
            ws_cf.append([item["label"], float(item["actual"]), float(item["budget"]), float(item["prior"])])
        for item in payload["cash_flow"]["financing_activities"]:
            ws_cf.append([item["label"], float(item["actual"]), float(item["budget"]), float(item["prior"])])

        # Sheet 5: KPIs
        ws_kpis = wb.create_sheet(title="KPIs")
        ws_kpis.append(["Metric Name", "Category", "Actual", "Target", "Prior", "Unit", "Definition", "Higher Is Better"])
        for k in payload["kpis"]:
            ws_kpis.append([k["name"], k["category"], float(k["current_value"]), float(k["target_value"]), float(k["prior_value"]), k["unit"], k["definition"], k["higher_is_better"]])

        # Sheet 6: Risks & Actions
        ws_ra = wb.create_sheet(title="Risks & Actions")
        ws_ra.append(["Risks"])
        ws_ra.append(["Title", "Category", "Severity", "Likelihood", "Description", "Mitigation Plan", "Owner"])
        for r in payload["risks"]:
            ws_ra.append([r["title"], r["category"], r["severity"], r["likelihood"], r["description"], r["mitigation_plan"], r["owner"]])
        ws_ra.append([])
        ws_ra.append(["Actions"])
        ws_ra.append(["Action", "Rationale", "Priority", "Owner", "Due Date", "Status"])
        for a in payload["actions"]:
            ws_ra.append([a["title"], a["reason"], a["priority"], a["owner"], a["due_date"], a["status"]])

        # Sheet 7: Multi-period Trend
        ws_trend = wb.create_sheet(title="Trend")
        ws_trend.append(["Period", "Revenue", "Opex", "Operating Profit"])
        for t in payload["income_statement"]["monthly_trend"]:
            ws_trend.append([t["period"], float(t["revenue"]), float(t["opex"]), float(t["profit"])])

        wb.save(filepath)
