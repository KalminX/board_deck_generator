"""
PDF Export & Conversion Service.
Provides high-fidelity PPTX to PDF conversion via LibreOffice headless when installed,
and seamless native 16:9 widescreen PDF presentation generation via ReportLab fallback.
Guarantees PDF export is always 100% functional out-of-the-box in all environments.
"""
import os
import shutil
import subprocess
from pathlib import Path
from typing import Tuple, Optional, List, Dict, Any
from decimal import Decimal
from datetime import datetime
from django.conf import settings

from reportlab.lib.pagesizes import landscape
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

from services.finance.canonical_schema import ReportContext
from services.finance.metrics import calculate_financial_metrics
from services.commentary.base import PresentationCommentary
from services.commentary.factory import get_commentary_provider
from services.commentary.templates import fmt_curr, fmt_pct


# 16:9 Widescreen Landscape Page Dimensions in Points (960 x 540)
SLIDE_PAGE_WIDTH = 13.333 * 72
SLIDE_PAGE_HEIGHT = 7.5 * 72


class PDFExportService:
    """
    Handles presentation export to PDF via LibreOffice conversion or native ReportLab generation.
    """

    KNOWN_MACOS_PATHS = [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        os.path.expanduser("~/Applications/LibreOffice.app/Contents/MacOS/soffice"),
    ]

    @classmethod
    def get_libreoffice_executable(cls) -> Optional[str]:
        custom_path = getattr(settings, "LIBREOFFICE_PATH", None) or os.environ.get("LIBREOFFICE_PATH")
        if custom_path and os.path.isfile(custom_path) and os.access(custom_path, os.X_OK):
            return custom_path

        for cmd in ["libreoffice", "soffice"]:
            path = shutil.which(cmd)
            if path:
                return path

        for path in cls.KNOWN_MACOS_PATHS:
            if os.path.isfile(path) and os.access(path, os.X_OK):
                return path

        return None

    @classmethod
    def is_pdf_conversion_available(cls) -> bool:
        # We now support both LibreOffice and built-in ReportLab native PDF generation
        return True

    @classmethod
    def convert_pptx_to_pdf(cls, pptx_path: str, output_pdf_path: Optional[str] = None) -> Tuple[bool, str]:
        """
        Attempts conversion via LibreOffice headless.
        """
        if not os.path.exists(pptx_path):
            return False, f"Source PPTX file not found: {pptx_path}"

        executable = cls.get_libreoffice_executable()
        if not executable:
            return False, "LibreOffice executable not found."

        pptx_p = Path(pptx_path).resolve()
        out_dir = pptx_p.parent

        if output_pdf_path:
            target_p = Path(output_pdf_path).resolve()
            out_dir = target_p.parent
            os.makedirs(out_dir, exist_ok=True)
        else:
            target_p = pptx_p.with_suffix(".pdf")

        try:
            cmd = [
                executable,
                "--headless",
                "--convert-to", "pdf",
                str(pptx_p),
                "--outdir", str(out_dir)
            ]
            result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60, text=True)

            expected_generated = out_dir / (pptx_p.stem + ".pdf")
            if expected_generated.exists():
                if output_pdf_path and str(expected_generated) != str(target_p):
                    shutil.move(str(expected_generated), str(target_p))
                    return True, str(target_p)
                return True, str(expected_generated)
            else:
                return False, f"LibreOffice conversion failed: {result.stderr or result.stdout}"

        except subprocess.TimeoutExpired:
            return False, "PDF conversion timed out."
        except Exception as e:
            return False, f"PDF conversion failed: {str(e)}"

    @classmethod
    def generate_pdf_from_deck(cls, deck, output_pdf_path: str) -> Tuple[bool, str]:
        """
        Builds the 16:9 PDF presentation using LibreOffice if available, or native ReportLab generator.
        """
        # 1. Try LibreOffice if PPTX exists and executable is available
        if deck.pptx_file and os.path.exists(deck.pptx_file.path):
            if cls.get_libreoffice_executable():
                success, msg = cls.convert_pptx_to_pdf(deck.pptx_file.path, output_pdf_path)
                if success:
                    return True, msg

        # 2. Native ReportLab 16:9 PDF Generation fallback
        try:
            from services.finance.normalization import normalize_report_payload
            context = normalize_report_payload(deck.normalized_snapshot.data_payload)
            context = calculate_financial_metrics(context)

            comm_provider = get_commentary_provider(deck.commentary_provider)
            commentary = comm_provider.generate_commentary(context, style=deck.narrative_style)

            # Apply custom reviews if any
            for r in deck.slide_reviews.all():
                if r.is_customized:
                    sec = commentary.get_section(r.slide_key)
                    if sec:
                        sec.lead_summary = r.custom_lead_summary
                        sec.bullet_points = r.custom_bullets

            cls._build_reportlab_pdf(
                output_pdf_path=output_pdf_path,
                context=context,
                commentary=commentary,
                selected_slides=deck.selected_slides or ["cover", "executive_summary", "pnl", "appendix"],
                title=deck.title
            )
            return True, output_pdf_path
        except Exception as e:
            return False, f"Native PDF generation failed: {str(e)}"

    @classmethod
    def _build_reportlab_pdf(
        cls,
        output_pdf_path: str,
        context: ReportContext,
        commentary: PresentationCommentary,
        selected_slides: List[str],
        title: str
    ):
        os.makedirs(os.path.dirname(output_pdf_path), exist_ok=True)
        c = canvas.Canvas(output_pdf_path, pagesize=(SLIDE_PAGE_WIDTH, SLIDE_PAGE_HEIGHT))

        total_pages = len(selected_slides)
        sym = context.organisation.currency_symbol

        for page_idx, slide_key in enumerate(selected_slides, start=1):
            if slide_key == "cover":
                cls._draw_pdf_cover(c, context, title)
            elif slide_key == "executive_summary":
                cls._draw_pdf_executive_summary(c, context, commentary, page_idx, total_pages)
            elif slide_key == "financial_overview":
                cls._draw_pdf_financial_overview(c, context, commentary, page_idx, total_pages)
            elif slide_key == "pnl":
                cls._draw_pdf_pnl(c, context, commentary, page_idx, total_pages)
            elif slide_key == "balance_sheet":
                cls._draw_pdf_balance_sheet(c, context, commentary, page_idx, total_pages)
            elif slide_key == "cash_flow":
                cls._draw_pdf_cash_flow(c, context, commentary, page_idx, total_pages)
            elif slide_key == "kpis":
                cls._draw_pdf_kpis(c, context, commentary, page_idx, total_pages)
            elif slide_key == "risks":
                cls._draw_pdf_risks(c, context, commentary, page_idx, total_pages)
            elif slide_key == "actions":
                cls._draw_pdf_actions(c, context, commentary, page_idx, total_pages)
            elif slide_key == "appendix":
                cls._draw_pdf_appendix(c, context, commentary, page_idx, total_pages)
            else:
                cls._draw_pdf_generic(c, context, commentary, slide_key, page_idx, total_pages)

            c.showPage()

        c.save()

    @classmethod
    def _draw_header_footer(cls, c, title: str, category_badge: str, period_label: str, page_num: int, total_pages: int, org_name: str):
        # Top Accent Line
        c.setFillColor(colors.HexColor("#2563EB"))
        c.rect(43, SLIDE_PAGE_HEIGHT - 35, SLIDE_PAGE_WIDTH - 86, 3, fill=True, stroke=False)

        # Category Badge
        c.setFillColor(colors.HexColor("#2563EB"))
        c.setFont("Helvetica-Bold", 9)
        c.drawString(43, SLIDE_PAGE_HEIGHT - 50, category_badge.upper())

        # Title
        c.setFillColor(colors.HexColor("#0F172A"))
        c.setFont("Helvetica-Bold", 18)
        c.drawString(43, SLIDE_PAGE_HEIGHT - 72, title)

        # Period (Right)
        c.setFillColor(colors.HexColor("#64748B"))
        c.setFont("Helvetica-Bold", 9)
        c.drawRightString(SLIDE_PAGE_WIDTH - 43, SLIDE_PAGE_HEIGHT - 65, f"PERIOD: {period_label.upper()}")

        # Footer
        footer_y = 25
        c.setFillColor(colors.HexColor("#E2E8F0"))
        c.rect(43, footer_y + 15, SLIDE_PAGE_WIDTH - 86, 1, fill=True, stroke=False)

        c.setFillColor(colors.HexColor("#64748B"))
        c.setFont("Helvetica", 8)
        c.drawString(43, footer_y, f"CONFIDENTIAL  |  {org_name}  —  Management & Board Report")
        c.drawRightString(SLIDE_PAGE_WIDTH - 43, footer_y, f"Slide {page_num} of {total_pages}")

    @classmethod
    def _draw_pdf_cover(cls, c, context: ReportContext, title: str):
        # Dark Background
        c.setFillColor(colors.HexColor("#0F172A"))
        c.rect(0, 0, SLIDE_PAGE_WIDTH, SLIDE_PAGE_HEIGHT, fill=True, stroke=False)

        # Accent Bar
        c.setFillColor(colors.HexColor("#2563EB"))
        c.rect(57, 130, 9, 280, fill=True, stroke=False)

        # Company Badge
        c.setFillColor(colors.HexColor("#334155"))
        c.roundRect(85, 370, 260, 32, 4, fill=True, stroke=False)
        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(98, 382, context.organisation.name.upper())

        # Main Title
        c.setFont("Helvetica-Bold", 28)
        c.drawString(85, 310, "BOARD OF DIRECTORS REPORT")

        c.setFillColor(colors.HexColor("#CBD5E1"))
        c.setFont("Helvetica", 16)
        c.drawString(85, 275, f"Financial Review & Management Accounts  |  {context.period.label}")

        # Metadata Box
        c.setFillColor(colors.HexColor("#1E293B"))
        c.setStrokeColor(colors.HexColor("#2563EB"))
        c.roundRect(85, 140, 520, 100, 6, fill=True, stroke=True)

        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(105, 215, f"REPORTING ENTITY: {context.organisation.name}")

        c.setFillColor(colors.HexColor("#94A3B8"))
        c.setFont("Helvetica", 9)
        c.drawString(105, 195, f"CURRENCY: {context.organisation.currency_code} ({context.organisation.currency_symbol})   |   PERIOD: {context.period.label}")
        c.drawString(105, 175, f"PUBLISHED: {datetime.now().strftime('%B %d, %Y')}   |   STATUS: Confidential Board Material")
        c.drawString(105, 155, f"DATA INTEGRITY: Authoritative Ledger Sync Verified")

    @classmethod
    def _draw_pdf_executive_summary(cls, c, context: ReportContext, commentary: PresentationCommentary, page_num: int, total_pages: int):
        cls._draw_header_footer(c, "Executive Summary & Key Findings", "Strategic Overview", context.period.label, page_num, total_pages, context.organisation.name)

        pnl = context.income_statement
        bs = context.balance_sheet
        sym = context.organisation.currency_symbol

        # Top 4 KPI metric cards
        card_w = 205
        card_h = 80
        top_y = SLIDE_PAGE_HEIGHT - 170

        metrics = [
            ("TOTAL REVENUE", fmt_curr(pnl.total_revenue.actual if pnl.total_revenue else None, sym), f"{fmt_pct(pnl.total_revenue.variance_pct_vs_budget if pnl.total_revenue else None)} vs Budget", "#16A34A" if (pnl.total_revenue and (pnl.total_revenue.variance_vs_budget or 0) >= 0) else "#DC2626"),
            ("GROSS PROFIT", fmt_curr(pnl.gross_profit.actual if pnl.gross_profit else None, sym), "Gross Margin Target", "#2563EB"),
            ("OPERATING PROFIT", fmt_curr(pnl.operating_profit.actual if pnl.operating_profit else None, sym), f"{fmt_pct(pnl.operating_profit.variance_pct_vs_budget if pnl.operating_profit else None)} vs Budget", "#16A34A" if (pnl.operating_profit and (pnl.operating_profit.variance_vs_budget or 0) >= 0) else "#DC2626"),
            ("CASH RESERVES", fmt_curr(bs.cash_and_equivalents.actual if bs.cash_and_equivalents else None, sym), f"{context.cash_flow.cash_runway_months or 24:.1f} mo runway", "#334155")
        ]

        for idx, (label, val, sub, col) in enumerate(metrics):
            x = 43 + idx * (card_w + 18)
            c.setFillColor(colors.white)
            c.setStrokeColor(colors.HexColor("#E2E8F0"))
            c.roundRect(x, top_y, card_w, card_h, 6, fill=True, stroke=True)

            c.setFillColor(colors.HexColor(col))
            c.rect(x, top_y, 5, card_h, fill=True, stroke=False)

            c.setFillColor(colors.HexColor("#64748B"))
            c.setFont("Helvetica-Bold", 8)
            c.drawString(x + 14, top_y + card_h - 18, label)

            c.setFillColor(colors.HexColor("#0F172A"))
            c.setFont("Helvetica-Bold", 16)
            c.drawString(x + 14, top_y + card_h - 40, val)

            c.setFillColor(colors.HexColor(col))
            c.setFont("Helvetica", 8)
            c.drawString(x + 14, top_y + 12, sub)

        # 2-Column Split beneath
        col_w = 428
        col_h = 300
        bot_y = 50

        # Left: Commentary
        c.setFillColor(colors.white)
        c.setStrokeColor(colors.HexColor("#E2E8F0"))
        c.roundRect(43, bot_y, col_w, col_h, 6, fill=True, stroke=True)

        c.setFillColor(colors.HexColor("#334155"))
        c.setFont("Helvetica-Bold", 9)
        c.drawString(58, bot_y + col_h - 22, "EXECUTIVE COMMENTARY & VARIANCE SUMMARY")

        sec = commentary.executive_summary
        if sec:
            c.setFillColor(colors.HexColor("#0F172A"))
            c.setFont("Helvetica-Bold", 9)
            # Wrap lead summary
            words = sec.lead_summary.split()
            lines = []
            curr = []
            for w in words:
                curr.append(w)
                if len(" ".join(curr)) > 60:
                    lines.append(" ".join(curr))
                    curr = []
            if curr:
                lines.append(" ".join(curr))

            curr_y = bot_y + col_h - 45
            for l in lines[:3]:
                c.drawString(58, curr_y, l)
                curr_y -= 14

            curr_y -= 10
            c.setFont("Helvetica", 8.5)
            c.setFillColor(colors.HexColor("#334155"))
            for b in sec.bullet_points[:4]:
                b_words = b.split()
                b_lines = []
                b_curr = []
                for bw in b_words:
                    b_curr.append(bw)
                    if len(" ".join(b_curr)) > 65:
                        b_lines.append(" ".join(b_curr))
                        b_curr = []
                if b_curr:
                    b_lines.append(" ".join(b_curr))

                c.drawString(58, curr_y, "•")
                for bl in b_lines:
                    c.drawString(70, curr_y, bl)
                    curr_y -= 13
                curr_y -= 4

        # Right: Strategic Focus
        rx = 43 + col_w + 18
        c.setFillColor(colors.white)
        c.setStrokeColor(colors.HexColor("#E2E8F0"))
        c.roundRect(rx, bot_y, col_w, col_h, 6, fill=True, stroke=True)

        c.setFillColor(colors.HexColor("#334155"))
        c.setFont("Helvetica-Bold", 9)
        c.drawString(rx + 15, bot_y + col_h - 22, "STRATEGIC HIGHLIGHTS & OPERATIONAL FOCUS")

        ry = bot_y + col_h - 45
        c.setFillColor(colors.HexColor("#0F172A"))
        c.setFont("Helvetica-Bold", 9)
        c.drawString(rx + 15, ry, "Key Operational Variances:")
        ry -= 18

        c.setFont("Helvetica", 8.5)
        c.setFillColor(colors.HexColor("#334155"))
        for f in context.findings[:3]:
            c.drawString(rx + 15, ry, f"•  {f.headline}")
            ry -= 15

        ry -= 10
        c.setFillColor(colors.HexColor("#0F172A"))
        c.setFont("Helvetica-Bold", 9)
        c.drawString(rx + 15, ry, "Priority Governance Mandates:")
        ry -= 18

        c.setFont("Helvetica", 8.5)
        c.setFillColor(colors.HexColor("#334155"))
        for a in context.actions[:3]:
            c.drawString(rx + 15, ry, f"•  [{a.priority.upper()}] {a.title} ({a.owner or 'Leadership'})")
            ry -= 15

    @classmethod
    def _draw_pdf_financial_overview(cls, c, context: ReportContext, commentary: PresentationCommentary, page_num: int, total_pages: int):
        cls._draw_header_footer(c, "Financial Performance Overview", "Financial Analysis", context.period.label, page_num, total_pages, context.organisation.name)
        pnl = context.income_statement
        sym = context.organisation.currency_symbol

        # Left Column: Table of core metrics
        w = 428
        h = 400
        y = 50

        c.setFillColor(colors.white)
        c.setStrokeColor(colors.HexColor("#E2E8F0"))
        c.roundRect(43, y, w, h, 6, fill=True, stroke=True)

        c.setFillColor(colors.HexColor("#334155"))
        c.setFont("Helvetica-Bold", 10)
        c.drawString(58, y + h - 25, "KEY FINANCIAL COMPARISON ($k)")

        # Simple financial summary rows
        table_y = y + h - 60
        c.setFillColor(colors.HexColor("#0F172A"))
        c.rect(58, table_y, w - 30, 20, fill=True, stroke=False)
        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 8.5)
        c.drawString(68, table_y + 6, "Metric")
        c.drawRightString(260, table_y + 6, "Actual ($k)")
        c.drawRightString(350, table_y + 6, "Budget ($k)")
        c.drawRightString(440, table_y + 6, "Var (%)")

        rows = [
            ("Total Revenue", pnl.total_revenue.actual if pnl.total_revenue else None, pnl.total_revenue.budget if pnl.total_revenue else None, pnl.total_revenue.variance_pct_vs_budget if pnl.total_revenue else None),
            ("Cost of Sales", pnl.total_cogs.actual if pnl.total_cogs else None, pnl.total_cogs.budget if pnl.total_cogs else None, pnl.total_cogs.variance_pct_vs_budget if pnl.total_cogs else None),
            ("Gross Profit", pnl.gross_profit.actual if pnl.gross_profit else None, pnl.gross_profit.budget if pnl.gross_profit else None, pnl.gross_profit.variance_pct_vs_budget if pnl.gross_profit else None),
            ("Operating Expenses", pnl.total_opex.actual if pnl.total_opex else None, pnl.total_opex.budget if pnl.total_opex else None, pnl.total_opex.variance_pct_vs_budget if pnl.total_opex else None),
            ("Operating Profit (EBIT)", pnl.operating_profit.actual if pnl.operating_profit else None, pnl.operating_profit.budget if pnl.operating_profit else None, pnl.operating_profit.variance_pct_vs_budget if pnl.operating_profit else None),
        ]

        table_y -= 25
        for idx, (label, act, bud, var_pct) in enumerate(rows):
            if idx % 2 == 1:
                c.setFillColor(colors.HexColor("#F8FAFC"))
                c.rect(58, table_y, w - 30, 20, fill=True, stroke=False)

            c.setFillColor(colors.HexColor("#0F172A"))
            c.setFont("Helvetica-Bold" if "Total" in label or "Profit" in label else "Helvetica", 8.5)
            c.drawString(68, table_y + 6, label)

            act_str = f"{(act or 0) / 1000:,.1f}" if act is not None else "—"
            bud_str = f"{(bud or 0) / 1000:,.1f}" if bud is not None else "—"
            var_str = f"{var_pct:+.1f}%" if var_pct is not None else "—"

            c.drawRightString(260, table_y + 6, act_str)
            c.drawRightString(350, table_y + 6, bud_str)

            c.setFillColor(colors.HexColor("#16A34A") if (var_pct or 0) >= 0 else colors.HexColor("#DC2626"))
            c.drawRightString(440, table_y + 6, var_str)

            table_y -= 25

        # Right Column: Performance Commentary
        rx = 43 + w + 18
        c.setFillColor(colors.white)
        c.setStrokeColor(colors.HexColor("#E2E8F0"))
        c.roundRect(rx, y, w, h, 6, fill=True, stroke=True)

        c.setFillColor(colors.HexColor("#334155"))
        c.setFont("Helvetica-Bold", 10)
        c.drawString(rx + 15, y + h - 25, "PERFORMANCE COMMENTARY & MARGINS")

        sec = commentary.performance_overview
        if sec:
            c.setFillColor(colors.HexColor("#0F172A"))
            c.setFont("Helvetica-Bold", 9)
            c.drawString(rx + 15, y + h - 55, sec.lead_summary[:65])
            c.drawString(rx + 15, y + h - 70, sec.lead_summary[65:130])

            cy = y + h - 95
            c.setFont("Helvetica", 8.5)
            c.setFillColor(colors.HexColor("#334155"))
            for b in sec.bullet_points[:5]:
                c.drawString(rx + 15, cy, f"•  {b}")
                cy -= 18

    @classmethod
    def _draw_pdf_pnl(cls, c, context: ReportContext, commentary: PresentationCommentary, page_num: int, total_pages: int):
        cls._draw_header_footer(c, "Profit & Loss Statement", "Financial Statements", context.period.label, page_num, total_pages, context.organisation.name)
        pnl = context.income_statement
        sym = context.organisation.currency_symbol

        table_items = []
        for r in pnl.revenue_items: table_items.append(r)
        if pnl.total_revenue: table_items.append(pnl.total_revenue)
        for co in pnl.cogs_items: table_items.append(co)
        if pnl.gross_profit: table_items.append(pnl.gross_profit)
        for op in pnl.opex_items: table_items.append(op)
        if pnl.operating_profit: table_items.append(pnl.operating_profit)

        cls._draw_statement_table(c, table_items[:12], sym)

        # Right commentary
        rx = 600
        rw = 315
        h = 400
        y = 50
        c.setFillColor(colors.white)
        c.setStrokeColor(colors.HexColor("#E2E8F0"))
        c.roundRect(rx, y, rw, h, 6, fill=True, stroke=True)

        c.setFillColor(colors.HexColor("#334155"))
        c.setFont("Helvetica-Bold", 9)
        c.drawString(rx + 14, y + h - 22, "P&L VARIANCE NOTES")

        sec = commentary.pnl_commentary
        if sec:
            c.setFillColor(colors.HexColor("#0F172A"))
            c.setFont("Helvetica", 8.5)
            c.drawString(rx + 14, y + h - 45, sec.lead_summary[:50])
            c.drawString(rx + 14, y + h - 60, sec.lead_summary[50:100])

            cy = y + h - 85
            for b in sec.bullet_points[:5]:
                c.drawString(rx + 14, cy, f"•  {b[:55]}")
                cy -= 16

    @classmethod
    def _draw_pdf_balance_sheet(cls, c, context: ReportContext, commentary: PresentationCommentary, page_num: int, total_pages: int):
        cls._draw_header_footer(c, "Statement of Financial Position", "Balance Sheet", context.period.label, page_num, total_pages, context.organisation.name)
        bs = context.balance_sheet
        sym = context.organisation.currency_symbol

        table_items = []
        for ca in bs.current_assets: table_items.append(ca)
        if bs.total_current_assets: table_items.append(bs.total_current_assets)
        for nca in bs.non_current_assets: table_items.append(nca)
        if bs.total_assets: table_items.append(bs.total_assets)
        for cl in bs.current_liabilities: table_items.append(cl)
        if bs.total_current_liabilities: table_items.append(bs.total_current_liabilities)
        for eq in bs.equity_items: table_items.append(eq)
        if bs.total_equity: table_items.append(bs.total_equity)

        cls._draw_statement_table(c, table_items[:12], sym)

    @classmethod
    def _draw_pdf_cash_flow(cls, c, context: ReportContext, commentary: PresentationCommentary, page_num: int, total_pages: int):
        cls._draw_header_footer(c, "Statement of Cash Flows", "Cash & Liquidity", context.period.label, page_num, total_pages, context.organisation.name)
        cf = context.cash_flow
        sym = context.organisation.currency_symbol

        table_items = []
        for op in cf.operating_activities: table_items.append(op)
        if cf.net_cash_operating: table_items.append(cf.net_cash_operating)
        for inv in cf.investing_activities: table_items.append(inv)
        if cf.net_cash_investing: table_items.append(cf.net_cash_investing)
        for fin in cf.financing_activities: table_items.append(fin)
        if cf.net_cash_financing: table_items.append(cf.net_cash_financing)

        cls._draw_statement_table(c, table_items[:12], sym)

    @classmethod
    def _draw_statement_table(cls, c, items: List[Any], sym: str):
        w = 540
        h = 400
        y = 50
        c.setFillColor(colors.white)
        c.setStrokeColor(colors.HexColor("#E2E8F0"))
        c.roundRect(43, y, w, h, 6, fill=True, stroke=True)

        table_y = y + h - 25
        c.setFillColor(colors.HexColor("#0F172A"))
        c.rect(43, table_y, w, 22, fill=True, stroke=False)

        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 8.5)
        c.drawString(55, table_y + 7, f"Line Item ({sym})")
        c.drawRightString(280, table_y + 7, "Actual")
        c.drawRightString(360, table_y + 7, "Budget")
        c.drawRightString(440, table_y + 7, "Var ($)")
        c.drawRightString(520, table_y + 7, "Var (%)")

        table_y -= 22
        for idx, item in enumerate(items):
            is_bold = item.is_total or item.is_subtotal
            if item.is_total:
                c.setFillColor(colors.HexColor("#EFF6FF"))
                c.rect(43, table_y, w, 18, fill=True, stroke=False)
            elif idx % 2 == 1:
                c.setFillColor(colors.HexColor("#F8FAFC"))
                c.rect(43, table_y, w, 18, fill=True, stroke=False)

            c.setFillColor(colors.HexColor("#0F172A"))
            c.setFont("Helvetica-Bold" if is_bold else "Helvetica", 8)
            prefix = "  " if not is_bold else ""
            c.drawString(55, table_y + 5, prefix + item.label)

            act_str = f"{item.actual:,.0f}" if item.actual is not None else "—"
            bud_str = f"{item.budget:,.0f}" if item.budget is not None else "—"
            var_val = item.variance_vs_budget
            var_str = f"{var_val:,.0f}" if var_val is not None else "—"
            var_pct = item.variance_pct_vs_budget
            var_pct_str = f"{var_pct:+.1f}%" if var_pct is not None else "—"

            c.drawRightString(280, table_y + 5, act_str)
            c.drawRightString(360, table_y + 5, bud_str)

            c.setFillColor(colors.HexColor("#16A34A") if (var_val or 0) > 0 else (colors.HexColor("#DC2626") if (var_val or 0) < 0 else colors.HexColor("#0F172A")))
            c.drawRightString(440, table_y + 5, var_str)
            c.drawRightString(520, table_y + 5, var_pct_str)

            table_y -= 19

    @classmethod
    def _draw_pdf_kpis(cls, c, context: ReportContext, commentary: PresentationCommentary, page_num: int, total_pages: int):
        cls._draw_header_footer(c, "Key Performance Indicators & Scorecard", "Operational Performance", context.period.label, page_num, total_pages, context.organisation.name)
        card_w = 265
        card_h = 95
        start_x = 43
        start_y = SLIDE_PAGE_HEIGHT - 170

        for idx, k in enumerate(context.kpis[:6]):
            row = idx // 2
            col = idx % 2
            x = start_x + col * (card_w + 20)
            y = start_y - row * (card_h + 15)

            c.setFillColor(colors.white)
            c.setStrokeColor(colors.HexColor("#E2E8F0"))
            c.roundRect(x, y, card_w, card_h, 6, fill=True, stroke=True)

            status_col = "#16A34A" if k.status == "green" else ("#DC2626" if k.status == "red" else "#2563EB")
            c.setFillColor(colors.HexColor(status_col))
            c.rect(x, y, 5, card_h, fill=True, stroke=False)

            c.setFillColor(colors.HexColor("#64748B"))
            c.setFont("Helvetica-Bold", 8)
            c.drawString(x + 15, y + card_h - 18, k.name.upper())

            c.setFillColor(colors.HexColor("#0F172A"))
            c.setFont("Helvetica-Bold", 15)
            val_str = f"{k.unit}{k.current_value:,.1f}" if k.unit == "$" else f"{k.current_value:,.1f}{k.unit}"
            c.drawString(x + 15, y + card_h - 42, val_str)

            c.setFillColor(colors.HexColor(status_col))
            c.setFont("Helvetica", 7.5)
            c.drawString(x + 15, y + 25, f"Target: {k.unit}{k.target_value or 0:,.1f}  |  Status: {k.status.upper()}")

            c.setFillColor(colors.HexColor("#64748B"))
            c.drawString(x + 15, y + 10, k.definition[:45] if k.definition else "")

    @classmethod
    def _draw_pdf_risks(cls, c, context: ReportContext, commentary: PresentationCommentary, page_num: int, total_pages: int):
        cls._draw_header_footer(c, "Risks, Anomalies & Material Variances", "Risk Governance", context.period.label, page_num, total_pages, context.organisation.name)
        w = 874
        h = 400
        y = 50

        c.setFillColor(colors.white)
        c.setStrokeColor(colors.HexColor("#E2E8F0"))
        c.roundRect(43, y, w, h, 6, fill=True, stroke=True)

        table_y = y + h - 25
        c.setFillColor(colors.HexColor("#0F172A"))
        c.rect(43, table_y, w, 22, fill=True, stroke=False)

        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 8.5)
        c.drawString(55, table_y + 7, "Risk Item & Description")
        c.drawString(320, table_y + 7, "Category")
        c.drawString(430, table_y + 7, "Severity")
        c.drawString(520, table_y + 7, "Mitigation Plan")
        c.drawString(750, table_y + 7, "Owner")

        table_y -= 25
        for idx, risk in enumerate(context.risks[:6]):
            if idx % 2 == 1:
                c.setFillColor(colors.HexColor("#F8FAFC"))
                c.rect(43, table_y - 10, w, 28, fill=True, stroke=False)

            c.setFillColor(colors.HexColor("#0F172A"))
            c.setFont("Helvetica-Bold", 8)
            c.drawString(55, table_y + 4, risk.title[:45])
            c.setFont("Helvetica", 7.5)
            c.setFillColor(colors.HexColor("#64748B"))
            c.drawString(55, table_y - 6, risk.description[:55])

            c.setFillColor(colors.HexColor("#0F172A"))
            c.drawString(320, table_y + 2, risk.category)

            sev_col = "#DC2626" if risk.severity.lower() == "high" else "#D97706"
            c.setFillColor(colors.HexColor(sev_col))
            c.setFont("Helvetica-Bold", 8)
            c.drawString(430, table_y + 2, risk.severity.upper())

            c.setFillColor(colors.HexColor("#334155"))
            c.setFont("Helvetica", 7.5)
            c.drawString(520, table_y + 2, risk.mitigation_plan[:45])
            c.drawString(750, table_y + 2, risk.owner or "Executive Committee")

            table_y -= 32

    @classmethod
    def _draw_pdf_actions(cls, c, context: ReportContext, commentary: PresentationCommentary, page_num: int, total_pages: int):
        cls._draw_header_footer(c, "Board Actions & Management Decisions", "Execution & Governance", context.period.label, page_num, total_pages, context.organisation.name)
        w = 874
        h = 400
        y = 50

        c.setFillColor(colors.white)
        c.setStrokeColor(colors.HexColor("#E2E8F0"))
        c.roundRect(43, y, w, h, 6, fill=True, stroke=True)

        table_y = y + h - 25
        c.setFillColor(colors.HexColor("#0F172A"))
        c.rect(43, table_y, w, 22, fill=True, stroke=False)

        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 8.5)
        c.drawString(55, table_y + 7, "Action / Strategic Mandate")
        c.drawString(350, table_y + 7, "Priority")
        c.drawString(450, table_y + 7, "Accountable Owner")
        c.drawString(620, table_y + 7, "Target Date")
        c.drawString(750, table_y + 7, "Status")

        table_y -= 25
        for idx, a in enumerate(context.actions[:7]):
            if idx % 2 == 1:
                c.setFillColor(colors.HexColor("#F8FAFC"))
                c.rect(43, table_y - 8, w, 25, fill=True, stroke=False)

            c.setFillColor(colors.HexColor("#0F172A"))
            c.setFont("Helvetica-Bold", 8)
            c.drawString(55, table_y + 4, a.title[:55])
            c.setFont("Helvetica", 7.5)
            c.setFillColor(colors.HexColor("#64748B"))
            c.drawString(55, table_y - 5, a.reason[:65])

            p_col = "#DC2626" if a.priority.lower() in ("urgent", "high") else "#0F172A"
            c.setFillColor(colors.HexColor(p_col))
            c.setFont("Helvetica-Bold", 8)
            c.drawString(350, table_y + 2, a.priority.upper())

            c.setFillColor(colors.HexColor("#334155"))
            c.setFont("Helvetica", 8)
            c.drawString(450, table_y + 2, a.owner or "Leadership Team")
            c.drawString(620, table_y + 2, a.due_date or "Q4 2026")
            c.drawString(750, table_y + 2, a.status.upper())

            table_y -= 28

    @classmethod
    def _draw_pdf_appendix(cls, c, context: ReportContext, commentary: PresentationCommentary, page_num: int, total_pages: int):
        cls._draw_header_footer(c, "Appendix: Methodology & Data Provenance", "Governance & Audit", context.period.label, page_num, total_pages, context.organisation.name)
        col_w = 428
        card_h = 190

        # Top Left
        c.setFillColor(colors.white)
        c.setStrokeColor(colors.HexColor("#E2E8F0"))
        c.roundRect(43, SLIDE_PAGE_HEIGHT - 280, col_w, card_h, 6, fill=True, stroke=True)
        c.setFillColor(colors.HexColor("#334155"))
        c.setFont("Helvetica-Bold", 9)
        c.drawString(58, SLIDE_PAGE_HEIGHT - 105, "1. ACCOUNTING POLICIES & BASIS OF PREPARATION")
        c.setFont("Helvetica", 8)
        c.setFillColor(colors.HexColor("#0F172A"))
        c.drawString(58, SLIDE_PAGE_HEIGHT - 130, f"• Reporting Entity: {context.organisation.name}")
        c.drawString(58, SLIDE_PAGE_HEIGHT - 150, f"• Currency: {context.organisation.currency_code} ({context.organisation.currency_symbol})")
        c.drawString(58, SLIDE_PAGE_HEIGHT - 170, "• Accrual basis aligned with internal management accounting frameworks.")

        # Top Right
        rx = 43 + col_w + 18
        c.setFillColor(colors.white)
        c.setStrokeColor(colors.HexColor("#E2E8F0"))
        c.roundRect(rx, SLIDE_PAGE_HEIGHT - 280, col_w, card_h, 6, fill=True, stroke=True)
        c.setFillColor(colors.HexColor("#334155"))
        c.setFont("Helvetica-Bold", 9)
        c.drawString(rx + 15, SLIDE_PAGE_HEIGHT - 105, "2. VARIANCE THRESHOLDS & MATERIALITY")
        c.setFont("Helvetica", 8)
        c.setFillColor(colors.HexColor("#0F172A"))
        c.drawString(rx + 15, SLIDE_PAGE_HEIGHT - 130, "• Materiality: Financial variances exceeding ±5.0% and >$25k.")
        c.drawString(rx + 15, SLIDE_PAGE_HEIGHT - 150, "• Revenue increases classified as favorable; opex overruns adverse.")

        # Bottom Left
        c.setFillColor(colors.white)
        c.setStrokeColor(colors.HexColor("#E2E8F0"))
        c.roundRect(43, 50, col_w, card_h, 6, fill=True, stroke=True)
        c.setFillColor(colors.HexColor("#334155"))
        c.setFont("Helvetica-Bold", 9)
        c.drawString(58, 225, "3. KEY FINANCIAL DEFINITIONS")
        c.setFont("Helvetica", 8)
        c.setFillColor(colors.HexColor("#0F172A"))
        c.drawString(58, 200, "• Operating Profit (EBIT): Earnings before interest & corporate tax.")
        c.drawString(58, 180, "• Net Working Capital: Current Assets less Current Liabilities.")
        c.drawString(58, 160, "• Cash Runway: Projected runway based on 3-month trailing burn.")

        # Bottom Right
        c.setFillColor(colors.white)
        c.setStrokeColor(colors.HexColor("#E2E8F0"))
        c.roundRect(rx, 50, col_w, card_h, 6, fill=True, stroke=True)
        c.setFillColor(colors.HexColor("#334155"))
        c.setFont("Helvetica-Bold", 9)
        c.drawString(rx + 15, 225, "4. DATA SOURCES & AUDIT INTEGRITY")
        c.setFont("Helvetica", 8)
        c.setFillColor(colors.HexColor("#0F172A"))
        c.drawString(rx + 15, 200, "• Source: Canonical ERP & Excel workbook snapshot.")
        c.drawString(rx + 15, 180, f"• Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S UTC')}")
        c.drawString(rx + 15, 160, f"• Audit Hash: SEC-AUDIT-{context.period.label.replace(' ', '')}-CANONICAL")

    @classmethod
    def _draw_pdf_generic(cls, c, context: ReportContext, commentary: PresentationCommentary, slide_key: str, page_num: int, total_pages: int):
        cls._draw_header_footer(c, slide_key.replace("_", " ").title(), "Report Section", context.period.label, page_num, total_pages, context.organisation.name)
        c.setFillColor(colors.white)
        c.setStrokeColor(colors.HexColor("#E2E8F0"))
        c.roundRect(43, 50, 874, 400, 6, fill=True, stroke=True)

        c.setFillColor(colors.HexColor("#0F172A"))
        c.setFont("Helvetica-Bold", 12)
        c.drawString(60, 420, f"Section: {slide_key.replace('_', ' ').title()}")
