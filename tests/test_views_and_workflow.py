"""
Integration Tests for Django Views, Workflow, Narrative Editing & Approval Invalidation.
"""
import os
import json
import pytest
from django.test import TestCase, Client
from django.urls import reverse
from django.core.files.uploadedfile import SimpleUploadedFile
from board_deck.models import (
    Organisation, ReportingPeriod, NormalizedSnapshot,
    PresentationDeck, SlideReviewNarrative, AuditLog
)
from services.finance.canonical_schema import ReportContext
from services.finance.normalization import normalize_report_payload
from services.presentations.builder import PresentationBuilder, DEFAULT_BRAND


class PresentationWorkflowTestCase(TestCase):
    def setUp(self):
        self.client = Client()
        self.org = Organisation.objects.create(name="Apex Technologies Holdings Ltd")
        self.period = ReportingPeriod.objects.create(organisation=self.org, label="Q3 2026")

        self.payload = {
            "organisation": {"name": "Apex Technologies Holdings Ltd", "currency_code": "USD", "currency_symbol": "$"},
            "period": {"label": "Q3 2026", "period_type": "quarter"},
            "income_statement": {
                "revenue_items": [
                    {"key": "rev_sub", "label": "Enterprise Subscriptions", "actual": "1000000", "budget": "900000", "prior": "850000", "category": "revenue"}
                ],
                "cogs_items": [
                    {"key": "cogs_host", "label": "Cloud Hosting", "actual": "200000", "budget": "180000", "prior": "170000", "category": "cogs"}
                ],
                "opex_items": [
                    {"key": "opex_sales", "label": "Sales & Marketing", "actual": "400000", "budget": "380000", "prior": "350000", "category": "opex"}
                ]
            },
            "balance_sheet": {
                "current_assets": [{"key": "cash", "label": "Cash", "actual": "3000000", "prior": "2500000", "category": "current_asset"}]
            },
            "cash_flow": {
                "operating_activities": [{"key": "cf_net", "label": "Operating Cash", "actual": "400000"}],
                "cash_runway_months": "24.0"
            },
            "kpis": [],
            "risks": [],
            "actions": []
        }

        self.snapshot = NormalizedSnapshot.objects.create(
            organisation=self.org,
            reporting_period=self.period,
            data_payload=self.payload
        )

    def test_dashboard_view_loads(self):
        response = self.client.get(reverse("board_deck:dashboard"))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Executive Presentation Runs")

    def test_sample_workbook_download(self):
        response = self.client.get(reverse("board_deck:download_sample_workbook"))
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response["Content-Type"], "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    def test_upload_workbook_post_creates_snapshot_and_redirects(self):
        sample_path = "sample_data/sample_board_deck_data.xlsx"
        with open(sample_path, "rb") as f:
            excel_content = f.read()

        uploaded_file = SimpleUploadedFile(
            "sample_board_deck_data.xlsx",
            excel_content,
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        upload_url = reverse("board_deck:upload")
        response = self.client.post(upload_url, {
            "organisation_name": "Apex Technologies Holdings Ltd",
            "period_label": "Q3 2026",
            "file": uploaded_file
        }, follow=True)

        self.assertEqual(response.status_code, 200)
        # Should redirect to create_presentation_from_snapshot
        snapshot = NormalizedSnapshot.objects.filter(organisation__name="Apex Technologies Holdings Ltd").latest("created_at")
        self.assertIsNotNone(snapshot)
        self.assertIn("income_statement", snapshot.data_payload)

    def test_generate_presentation_view(self):
        url = reverse("board_deck:create_presentation_from_snapshot", args=[self.snapshot.id])
        # GET form
        response = self.client.get(url)
        self.assertEqual(response.status_code, 200)

        # POST form
        data = {
            "title": "Q3 2026 Board Review",
            "narrative_style": "standard",
            "brand_theme": "Corporate Slate",
            "selected_slides": ["cover", "executive_summary", "pnl", "appendix"]
        }
        response = self.client.post(url, data, follow=True)
        self.assertEqual(response.status_code, 200)

        deck = PresentationDeck.objects.get(title="Q3 2026 Board Review")
        self.assertEqual(deck.status, "review")
        self.assertEqual(deck.version, 1)
        self.assertTrue(bool(deck.pptx_file))
        self.assertEqual(len(deck.selected_slides), 4)

    def test_edit_narrative_and_approval_invalidation(self):
        # Create deck
        context = normalize_report_payload(self.payload)
        builder = PresentationBuilder()
        prs, manifest, _ = builder.build_presentation(context, selected_slides=["cover", "executive_summary", "pnl"])
        bio = builder.save_to_bytes(prs)

        deck = PresentationDeck.objects.create(
            organisation=self.org,
            reporting_period=self.period,
            normalized_snapshot=self.snapshot,
            title="Approval Test Deck",
            selected_slides=["cover", "executive_summary", "pnl"],
            status="approved",
            version=1,
            generation_manifest=manifest
        )
        deck.pptx_file.save("test_deck.pptx", bio)

        # Approve deck
        deck.status = "approved"
        deck.approved_version = 1
        deck.save()
        self.assertEqual(deck.status, "approved")

        # Edit narrative via POST endpoint
        edit_url = reverse("board_deck:save_narrative_edit", args=[deck.id])
        edit_data = {
            "slide_key": "executive_summary",
            "lead_summary": "Updated executive revenue growth commentary.",
            "bullets_text": "First bullet\nSecond bullet"
        }
        response = self.client.post(edit_url, edit_data, follow=True)
        self.assertEqual(response.status_code, 200)

        # Reload deck from db
        deck.refresh_from_db()
        # Verifies that approval was invalidated!
        self.assertEqual(deck.status, "review")
        self.assertIsNone(deck.approved_version)
        self.assertEqual(deck.version, 2)

        # Verify review record
        rev = deck.slide_reviews.get(slide_key="executive_summary")
        self.assertEqual(rev.custom_lead_summary, "Updated executive revenue growth commentary.")
        self.assertEqual(rev.custom_bullets, ["First bullet", "Second bullet"])
        self.assertTrue(rev.is_customized)

    def test_approve_deck_view(self):
        deck = PresentationDeck.objects.create(
            organisation=self.org,
            reporting_period=self.period,
            normalized_snapshot=self.snapshot,
            title="Signoff Deck",
            selected_slides=["cover", "executive_summary"],
            status="review",
            version=1
        )
        url = reverse("board_deck:approve_deck", args=[deck.id])
        response = self.client.post(url, follow=True)
        self.assertEqual(response.status_code, 200)

        deck.refresh_from_db()
        self.assertEqual(deck.status, "approved")
        self.assertEqual(deck.approved_version, 1)
        self.assertIsNotNone(deck.approved_at)

    def test_download_endpoints(self):
        context = normalize_report_payload(self.payload)
        builder = PresentationBuilder()
        prs, manifest, _ = builder.build_presentation(context, selected_slides=["cover", "pnl"])
        bio = builder.save_to_bytes(prs)

        deck = PresentationDeck.objects.create(
            organisation=self.org,
            reporting_period=self.period,
            normalized_snapshot=self.snapshot,
            title="Download Deck",
            selected_slides=["cover", "pnl"],
            status="approved",
            version=1,
            generation_manifest=manifest
        )
        deck.pptx_file.save("download_test.pptx", bio)

        # Download PPTX
        pptx_resp = self.client.get(reverse("board_deck:download_pptx", args=[deck.id]))
        self.assertEqual(pptx_resp.status_code, 200)
        self.assertIn("presentationml", pptx_resp["Content-Type"])

        # Download Data JSON
        data_resp = self.client.get(reverse("board_deck:download_data_summary", args=[deck.id]))
        self.assertEqual(data_resp.status_code, 200)
        self.assertEqual(data_resp["Content-Type"], "application/json")

        # Download Manifest JSON
        man_resp = self.client.get(reverse("board_deck:download_manifest", args=[deck.id]))
        self.assertEqual(man_resp.status_code, 200)
        self.assertEqual(man_resp["Content-Type"], "application/json")

        # Download PDF
        pdf_resp = self.client.get(reverse("board_deck:download_pdf", args=[deck.id]))
        self.assertEqual(pdf_resp.status_code, 200)
        self.assertEqual(pdf_resp["Content-Type"], "application/pdf")
