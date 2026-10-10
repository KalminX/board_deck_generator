"""
Database Models for Board Deck & Management Accounts Generator.
Persists organisations, reporting periods, uploaded workbooks, normalized snapshots,
presentations, slide reviews, and audit trails.
"""
from django.db import models
from django.contrib.auth.models import User
from django.utils import timezone
from django.core.serializers.json import DjangoJSONEncoder
from pptx.dml.color import RGBColor
import hashlib
import json


class Organisation(models.Model):
    name = models.CharField(max_length=255, default="Apex Technologies Holdings Ltd")
    legal_entity_number = models.CharField(max_length=100, blank=True, null=True)
    industry = models.CharField(max_length=100, default="Technology & Software")
    currency_code = models.CharField(max_length=10, default="USD")
    currency_symbol = models.CharField(max_length=5, default="$")
    fiscal_year_end_month = models.IntegerField(default=12)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return self.name


class ReportingPeriod(models.Model):
    organisation = models.ForeignKey(Organisation, on_delete=models.CASCADE, related_name="reporting_periods")
    label = models.CharField(max_length=50, default="Q3 2026")  # e.g., "Q3 2026", "Sep 2026", "FY 2026"
    period_type = models.CharField(max_length=20, choices=[
        ("month", "Month"),
        ("quarter", "Quarter"),
        ("year", "Full Year")
    ], default="quarter")
    start_date = models.DateField(blank=True, null=True)
    end_date = models.DateField(blank=True, null=True)
    prior_period_label = models.CharField(max_length=50, default="Q2 2026")
    budget_period_label = models.CharField(max_length=50, default="Q3 2026 Budget")
    is_closed = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["-created_at"]

    def __str__(self):
        return f"{self.organisation.name} - {self.label}"


class BrandConfiguration(models.Model):
    name = models.CharField(max_length=100, default="Corporate Slate")
    primary_color_hex = models.CharField(max_length=7, default="#0F172A")
    secondary_color_hex = models.CharField(max_length=7, default="#334155")
    accent_color_hex = models.CharField(max_length=7, default="#2563EB")
    font_title = models.CharField(max_length=50, default="Calibri")
    font_body = models.CharField(max_length=50, default="Calibri")
    is_default = models.BooleanField(default=False)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return self.name

    def to_brand_config(self):
        from services.presentations.tokens import BrandConfig

        def hex_to_rgb(h):
            h = h.lstrip("#")
            return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

        return BrandConfig(
            brand_name=self.name,
            primary_color=hex_to_rgb(self.primary_color_hex),
            secondary_color=hex_to_rgb(self.secondary_color_hex),
            accent_color=hex_to_rgb(self.accent_color_hex),
            font_title=self.font_title,
            font_body=self.font_body
        )


class UploadedWorkbook(models.Model):
    organisation = models.ForeignKey(Organisation, on_delete=models.CASCADE, related_name="uploaded_workbooks")
    original_filename = models.CharField(max_length=255)
    file = models.FileField(upload_to="uploads/workbooks/%Y/%m/")
    file_size_bytes = models.IntegerField(default=0)
    sheet_count = models.IntegerField(default=0)
    uploaded_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True)
    is_valid = models.BooleanField(default=False)
    validation_errors = models.JSONField(default=list, blank=True, encoder=DjangoJSONEncoder)
    validation_warnings = models.JSONField(default=list, blank=True, encoder=DjangoJSONEncoder)
    uploaded_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.original_filename} ({self.uploaded_at.strftime('%Y-%m-%d')})"


class NormalizedSnapshot(models.Model):
    organisation = models.ForeignKey(Organisation, on_delete=models.CASCADE, related_name="snapshots")
    reporting_period = models.ForeignKey(ReportingPeriod, on_delete=models.CASCADE, related_name="snapshots")
    uploaded_workbook = models.ForeignKey(UploadedWorkbook, on_delete=models.SET_NULL, null=True, blank=True, related_name="snapshots")
    data_payload = models.JSONField(help_text="Canonical report payload dictionary", encoder=DjangoJSONEncoder)
    snapshot_hash = models.CharField(max_length=64, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    def save(self, *args, **kwargs):
        if not self.snapshot_hash:
            data_str = json.dumps(self.data_payload, sort_keys=True, default=str)
            self.snapshot_hash = hashlib.sha256(data_str.encode("utf-8")).hexdigest()
        super().save(*args, **kwargs)

    def __str__(self):
        return f"Snapshot {self.organisation.name} - {self.reporting_period.label} ({self.created_at.strftime('%Y-%m-%d')})"


class PresentationDeck(models.Model):
    STATUS_CHOICES = [
        ("draft", "Draft"),
        ("generating", "Generating"),
        ("review", "In Review"),
        ("approved", "Approved"),
        ("failed", "Failed")
    ]

    organisation = models.ForeignKey(Organisation, on_delete=models.CASCADE, related_name="decks")
    reporting_period = models.ForeignKey(ReportingPeriod, on_delete=models.CASCADE, related_name="decks")
    normalized_snapshot = models.ForeignKey(NormalizedSnapshot, on_delete=models.CASCADE, related_name="decks")
    brand_config = models.ForeignKey(BrandConfiguration, on_delete=models.SET_NULL, null=True, blank=True)

    title = models.CharField(max_length=255, default="Board of Directors Presentation")
    narrative_style = models.CharField(max_length=20, choices=[
        ("concise", "Concise (Board Summary)"),
        ("standard", "Standard (Management Narrative)"),
        ("detailed", "Detailed (Analyst Deep Dive)")
    ], default="standard")
    commentary_provider = models.CharField(max_length=50, default="template")
    selected_slides = models.JSONField(default=list, encoder=DjangoJSONEncoder)

    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default="draft")
    version = models.IntegerField(default=1)

    pptx_file = models.FileField(upload_to="exports/pptx/%Y/%m/", blank=True, null=True)
    pdf_file = models.FileField(upload_to="exports/pdf/%Y/%m/", blank=True, null=True)

    generation_manifest = models.JSONField(default=dict, blank=True, encoder=DjangoJSONEncoder)
    validation_warnings = models.JSONField(default=list, blank=True, encoder=DjangoJSONEncoder)

    reviewed_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name="reviewed_decks")
    reviewed_at = models.DateTimeField(blank=True, null=True)
    approved_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name="approved_decks")
    approved_at = models.DateTimeField(blank=True, null=True)
    approved_version = models.IntegerField(blank=True, null=True)

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["-created_at"]

    def __str__(self):
        return f"{self.title} - {self.reporting_period.label} (v{self.version})"

    def invalidate_approval_on_edit(self):
        """If deck is edited after approval, revoke approval status."""
        if self.status == "approved":
            self.status = "review"
            self.approved_by = None
            self.approved_at = None
            self.approved_version = None


class SlideReviewNarrative(models.Model):
    presentation_deck = models.ForeignKey(PresentationDeck, on_delete=models.CASCADE, related_name="slide_reviews")
    slide_key = models.CharField(max_length=50)
    title = models.CharField(max_length=255)
    custom_lead_summary = models.TextField(blank=True)
    custom_bullets = models.JSONField(default=list, blank=True, encoder=DjangoJSONEncoder)
    is_customized = models.BooleanField(default=False)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        unique_together = ("presentation_deck", "slide_key")

    def __str__(self):
        return f"{self.presentation_deck.title} - Slide: {self.slide_key}"


class AuditLog(models.Model):
    presentation_deck = models.ForeignKey(PresentationDeck, on_delete=models.SET_NULL, null=True, blank=True, related_name="audit_logs")
    user = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True)
    action = models.CharField(max_length=50)  # e.g., 'UPLOAD', 'GENERATE', 'EDIT_NARRATIVE', 'APPROVE', 'REJECT', 'DOWNLOAD_PPTX', 'DOWNLOAD_PDF'
    details = models.JSONField(default=dict, blank=True, encoder=DjangoJSONEncoder)
    timestamp = models.DateTimeField(default=timezone.now)

    class Meta:
        ordering = ["-timestamp"]

    def __str__(self):
        return f"[{self.timestamp.strftime('%Y-%m-%d %H:%M')}] {self.action} by {self.user or 'System'}"
