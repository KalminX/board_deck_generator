"""
Django Forms for Board Deck Generator.
"""
from django import forms
from board_deck.models import Organisation, ReportingPeriod, BrandConfiguration, PresentationDeck


SLIDE_CHOICES = [
    ("cover", "Slide 1: Cover Page"),
    ("executive_summary", "Slide 2: Executive Summary & Highlights"),
    ("financial_overview", "Slide 3: Financial Performance Overview & Charts"),
    ("pnl", "Slide 4: Profit & Loss Statement"),
    ("balance_sheet", "Slide 5: Statement of Financial Position (Balance Sheet)"),
    ("cash_flow", "Slide 6: Statement of Cash Flows"),
    ("kpis", "Slide 7: Key Performance Indicators (KPI Scorecard)"),
    ("risks", "Slide 8: Risks, Anomalies & Material Variances"),
    ("actions", "Slide 9: Board Actions & Management Decisions"),
    ("appendix", "Slide 10: Appendix, Methodology & Data Provenance"),
]


class WorkbookUploadForm(forms.Form):
    organisation_name = forms.CharField(
        max_length=255,
        required=False,
        initial="Apex Technologies Holdings Ltd",
        widget=forms.TextInput(attrs={"class": "form-input", "placeholder": "e.g. Apex Technologies Holdings Ltd"})
    )
    period_label = forms.CharField(
        max_length=50,
        required=False,
        initial="Q3 2026",
        widget=forms.TextInput(attrs={"class": "form-input", "placeholder": "e.g. Q3 2026"})
    )
    file = forms.FileField(
        label="Excel Financial Workbook (.xlsx)",
        help_text="Upload a .xlsx workbook containing P&L, Balance Sheet, Cash Flow, KPIs, and Actions.",
        widget=forms.FileInput(attrs={"class": "form-file-input", "accept": ".xlsx"})
    )

    def clean_file(self):
        uploaded_file = self.cleaned_data.get("file")
        if uploaded_file:
            if not uploaded_file.name.endswith(".xlsx"):
                raise forms.ValidationError("Only Excel .xlsx files are supported.")
            if uploaded_file.size > 15 * 1024 * 1024:
                raise forms.ValidationError("File size exceeds 15 MB limit.")
        return uploaded_file


class PresentationConfigForm(forms.Form):
    title = forms.CharField(
        max_length=255,
        initial="Board of Directors Presentation",
        widget=forms.TextInput(attrs={"class": "form-input"})
    )
    narrative_style = forms.ChoiceField(
        choices=[
            ("concise", "Concise (High-Level Board Pack)"),
            ("standard", "Standard (Executive Management Narrative)"),
            ("detailed", "Detailed (In-Depth Analyst Commentary)")
        ],
        initial="standard",
        widget=forms.Select(attrs={"class": "form-select"})
    )
    brand_theme = forms.ChoiceField(
        choices=[
            ("Corporate Slate", "Corporate Slate (Classic Dark Navy & Slate)"),
            ("Executive Navy", "Executive Navy (Deep Blue & Sky)"),
            ("Emerald Finance", "Emerald Finance (Emerald & Teal)")
        ],
        initial="Corporate Slate",
        widget=forms.Select(attrs={"class": "form-select"})
    )
    selected_slides = forms.MultipleChoiceField(
        choices=SLIDE_CHOICES,
        initial=[k for k, _ in SLIDE_CHOICES],
        widget=forms.CheckboxSelectMultiple(attrs={"class": "form-checkbox"})
    )


class SlideReviewEditForm(forms.Form):
    slide_key = forms.CharField(widget=forms.HiddenInput())
    lead_summary = forms.CharField(
        widget=forms.Textarea(attrs={"class": "form-textarea", "rows": 3}),
        required=False
    )
    bullets_text = forms.CharField(
        widget=forms.Textarea(attrs={"class": "form-textarea", "rows": 6}),
        help_text="One bullet point per line",
        required=False
    )
