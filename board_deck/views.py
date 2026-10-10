"""
Django Views for Board Deck & Management Accounts Presentation Generator.
"""
import os
import json
from pathlib import Path
from django.shortcuts import render, redirect, get_object_or_404
from django.http import HttpResponse, Http404, JsonResponse, FileResponse
from django.views import View
from django.contrib import messages
from django.utils import timezone
from django.core.files.base import ContentFile
from django.conf import settings

from board_deck.models import (
    Organisation, ReportingPeriod, BrandConfiguration,
    UploadedWorkbook, NormalizedSnapshot, PresentationDeck,
    SlideReviewNarrative, AuditLog
)
from board_deck.forms import WorkbookUploadForm, PresentationConfigForm, SlideReviewEditForm
from services.imports.excel import ExcelImportService
from services.imports.validation import FinancialDataValidator
from services.finance.normalization import normalize_report_payload
from services.finance.metrics import calculate_financial_metrics
from services.presentations.builder import PresentationBuilder, SLIDE_REGISTRY, SLIDE_TITLES
from services.presentations.tokens import DEFAULT_BRAND, MODERN_NAVY_BRAND, MINIMAL_EMERALD_BRAND, BrandConfig
from services.exports.pdf import PDFExportService
from services.commentary.factory import get_commentary_provider


class DashboardView(View):
    def get(self, request):
        decks = PresentationDeck.objects.all().select_related("organisation", "reporting_period")
        snapshots = NormalizedSnapshot.objects.all().select_related("organisation", "reporting_period")
        recent_uploads = UploadedWorkbook.objects.all().select_related("organisation")[:5]

        total_decks = decks.count()
        approved_decks = decks.filter(status="approved").count()
        in_review = decks.filter(status="review").count()
        drafts = decks.filter(status="draft").count()

        context = {
            "decks": decks[:10],
            "snapshots": snapshots[:5],
            "recent_uploads": recent_uploads,
            "total_decks": total_decks,
            "approved_decks": approved_decks,
            "in_review": in_review,
            "drafts": drafts,
            "pdf_available": PDFExportService.is_pdf_conversion_available(),
        }
        return render(request, "board_deck/dashboard.html", context)


class UploadWorkbookView(View):
    def get(self, request):
        form = WorkbookUploadForm()
        return render(request, "board_deck/upload.html", {"form": form})

    def post(self, request):
        form = WorkbookUploadForm(request.POST, request.FILES)
        if not form.is_valid():
            return render(request, "board_deck/upload.html", {"form": form})

        uploaded_file = form.cleaned_data["file"]
        file_bytes = uploaded_file.read()

        # 1. Inspect & Parse Workbook
        try:
            sheet_info = ExcelImportService.inspect_workbook(file_bytes)
            raw_payload = ExcelImportService.parse_workbook(file_bytes)
        except Exception as e:
            messages.error(request, f"Failed to parse Excel workbook: {str(e)}")
            return render(request, "board_deck/upload.html", {"form": form})

        # 2. Validate parsed data
        val_res = FinancialDataValidator.validate_raw_payload(raw_payload)

        # 3. Resolve Organisation & Period
        org_name = form.cleaned_data.get("organisation_name") or raw_payload["organisation"].get("name") or "Apex Technologies Holdings Ltd"
        org, _ = Organisation.objects.get_or_create(
            name=org_name,
            defaults={
                "currency_code": raw_payload["organisation"].get("currency_code", "USD"),
                "currency_symbol": raw_payload["organisation"].get("currency_symbol", "$")
            }
        )

        period_label = form.cleaned_data.get("period_label") or raw_payload["period"].get("label") or "Q3 2026"
        period, _ = ReportingPeriod.objects.get_or_create(
            organisation=org,
            label=period_label,
            defaults={"period_type": raw_payload["period"].get("period_type", "quarter")}
        )

        # 4. Save UploadedWorkbook
        wb_record = UploadedWorkbook.objects.create(
            organisation=org,
            original_filename=uploaded_file.name,
            file_size_bytes=len(file_bytes),
            sheet_count=len(sheet_info),
            uploaded_by=request.user if request.user.is_authenticated else None,
            is_valid=val_res.is_valid,
            validation_errors=[issue.message for issue in val_res.errors],
            validation_warnings=[issue.message for issue in val_res.warnings],
        )
        wb_record.file.save(uploaded_file.name, ContentFile(file_bytes))

        # 5. Save Normalized Snapshot
        snapshot = NormalizedSnapshot.objects.create(
            organisation=org,
            reporting_period=period,
            uploaded_workbook=wb_record,
            data_payload=raw_payload
        )

        AuditLog.objects.create(
            user=request.user if request.user.is_authenticated else None,
            action="UPLOAD_WORKBOOK",
            details={
                "workbook_id": wb_record.id,
                "filename": uploaded_file.name,
                "snapshot_id": snapshot.id,
                "is_valid": val_res.is_valid,
                "errors_count": len(val_res.errors),
                "warnings_count": len(val_res.warnings)
            }
        )

        if not val_res.is_valid:
            messages.error(request, f"Workbook validation failed with {len(val_res.errors)} errors. Please correct issues before generating.")
            return render(request, "board_deck/upload_preview.html", {
                "snapshot": snapshot,
                "wb_record": wb_record,
                "validation": val_res,
                "raw_payload": raw_payload,
                "sheet_info": sheet_info
            })

        messages.success(request, f"Workbook '{uploaded_file.name}' successfully parsed and validated!")
        return redirect("board_deck:create_presentation_from_snapshot", snapshot_id=snapshot.id)


class CreatePresentationView(View):
    def get(self, request, snapshot_id):
        snapshot = get_object_or_404(NormalizedSnapshot, id=snapshot_id)
        form = PresentationConfigForm(initial={
            "title": f"Board Presentation - {snapshot.reporting_period.label}",
            "narrative_style": "standard"
        })
        return render(request, "board_deck/create_presentation.html", {
            "snapshot": snapshot,
            "form": form,
        })

    def post(self, request, snapshot_id):
        snapshot = get_object_or_404(NormalizedSnapshot, id=snapshot_id)
        form = PresentationConfigForm(request.POST)
        if not form.is_valid():
            return render(request, "board_deck/create_presentation.html", {
                "snapshot": snapshot,
                "form": form
            })

        title = form.cleaned_data["title"]
        narrative_style = form.cleaned_data["narrative_style"]
        brand_theme_name = form.cleaned_data["brand_theme"]
        selected_slides = form.cleaned_data["selected_slides"]

        brand = DEFAULT_BRAND
        if brand_theme_name == "Executive Navy":
            brand = MODERN_NAVY_BRAND
        elif brand_theme_name == "Emerald Finance":
            brand = MINIMAL_EMERALD_BRAND

        # 1. Normalize payload to canonical ReportContext
        context = normalize_report_payload(snapshot.data_payload)

        # 2. Build presentation
        builder = PresentationBuilder(brand=brand, narrative_style=narrative_style)
        try:
            prs, manifest, val_out = builder.build_presentation(context, selected_slides=selected_slides)
        except Exception as e:
            messages.error(request, f"Generation failed: {str(e)}")
            return render(request, "board_deck/create_presentation.html", {"snapshot": snapshot, "form": form})

        # 3. Save PPTX to bytes
        pptx_bytes = builder.save_to_bytes(prs)

        # 4. Create PresentationDeck record
        deck = PresentationDeck.objects.create(
            organisation=snapshot.organisation,
            reporting_period=snapshot.reporting_period,
            normalized_snapshot=snapshot,
            title=title,
            narrative_style=narrative_style,
            commentary_provider="template",
            selected_slides=selected_slides,
            status="review",
            version=1,
            generation_manifest=manifest,
            validation_warnings=manifest.get("warnings", [])
        )

        filename = f"{snapshot.organisation.name.replace(' ', '_')}_{snapshot.reporting_period.label.replace(' ', '_')}_Board_Deck_v1.pptx"
        deck.pptx_file.save(filename, ContentFile(pptx_bytes.getvalue()))

        # 5. Populate initial SlideReviewNarrative records
        comm_provider = get_commentary_provider("template")
        commentary = comm_provider.generate_commentary(calculate_financial_metrics(context), style=narrative_style)

        for slide_key in selected_slides:
            sec = commentary.get_section(slide_key)
            if sec:
                SlideReviewNarrative.objects.create(
                    presentation_deck=deck,
                    slide_key=slide_key,
                    title=sec.title,
                    custom_lead_summary=sec.lead_summary,
                    custom_bullets=sec.bullet_points,
                    is_customized=False
                )

        AuditLog.objects.create(
            presentation_deck=deck,
            user=request.user if request.user.is_authenticated else None,
            action="GENERATE_PRESENTATION",
            details={"manifest": manifest}
        )

        messages.success(request, f"Presentation '{title}' generated successfully! Ready for review.")
        return redirect("board_deck:presentation_review", deck_id=deck.id)


class PresentationReviewView(View):
    def get(self, request, deck_id):
        deck = get_object_or_404(PresentationDeck, id=deck_id)
        reviews = deck.slide_reviews.all()
        manifest = deck.generation_manifest
        context_data = normalize_report_payload(deck.normalized_snapshot.data_payload)
        context_data = calculate_financial_metrics(context_data)

        return render(request, "board_deck/review.html", {
            "deck": deck,
            "reviews": reviews,
            "manifest": manifest,
            "context_data": context_data,
            "pdf_available": PDFExportService.is_pdf_conversion_available()
        })


class SaveNarrativeEditView(View):
    def post(self, request, deck_id):
        deck = get_object_or_404(PresentationDeck, id=deck_id)
        slide_key = request.POST.get("slide_key")
        lead_summary = request.POST.get("lead_summary", "").strip()
        bullets_raw = request.POST.get("bullets_text", "").strip()

        bullets = [b.strip().lstrip("•-* ") for b in bullets_raw.split("\n") if b.strip()]

        review_narrative, _ = SlideReviewNarrative.objects.get_or_create(
            presentation_deck=deck,
            slide_key=slide_key,
            defaults={"title": SLIDE_TITLES.get(slide_key, slide_key)}
        )
        review_narrative.custom_lead_summary = lead_summary
        review_narrative.custom_bullets = bullets
        review_narrative.is_customized = True
        review_narrative.save()

        # Invalidate previous approval if modified
        deck.invalidate_approval_on_edit()
        deck.version += 1
        deck.save()

        # Re-build PowerPoint with updated custom commentary
        all_reviews = deck.slide_reviews.all()
        custom_overrides = {}
        for r in all_reviews:
            if r.is_customized:
                custom_overrides[r.slide_key] = {
                    "lead_summary": r.custom_lead_summary,
                    "bullet_points": r.custom_bullets
                }

        context = normalize_report_payload(deck.normalized_snapshot.data_payload)
        brand = DEFAULT_BRAND
        builder = PresentationBuilder(brand=brand, narrative_style=deck.narrative_style)
        prs, manifest, _ = builder.build_presentation(context, selected_slides=deck.selected_slides, custom_commentary=custom_overrides)

        pptx_bytes = builder.save_to_bytes(prs)
        filename = f"{deck.organisation.name.replace(' ', '_')}_{deck.reporting_period.label.replace(' ', '_')}_Board_Deck_v{deck.version}.pptx"
        deck.pptx_file.save(filename, ContentFile(pptx_bytes.getvalue()), save=True)

        AuditLog.objects.create(
            presentation_deck=deck,
            user=request.user if request.user.is_authenticated else None,
            action="EDIT_COMMENTARY",
            details={"slide_key": slide_key, "new_version": deck.version}
        )

        messages.success(request, f"Narrative for slide '{SLIDE_TITLES.get(slide_key, slide_key)}' updated and presentation rebuilt (v{deck.version}).")
        return redirect("board_deck:presentation_review", deck_id=deck.id)


class ApproveDeckView(View):
    def post(self, request, deck_id):
        deck = get_object_or_404(PresentationDeck, id=deck_id)
        deck.status = "approved"
        deck.approved_by = request.user if request.user.is_authenticated else None
        deck.approved_at = timezone.now()
        deck.approved_version = deck.version
        deck.save()

        AuditLog.objects.create(
            presentation_deck=deck,
            user=request.user if request.user.is_authenticated else None,
            action="APPROVE_DECK",
            details={"approved_version": deck.version}
        )

        messages.success(request, f"Presentation '{deck.title}' (v{deck.version}) has been formally APPROVED for board release.")
        return redirect("board_deck:presentation_review", deck_id=deck.id)


class DownloadPPTXView(View):
    def get(self, request, deck_id):
        deck = get_object_or_404(PresentationDeck, id=deck_id)
        if not deck.pptx_file:
            raise Http404("PPTX file not found.")

        AuditLog.objects.create(
            presentation_deck=deck,
            user=request.user if request.user.is_authenticated else None,
            action="DOWNLOAD_PPTX",
            details={"version": deck.version}
        )

        response = FileResponse(deck.pptx_file.open("rb"), content_type="application/vnd.openxmlformats-officedocument.presentationml.presentation")
        response["Content-Disposition"] = f'attachment; filename="{os.path.basename(deck.pptx_file.name)}"'
        return response


class DownloadPDFView(View):
    def get(self, request, deck_id):
        deck = get_object_or_404(PresentationDeck, id=deck_id)
        if not deck.pptx_file:
            raise Http404("Presentation file not found.")

        pptx_full_path = deck.pptx_file.path

        # If PDF already generated and matches version
        if deck.pdf_file and os.path.exists(deck.pdf_file.path):
            AuditLog.objects.create(
                presentation_deck=deck,
                user=request.user if request.user.is_authenticated else None,
                action="DOWNLOAD_PDF",
                details={"version": deck.version}
            )
            response = FileResponse(deck.pdf_file.open("rb"), content_type="application/pdf")
            response["Content-Disposition"] = f'attachment; filename="{os.path.basename(deck.pdf_file.name)}"'
            return response

        # Generate or convert PDF
        pdf_filename = Path(pptx_full_path).stem + ".pdf"
        target_pdf_path = os.path.join(settings.MEDIA_ROOT, "exports", "pdf", pdf_filename)

        success, msg = PDFExportService.generate_pdf_from_deck(deck, target_pdf_path)

        if success:
            with open(target_pdf_path, "rb") as f:
                deck.pdf_file.save(pdf_filename, ContentFile(f.read()), save=True)

            AuditLog.objects.create(
                presentation_deck=deck,
                user=request.user if request.user.is_authenticated else None,
                action="DOWNLOAD_PDF",
                details={"version": deck.version}
            )
            response = FileResponse(open(target_pdf_path, "rb"), content_type="application/pdf")
            response["Content-Disposition"] = f'attachment; filename="{pdf_filename}"'
            return response
        else:
            messages.warning(request, f"PDF export issue: {msg}")
            return redirect("board_deck:presentation_review", deck_id=deck.id)


class DownloadDataSummaryView(View):
    def get(self, request, deck_id):
        deck = get_object_or_404(PresentationDeck, id=deck_id)
        data = deck.normalized_snapshot.data_payload
        response = HttpResponse(json.dumps(data, indent=2, default=str), content_type="application/json")
        response["Content-Disposition"] = f'attachment; filename="data_summary_{deck.reporting_period.label.replace(" ", "_")}.json"'
        return response


class DownloadManifestView(View):
    def get(self, request, deck_id):
        deck = get_object_or_404(PresentationDeck, id=deck_id)
        data = deck.generation_manifest
        response = HttpResponse(json.dumps(data, indent=2, default=str), content_type="application/json")
        response["Content-Disposition"] = f'attachment; filename="generation_manifest_{deck.reporting_period.label.replace(" ", "_")}.json"'
        return response


class SampleWorkbookDownloadView(View):
    def get(self, request):
        sample_path = os.path.join(settings.BASE_DIR, "sample_data", "sample_board_deck_data.xlsx")
        if not os.path.exists(sample_path):
            raise Http404("Sample workbook file not found on server.")
        response = FileResponse(open(sample_path, "rb"), content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        response["Content-Disposition"] = 'attachment; filename="sample_board_deck_data.xlsx"'
        return response
