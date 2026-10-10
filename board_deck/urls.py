"""
URL Routing for board_deck app.
"""
from django.urls import path
from board_deck.views import (
    DashboardView, UploadWorkbookView, CreatePresentationView,
    PresentationReviewView, SaveNarrativeEditView, ApproveDeckView,
    DownloadPPTXView, DownloadPDFView, DownloadDataSummaryView,
    DownloadManifestView, SampleWorkbookDownloadView
)

app_name = "board_deck"

urlpatterns = [
    path("", DashboardView.as_view(), name="dashboard"),
    path("upload/", UploadWorkbookView.as_view(), name="upload"),
    path("generate/<int:snapshot_id>/", CreatePresentationView.as_view(), name="create_presentation_from_snapshot"),
    path("deck/<int:deck_id>/review/", PresentationReviewView.as_view(), name="presentation_review"),
    path("deck/<int:deck_id>/save-narrative/", SaveNarrativeEditView.as_view(), name="save_narrative_edit"),
    path("deck/<int:deck_id>/approve/", ApproveDeckView.as_view(), name="approve_deck"),
    path("deck/<int:deck_id>/download/pptx/", DownloadPPTXView.as_view(), name="download_pptx"),
    path("deck/<int:deck_id>/download/pdf/", DownloadPDFView.as_view(), name="download_pdf"),
    path("deck/<int:deck_id>/download/data/", DownloadDataSummaryView.as_view(), name="download_data_summary"),
    path("deck/<int:deck_id>/download/manifest/", DownloadManifestView.as_view(), name="download_manifest"),
    path("download/sample-workbook/", SampleWorkbookDownloadView.as_view(), name="download_sample_workbook"),
]
