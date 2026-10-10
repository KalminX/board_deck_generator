"""
Security and File Upload Integrity Tests.
"""
from django.test import TestCase, Client
from django.urls import reverse
from django.core.files.uploadedfile import SimpleUploadedFile


class SecurityTestCase(TestCase):
    def setUp(self):
        self.client = Client()

    def test_reject_non_xlsx_upload(self):
        fake_file = SimpleUploadedFile("malicious_script.exe", b"binary content", content_type="application/octet-stream")
        response = self.client.post(reverse("board_deck:upload"), {
            "organisation_name": "Test Org",
            "period_label": "Q3 2026",
            "file": fake_file
        })
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Only Excel .xlsx files are supported.")

    def test_reject_oversized_file(self):
        # 16 MB dummy content
        large_content = b"0" * (16 * 1024 * 1024)
        large_file = SimpleUploadedFile("huge.xlsx", large_content, content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        response = self.client.post(reverse("board_deck:upload"), {
            "organisation_name": "Test Org",
            "period_label": "Q3 2026",
            "file": large_file
        })
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "File size exceeds 15 MB limit.")
