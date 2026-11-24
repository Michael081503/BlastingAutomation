from __future__ import annotations

import unittest
from io import BytesIO
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook

from inbound_campaign_lookup import build_email_campaign_lookup
from inbound_summary import campaign_totals_key
from test_inbound_all_channels import REPORT_DATE, Upload, all_channel_uploads
from test_inbound_bucket_lookup import bucket_files
from test_inbound_campaign_lookup import workbook_bytes
from test_inbound_display import table_cells
from test_inbound_email import csv_bytes


ROOT = Path(__file__).resolve().parents[1]


def regular_catalog(last_level=6):
    return Upload("Regular.xlsx", workbook_bytes([("OVERALL", [
        ("Name", "debtor_email", "LEVEL", "phone_1"),
        *((f"Sample {level}", f"level{level}@example.com" if level != last_level else None,
           f"Level {level}", f"094{level}1234567" if level != last_level else None)
          for level in (1, 2, 3, 4, 5, last_level)),
    ])]))


def catalog_uploads():
    geo = all_channel_uploads()["inbound_reply_email_lookup_uploads"][1]
    return [regular_catalog(), geo, *(Upload(name, data) for name, data in bucket_files())]


class CampaignCatalogTests(unittest.TestCase):
    def test_campaigns_come_from_account_rows_even_without_usable_contacts(self):
        uploads = catalog_uploads()
        lookup = build_email_campaign_lookup([(u.name, u.getvalue()) for u in uploads])
        expected = ("SBC B2", "SBC B4", "SBC Level 1", "SBC Level 6", "SBC Mindanao")
        self.assertEqual(lookup.campaigns, expected)
        self.assertFalse(any("SBC Level 6" in labels for labels in lookup.campaigns_by_email.values()))
        self.assertFalse(any("SBC Level 6" in labels for labels in lookup.campaigns_by_phone.values()))
        reverse = build_email_campaign_lookup([(u.name, u.getvalue()) for u in reversed(uploads)])
        self.assertEqual(reverse.campaigns, expected)

    def test_empty_bucket_files_do_not_invent_an_active_campaign(self):
        lookup = build_email_campaign_lookup([("B4.csv", csv_bytes(("Email (01)", "Mobile 1"), []))])
        self.assertEqual(lookup.campaigns, ())


class CampaignCatalogPageTests(unittest.TestCase):
    def test_summary_and_excel_catalog_refresh_when_uploads_change(self):
        from streamlit.testing.v1 import AppTest
        from upload_state import render_clear_all_files

        expected = ["SBC B2", "SBC B4", "SBC Level 1", "SBC Level 6", "SBC Mindanao"]
        files = catalog_uploads()
        uploads = {}
        with patch("branding.render_bank_branding", side_effect=render_clear_all_files), \
                patch("inbound_reply_page.inbound_report_date", return_value=REPORT_DATE), \
                patch("streamlit.file_uploader", side_effect=lambda label, **kw: uploads.get(kw["key"])), \
                patch("streamlit.html") as html, patch("streamlit.download_button") as downloads:
            app = AppTest.from_file(str(ROOT / "pages" / "INBOUND_REPLY.py")).run(timeout=20)

            def summary():
                self.assertFalse(app.exception)
                call = next(call for call in reversed(html.call_args_list) if 'aria-label="Inbound reply summary"' in call.args[0])
                return {row[0]: row for row in table_cells(call.args[0])}

            self.assertEqual(list(summary()), ["\u2014"])
            self.assertEqual(len(app.number_input), 0)
            self.assertFalse(downloads.called)
            uploads["inbound_reply_email_lookup_uploads"] = files
            app.run(timeout=20)
            self.assertFalse(app.error)
            self.assertEqual(list(summary()), expected)
            self.assertEqual(len(app.number_input), 0)
            self.assertEqual(len(app.metric), 0)
            call = next(call for call in reversed(downloads.call_args_list) if call.kwargs.get("key") == "inbound_reply_summary_download")
            with BytesIO(call.kwargs["data"]()) as data:
                book = load_workbook(data)
            try:
                self.assertEqual([book["Summary"].cell(r, 1).value for r in range(4, 4 + len(expected))], expected)
                self.assertIsNone(book["Summary"]["B6"].value)
                self.assertIsNone(book["Summary"]["C6"].value)
                self.assertEqual(book["Summary"]["D6"].data_type, "f")
                self.assertIsNone(book["Summary"]["I6"].value)
            finally:
                book.close()

            app.session_state["inbound_reply_campaign_totals"] = {
                campaign_totals_key(REPORT_DATE, "SBC Mindanao"): {"active_ticket": 400},
                campaign_totals_key(REPORT_DATE, "SBC Region 3"): {"active_ticket": 300, "dc_email": 140},
            }
            uploads["inbound_reply_email_lookup_uploads"] = list(reversed(files))
            app.run(timeout=20)
            self.assertEqual(list(summary()), expected)
            self.assertEqual(summary()["SBC Mindanao"][1:4], ["700", "140", "20.00%"])

            uploads["inbound_reply_email_lookup_uploads"] = [u for u in files if u.name != "B4.xlsx"]
            app.run(timeout=20)
            self.assertNotIn("SBC B4", summary())
            uploads["inbound_reply_email_lookup_uploads"] = [regular_catalog(7), *files[1:]]
            app.run(timeout=20)
            self.assertNotIn("SBC Level 6", summary())
            self.assertIn("SBC Level 7", summary())

            uploads["inbound_reply_email_lookup_uploads"] = files
            uploads["inbound_reply_email_2_upload"] = Upload("dc.csv", csv_bytes(("sender",), [("level1@example.com",)]))
            app.run(timeout=20)
            self.assertEqual(list(summary()), expected)
            self.assertEqual(summary()["SBC Level 1"][8], "\u2014")
            uploads["inbound_reply_email_1_upload"] = Upload("history.csv", csv_bytes(("email", "sent"), [
                ("level1@example.com", "Agent reply"),
            ]))
            app.run(timeout=20)
            self.assertEqual(list(summary()), expected)
            self.assertEqual(summary()["SBC Level 1"][8], "1")
            app.button(key="clear_all_files").click().run(timeout=20)
            self.assertEqual(list(summary()), ["\u2014"])


if __name__ == "__main__":
    unittest.main()
