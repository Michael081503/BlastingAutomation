from __future__ import annotations

import re
import unittest
from hashlib import sha256
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


def total_key(campaign, field):
    identity = campaign_totals_key(REPORT_DATE, campaign)
    return "inbound_reply_total_" + sha256(identity.encode("utf-8")).hexdigest()[:16] + "_" + field


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


class CampaignEditorPageTests(unittest.TestCase):
    def test_lookup_only_uploads_label_manual_fields_refresh_campaigns_and_preserve_correct_totals(self):
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

            def editor_labels():
                self.assertFalse(app.exception)
                return [match[1] for element in app.markdown
                        if (match := re.fullmatch(r"\*\*(SBC (?:B[24]|Level \d+|Mindanao|Region 3))\*\*", element.value))]

            def summary():
                call = next(call for call in reversed(html.call_args_list) if 'aria-label="Inbound reply summary"' in call.args[0])
                return {row[0]: row for row in table_cells(call.args[0])}

            self.assertEqual(editor_labels(), [])
            self.assertEqual(len(app.number_input), 0)
            self.assertTrue(any("Upload campaign lookup files" in message.value for message in app.info))
            uploads["inbound_reply_email_lookup_uploads"] = files
            app.run(timeout=20)
            self.assertFalse(app.error)
            self.assertEqual(editor_labels(), expected)
            self.assertEqual(list(summary()), expected)
            self.assertEqual(len(app.number_input), 4 * len(expected))
            self.assertTrue(all(widget.value is None for widget in app.number_input))
            self.assertEqual(len(app.metric), 0)
            self.assertTrue(any("5 campaigns detected" in caption.value for caption in app.caption))
            for field, value in (("active_ticket", 5000), ("dc_email", 4337), ("dc_sms", 3500), ("dc_viber", 2500)):
                app.number_input(key=total_key("SBC Level 1", field)).set_value(value).run(timeout=20)
                self.assertFalse(app.exception)
            self.assertEqual(summary()["SBC Level 1"][1:8], ["5,000", "4,337", "86.74%", "3,500", "70.00%", "2,500", "50.00%"])
            self.assertTrue(any("Email: 86.74% · SMS: 70.00% · Viber: 50.00%" in c.value for c in app.caption))
            app.number_input(key=total_key("SBC B4", "active_ticket")).set_value(3000).run(timeout=20)
            app.number_input(key=total_key("SBC Level 6", "active_ticket")).set_value(6000).run(timeout=20)
            mindanao_active = total_key("SBC Mindanao", "active_ticket")
            app.number_input(key=mindanao_active).set_value(400).run(timeout=20)
            legacy_key = campaign_totals_key(REPORT_DATE, "SBC Region 3")
            app.session_state["inbound_reply_campaign_totals"][legacy_key] = {"active_ticket": 300, "dc_email": 140}
            uploads["inbound_reply_email_lookup_uploads"] = list(reversed(files))
            app.run(timeout=20)
            self.assertEqual(editor_labels(), expected)
            self.assertEqual(app.number_input(key=mindanao_active).value, 700)
            self.assertEqual(summary()["SBC Mindanao"][1:4], ["700", "140", "20.00%"])
            self.assertNotIn(legacy_key, app.session_state["inbound_reply_campaign_totals"])
            self.assertEqual(app.number_input(key=total_key("SBC Level 1", "active_ticket")).value, 5000)
            self.assertEqual(app.number_input(key=total_key("SBC B4", "active_ticket")).value, 3000)
            uploads["inbound_reply_email_lookup_uploads"] = [u for u in files if u.name != "B4.xlsx"]
            app.run(timeout=20)
            self.assertEqual(editor_labels(), [label for label in expected if label != "SBC B4"])
            self.assertNotIn("SBC B4", summary())
            uploads["inbound_reply_email_lookup_uploads"] = files
            app.run(timeout=20)
            self.assertEqual(app.number_input(key=total_key("SBC B4", "active_ticket")).value, 3000)
            uploads["inbound_reply_email_lookup_uploads"] = [regular_catalog(7), *files[1:]]
            app.run(timeout=20)
            self.assertNotIn("SBC Level 6", editor_labels())
            self.assertIn("SBC Level 7", editor_labels())
            self.assertIsNone(app.number_input(key=total_key("SBC Level 7", "active_ticket")).value)
            uploads["inbound_reply_email_lookup_uploads"] = files
            uploads["inbound_reply_email_2_upload"] = Upload("dc.csv", csv_bytes(("sender",), [("level1@example.com",)]))
            app.run(timeout=20)
            self.assertEqual(editor_labels(), expected)
            self.assertEqual(app.number_input(key=total_key("SBC Level 6", "active_ticket")).value, 6000)
            self.assertEqual(summary()["SBC Level 1"][8], "—")
            uploads["inbound_reply_email_1_upload"] = Upload("history.csv", csv_bytes(("email", "sent"), [("level1@example.com", "Agent reply")]))
            app.run(timeout=20)
            self.assertEqual(editor_labels(), expected)
            self.assertEqual(summary()["SBC Level 1"][8], "1")
            call = next(call for call in reversed(downloads.call_args_list) if call.kwargs.get("key") == "inbound_reply_summary_download")
            book = load_workbook(BytesIO(call.kwargs["data"]()))
            try:
                self.assertEqual([book["Summary"].cell(r, 1).value for r in range(4, 4 + len(expected))], expected)
                self.assertEqual(book["Summary"]["B5"].value, 3000)
                self.assertEqual(book["Summary"]["B6"].value, 5000)
                self.assertEqual(book["Summary"]["C6"].value, 4337)
                self.assertEqual(book["Summary"]["D6"].data_type, "f")
                self.assertEqual(book["Summary"]["I6"].value, 1)
                self.assertIsNone(book["Summary"]["I7"].value)
                self.assertEqual([book["Summary"][cell].value for cell in ("B8", "C8")], [700, 140])
            finally:
                book.close()
            app.button(key="clear_all_files").click().run(timeout=20)
            self.assertEqual(editor_labels(), [])
            self.assertEqual(len(app.number_input), 0)
            uploads["inbound_reply_email_lookup_uploads_clear_1"] = files
            app.run(timeout=20)
            self.assertEqual(editor_labels(), expected)
            self.assertEqual(app.number_input(key=total_key("SBC Level 1", "active_ticket")).value, 5000)
            self.assertEqual(app.number_input(key=total_key("SBC B4", "active_ticket")).value, 3000)


if __name__ == "__main__":
    unittest.main()
