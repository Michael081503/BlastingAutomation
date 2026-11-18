from __future__ import annotations

import unittest
from datetime import date
from hashlib import sha256
from io import BytesIO
from pathlib import Path
from unittest.mock import patch

from inbound_email import read_email_source
from inbound_reply import read_reply_source
from inbound_summary import campaign_totals_key, summary_rows
from test_inbound_display import table_cells
from test_inbound_email import csv_bytes
from test_inbound_summary import reports


ROOT = Path(__file__).resolve().parents[1]


class CampaignColumnTests(unittest.TestCase):
    def test_campaign_header_variants_work_in_excel_phone_and_email_files(self):
        from openpyxl import Workbook

        for header in ("Campaign", "Campaigns", "Bank Campaign", "Bank Campaigns", "Campaign Name"):
            for channel in ("phone", "email"):
                with self.subTest(header=header, channel=channel):
                    book = Workbook()
                    sheet = book.active
                    sheet.append(["Phone" if channel == "phone" else "sender", header])
                    sheet.append(["09171234567" if channel == "phone" else "one@example.com", "Source Campaign"])
                    data = BytesIO()
                    book.save(data)
                    book.close()
                    source = (read_reply_source("customers.xlsx", data.getvalue()) if channel == "phone"
                              else read_email_source("customers.xlsx", data.getvalue(), kind="dc"))
                    self.assertEqual([record.campaign for record in source.records], ["Source Campaign"])

    def test_source_campaign_names_merge_with_processed_campaigns_without_inventing_counts(self):
        viber, email = reports()
        rows = summary_rows("Manual fallback", date(2026, 11, 18), "YES", viber, email,
                            source_campaigns=(" a ", "C", "c", ""))
        self.assertEqual([row[0] for row in rows], ["A", "B", "C"])
        self.assertEqual(rows[0][8], 2)
        self.assertTrue(all(rows[2][index] is None for index in (*range(1, 11), *range(12, 18))))

    def test_campaign_column_populates_from_single_uploads_and_stays_in_sync_with_exports(self):
        from openpyxl import load_workbook
        from streamlit.testing.v1 import AppTest
        from upload_state import render_clear_all_files

        class Upload(BytesIO):
            def __init__(self, name, data):
                super().__init__(data)
                self.name = name

        uploads = {
            "inbound_reply_sms_2_upload": Upload("sms.csv", csv_bytes(("Phone", "Campaigns"), [
                ("09171234567", "SMS Campaign"), ("09181234567", "SMS Campaign"),
            ])),
            "inbound_reply_viber_2_upload": Upload("viber.csv", csv_bytes(("Phone", "Bank Campaigns"), [
                ("09191234567", "Viber Campaign"),
            ])),
            "inbound_reply_email_2_upload": Upload("dc.csv", csv_bytes(("sender", "Campaign Name"), [
                ("one@example.com", "Email Campaign"),
            ])),
        }
        with patch("branding.render_bank_branding", side_effect=render_clear_all_files), \
                patch("streamlit.file_uploader", side_effect=lambda label, **kwargs: uploads.get(kwargs["key"])), \
                patch("streamlit.html") as html, patch("streamlit.download_button") as downloads:
            app = AppTest.from_file(str(ROOT / "pages" / "INBOUND_REPLY.py")).run(timeout=20)

            def displayed_rows():
                self.assertFalse(app.exception)
                call = next(call for call in reversed(html.call_args_list)
                            if 'aria-label="Inbound reply summary"' in call.args[0])
                return table_cells(call.args[0])

            rows = displayed_rows()
            self.assertEqual([row[0] for row in rows], ["Email Campaign", "SMS Campaign", "Viber Campaign"])
            self.assertTrue(all(row[8:11] == ["—", "—", "—"] for row in rows))
            identity = campaign_totals_key(app.date_input(key="inbound_reply_date").value, "SMS Campaign")
            active_key = "inbound_reply_total_" + sha256(identity.encode("utf-8")).hexdigest()[:16] + "_active_ticket"
            app.number_input(key=active_key).set_value(100).run(timeout=20)
            call = next(call for call in reversed(downloads.call_args_list)
                        if call.kwargs.get("key") == "inbound_reply_summary_download")
            book = load_workbook(BytesIO(call.kwargs["data"]()))
            try:
                self.assertEqual([book["Summary"].cell(row, 1).value for row in range(4, 7)],
                                 ["Email Campaign", "SMS Campaign", "Viber Campaign"])
                self.assertEqual(book["Summary"]["B5"].value, 100)
                self.assertIsNone(book["Summary"]["J5"].value)
                self.assertIn("Incomplete", book["SMS Unreplied"]["B1"].value)
                self.assertIsNone(book["SMS Unreplied"]["B2"].value)
            finally:
                book.close()
            uploads["inbound_reply_sms_1_upload"] = Upload("agents.csv", csv_bytes(("Phone", "Campaigns"), [
                ("09171234567", "SMS Campaign"),
            ]))
            app.run(timeout=20)
            rows = displayed_rows()
            sms_row = next(row for row in rows if row[0] == "SMS Campaign")
            self.assertEqual(sms_row[1], "100")
            self.assertEqual(sms_row[9], "2")
            self.assertEqual(sms_row[14:16], ["1", "50.00%"])
            uploads["inbound_reply_email_2_upload"] = Upload("dc.csv", csv_bytes(("sender", "Campaign Name"), [
                ("two@example.com", "Replacement Campaign"),
            ]))
            del uploads["inbound_reply_viber_2_upload"]
            app.run(timeout=20)
            self.assertEqual([row[0] for row in displayed_rows()], ["Replacement Campaign", "SMS Campaign"])
            app.button(key="clear_all_files").click().run(timeout=20)
            self.assertEqual([row[0] for row in displayed_rows()], ["—"])
            self.assertNotIn("inbound_reply_campaign_labels", app.session_state)


if __name__ == "__main__":
    unittest.main()
