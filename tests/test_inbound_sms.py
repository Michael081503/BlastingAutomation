from __future__ import annotations

import unittest
from datetime import date
from io import BytesIO
from pathlib import Path
from unittest.mock import patch

from inbound_reply import match_viber_sources, read_reply_source
from inbound_summary import build_summary_workbook
from test_inbound_email import DC_HEADERS, HISTORY_HEADERS, csv_bytes
from test_inbound_summary import OUTPUT_SHEETS, detail_rows, reports


ROOT = Path(__file__).resolve().parents[1]
PHONE_HEADERS = ("ID", "Client Name", "Phone", "Bank Campaign", "Message Content")


def sms_files():
    return (
        csv_bytes(PHONE_HEADERS, [
            ("000007", "Replied", "09191234567", "A", "Account message"),
            ("000008", "Pending", "09201234567", "A", "First source message"),
            ("000009", "Other Name", "+63 920 123 4567", "A", "Duplicate phone"),
            ("000010", "Pending", "09201234567", "B", "Separate campaign"),
        ]),
        csv_bytes(PHONE_HEADERS, [("000011", "Replied", "+63 919 123 4567", "A", "Agent reply")]),
    )


class SmsOutputTests(unittest.TestCase):
    def test_sms_only_upload_accepts_a_fully_replied_account_list(self):
        from openpyxl import load_workbook
        from streamlit.testing.v1 import AppTest

        class Upload(BytesIO):
            def __init__(self, name, data):
                super().__init__(data)
                self.name = name

        data = csv_bytes(PHONE_HEADERS, [("0001", "One", "09171234567", "A", "Message")])
        uploads = {"inbound_reply_sms_1_upload": Upload("accounts.csv", data),
                   "inbound_reply_sms_2_upload": Upload("replied.csv", data)}
        with patch("branding.render_bank_branding"), \
                patch("streamlit.file_uploader", side_effect=lambda label, **kwargs: uploads.get(kwargs["key"])), \
                patch("streamlit.download_button") as downloads:
            app = AppTest.from_file(str(ROOT / "pages" / "INBOUND_REPLY.py")).run(timeout=20)
            self.assertFalse(app.exception)
            self.assertFalse(app.error)
            call = next(call for call in downloads.call_args_list if call.kwargs.get("key") == "inbound_reply_summary_download")
            book = load_workbook(BytesIO(call.kwargs["data"]()))
            try:
                self.assertEqual(book["SMS Unreplied"]["B1"].value, "Processed")
                self.assertEqual(book["SMS Unreplied"]["B2"].value, 0)
                self.assertEqual(detail_rows(book["SMS Unreplied"]), [])
                self.assertEqual(book["Summary"]["J4"].value, 1)
                self.assertEqual(book["Summary"]["O4"].value, 1)
            finally:
                book.close()

    def test_sms_viber_and_email_have_separate_complete_unreplied_sheets(self):
        from openpyxl import load_workbook

        viber, email = reports()
        accounts, replies = sms_files()
        sms = match_viber_sources(read_reply_source("accounts.csv", accounts), read_reply_source("replied.csv", replies),
                                  agent_replied_source="campaign_replies", match_mode="phone")
        book = load_workbook(BytesIO(build_summary_workbook("", date(2026, 11, 18), "YES", viber, email, sms_report=sms)))
        try:
            self.assertEqual(book.sheetnames, OUTPUT_SHEETS)
            self.assertEqual(book["Summary"]["J4"].value, 2)
            self.assertEqual(book["Summary"]["O4"].value, 1)
            self.assertEqual(book["Summary"]["P4"].value, .5)
            for name, expected, channel in (("SMS Unreplied", 2, "SMS"), ("Viber Unreplied", 1, "VIBER"),
                                              ("Email Unreplied", 1, "EMAIL")):
                sheet = book[name]
                self.assertEqual(sheet["B1"].value, "Processed")
                self.assertEqual(sheet["B2"].value, expected)
                rows = detail_rows(sheet)
                self.assertEqual(len(rows), expected)
                self.assertTrue(all(row["Channel"] == channel and row["Reply Status"] == "Not replied" for row in rows))
            sms_rows = detail_rows(book["SMS Unreplied"])
            self.assertEqual({(row["Campaign"], row["Phone"]) for row in sms_rows},
                             {("A", "09201234567"), ("B", "09201234567")})
            self.assertEqual({row["Source: ID"] for row in sms_rows}, {"000008", "000010"})
            self.assertEqual(sms_rows[0]["Source: Message Content"], "First source message")
            self.assertEqual(book["SMS Unreplied"]["D4"].number_format, "@")
        finally:
            book.close()

    def test_processed_zero_is_distinct_from_missing_channel_inputs(self):
        from openpyxl import load_workbook

        data = csv_bytes(PHONE_HEADERS, [("0001", "One", "09171234567", "A", "Message")])
        source = read_reply_source("accounts.csv", data)
        sms = match_viber_sources(source, source, agent_replied_source="campaign_replies", match_mode="phone")
        book = load_workbook(BytesIO(build_summary_workbook("", date(2026, 11, 18), "NO", sms_report=sms)))
        try:
            self.assertEqual(book["SMS Unreplied"]["B1"].value, "Processed")
            self.assertEqual(book["SMS Unreplied"]["B2"].value, 0)
            self.assertEqual(detail_rows(book["SMS Unreplied"]), [])
            self.assertIn("Source: ID", [cell.value for cell in book["SMS Unreplied"][3]])
            for name in ("Viber Unreplied", "Email Unreplied"):
                self.assertEqual(book[name]["B1"].value, "Not uploaded")
                self.assertIsNone(book[name]["B2"].value)
                self.assertEqual(detail_rows(book[name]), [])
        finally:
            book.close()

    def test_sms_uploads_replacements_and_clear_update_the_output_without_stale_accounts(self):
        from openpyxl import load_workbook
        from streamlit.testing.v1 import AppTest
        from upload_state import render_clear_all_files

        class Upload(BytesIO):
            def __init__(self, name, data):
                super().__init__(data)
                self.name = name

        accounts, replies = sms_files()
        uploads = {
            "inbound_reply_sms_1_upload": Upload("accounts.csv", accounts),
            "inbound_reply_email_1_upload": Upload("history.csv", csv_bytes(HISTORY_HEADERS, [])),
            "inbound_reply_email_2_upload": Upload("DC.csv", csv_bytes(DC_HEADERS, [
                ("E1", "base@example.com", "Respond Now", "Email Base", None, None),
            ])),
        }
        with patch("branding.render_bank_branding", side_effect=render_clear_all_files), \
                patch("streamlit.file_uploader", side_effect=lambda label, **kwargs: uploads.get(kwargs["key"])), \
                patch("streamlit.download_button") as downloads:
            app = AppTest.from_file(str(ROOT / "pages" / "INBOUND_REPLY.py")).run(timeout=20)

            def output_book():
                self.assertFalse(app.exception)
                call = next(call for call in reversed(downloads.call_args_list)
                            if call.kwargs.get("key") == "inbound_reply_summary_download")
                return load_workbook(BytesIO(call.kwargs["data"]()))

            book = output_book()
            self.assertIn("Incomplete", book["SMS Unreplied"]["B1"].value)
            self.assertIsNone(book["SMS Unreplied"]["B2"].value)
            book.close()
            uploads["inbound_reply_sms_2_upload"] = Upload("replied.csv", replies)
            app.run(timeout=20)
            book = output_book()
            self.assertEqual(book["SMS Unreplied"]["B2"].value, 2)
            self.assertEqual(len(detail_rows(book["SMS Unreplied"])), 2)
            book.close()
            self.assertIn("report", app.session_state["inbound_reply_sms_result"])
            uploads["inbound_reply_sms_2_upload"] = Upload("replied.csv", b"Wrong,Headers\n1,2\n")
            app.run(timeout=20)
            self.assertTrue(app.error)
            book = output_book()
            self.assertIn("Could not process", book["SMS Unreplied"]["B1"].value)
            self.assertIsNone(book["SMS Unreplied"]["B2"].value)
            self.assertEqual(detail_rows(book["SMS Unreplied"]), [])
            self.assertEqual(book["Email Unreplied"]["B2"].value, 1)
            book.close()
            self.assertNotIn("report", app.session_state["inbound_reply_sms_result"])
            del uploads["inbound_reply_sms_1_upload"]
            app.run(timeout=20)
            self.assertNotIn("inbound_reply_sms_result", app.session_state)
            downloads.reset_mock()
            app.button(key="clear_all_files").click().run(timeout=20)
            self.assertFalse(app.exception)
            self.assertNotIn("inbound_reply_sms_result", app.session_state)
            self.assertNotIn("inbound_reply_email_result", app.session_state)
            self.assertFalse(downloads.called)


if __name__ == "__main__":
    unittest.main()
