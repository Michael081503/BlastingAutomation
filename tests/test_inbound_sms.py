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
    def test_client_and_agent_touch_workbooks_keep_one_client_row_per_account(self):
        from openpyxl import load_workbook
        from streamlit.testing.v1 import AppTest
        from test_inbound_campaign_lookup import workbook_bytes

        class Upload(BytesIO):
            def __init__(self, name, data):
                super().__init__(data)
                self.name = name

        client_headers = ("Received Date / Time", "Debtor Name", "Client", "Account No.", "Phone Number", "Message")
        agent_headers = ("Call Type", "Customer", "Client", "Account", "Contact Number")
        first_message = "First message\n" * 10
        clients = Upload("CLIENT.xlsx", workbook_bytes([
            ("Comparison", [("Client", "Account No.", "Phone Number"), ("Sample", "0999", "09991234567")]),
            ("Incoming", [client_headers,
                          ("11/24", "One", "SBC CARDS B2", "0001", "09171234567", first_message),
                          ("11/24", "One", "SBC CARDS L1", "0001", "09181234567", "Repeated account"),
                          ("11/24", "Two", "SBC CARDS L1", "0002", "+63 9171234567", "Shared phone"),
                          ("11/24", "Three", "SBC CARDS L1", "0003", "09191234567", "Third account")]),
        ]))
        first = Upload("AGENT TOUCH RECO.xlsx", workbook_bytes([("Calls", [agent_headers,
            ("Outgoing Call", "One", "SBC CARDS B2", "0001", "09341234567"),
            ("Outgoing Call", "One", "SBC CARDS B2", "0001", "09341234567"),
            ("Outgoing Call", "Outside", "SBC CARDS B2", "0004", "09991234567"),
        ])]))
        second = Upload("AGENT TOUCH INSURANCE.xlsx", workbook_bytes([("Calls", [agent_headers,
            ("Outgoing Call", "Three", "SBC INSURANCE", "0003", "09191234567"),
        ])]))
        uploads = {"inbound_reply_sms_1_upload": [first, second], "inbound_reply_sms_2_upload": clients}
        with patch("branding.render_bank_branding"), \
                patch("streamlit.file_uploader", side_effect=lambda label, **kwargs: uploads.get(kwargs["key"])), \
                patch("streamlit.download_button") as downloads:
            app = AppTest.from_file(str(ROOT / "pages" / "INBOUND_REPLY.py")).run(timeout=20)
            self.assertFalse(app.exception)
            self.assertFalse(app.error)
            self.assertEqual([metric.value for metric in app.metric], ["3", "2", "1", "3"])
            self.assertTrue(any("Removed 1 duplicate Client" in caption.value for caption in app.caption))
            call = next(call for call in downloads.call_args_list
                        if call.kwargs.get("key") == "inbound_reply_client_accounts_download")
            book = load_workbook(BytesIO(call.kwargs["data"]()))
            try:
                headers, *rows = book["Client Accounts"].values
                self.assertEqual([(row[headers.index("Account No.")], row[1], row[-1]) for row in rows], [
                    ("0001", "Touched", first_message), ("0002", "No touch", "Shared phone"),
                    ("0003", "Touched", "Third account"),
                ])
                self.assertEqual(book["Client Accounts"]["H2"].hyperlink.location, "'Client Accounts Text'!A2")
            finally:
                book.close()

            uploads["inbound_reply_email_lookup_uploads"] = [
                Upload("Regular.csv", b"Phone,LEVEL\n09171234567,1\n09191234567,1\n"),
                Upload("B2.csv", b"Phone\n09171234567\n"),
            ]
            app.run(timeout=20)
            self.assertFalse(app.error)
            report = app.session_state["inbound_reply_sms_result"]["report"]
            self.assertEqual(report.message_logs.sheet, "Incoming")
            self.assertEqual([(row.account, row.campaign) for row, _ in report.client_touches], [
                ("0001", "SBC B2"), ("0002", "SBC Level 1"), ("0003", "SBC Level 1"),
            ])
            self.assertEqual((report.totals.unique_people, report.totals.agent_replied), (3, 2))

    def test_sms_only_upload_accepts_a_fully_replied_account_list(self):
        from openpyxl import load_workbook
        from streamlit.testing.v1 import AppTest

        class Upload(BytesIO):
            def __init__(self, name, data):
                super().__init__(data)
                self.name = name

        data = csv_bytes(PHONE_HEADERS, [("0001", "One", "09171234567", "A", "Message")])
        uploads = {"inbound_reply_sms_1_upload": Upload("replied.csv", data),
                   "inbound_reply_sms_2_upload": Upload("accounts.csv", data)}
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
        book = load_workbook(BytesIO(build_summary_workbook("", date(2026, 11, 18), "YES", viber, email, sms_report=sms)),
                             data_only=True)
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
            "inbound_reply_sms_2_upload": Upload("accounts.csv", accounts),
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
            uploads["inbound_reply_sms_1_upload"] = Upload("replied.csv", replies)
            app.run(timeout=20)
            book = output_book()
            self.assertEqual(book["SMS Unreplied"]["B2"].value, 2)
            self.assertEqual(len(detail_rows(book["SMS Unreplied"])), 2)
            self.assertEqual({row["Source: ID"] for row in detail_rows(book["SMS Unreplied"])}, {"000008", "000010"})
            book.close()
            self.assertIn("report", app.session_state["inbound_reply_sms_result"])
            uploads["inbound_reply_sms_1_upload"] = Upload("replied.csv", b"Wrong,Headers\n1,2\n")
            app.run(timeout=20)
            self.assertTrue(app.error)
            book = output_book()
            self.assertIn("Could not process", book["SMS Unreplied"]["B1"].value)
            self.assertIsNone(book["SMS Unreplied"]["B2"].value)
            self.assertEqual(detail_rows(book["SMS Unreplied"]), [])
            self.assertEqual(book["Email Unreplied"]["B2"].value, 1)
            book.close()
            self.assertNotIn("report", app.session_state["inbound_reply_sms_result"])
            del uploads["inbound_reply_sms_2_upload"]
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
