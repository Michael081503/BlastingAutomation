from __future__ import annotations

import unittest
from datetime import date
from io import BytesIO
from pathlib import Path
from tempfile import NamedTemporaryFile
from unittest.mock import patch

from openpyxl import load_workbook

from inbound_campaign_lookup import apply_phone_campaign_lookup, build_email_campaign_lookup
from inbound_email import build_email_unreplied_workbook
from inbound_reply import build_people_csv, build_unreplied_workbook, match_viber_sources, read_reply_source
from inbound_summary import campaign_totals_key
from test_inbound_campaign_lookup import workbook_bytes
from test_inbound_display import table_cells
from test_inbound_email import csv_bytes


ROOT = Path(__file__).resolve().parents[1]
REPORT_DATE = date(2026, 11, 18)


class Upload(BytesIO):
    def __init__(self, name, data):
        super().__init__(data)
        self.name = name


def all_channel_uploads(level="Level 1"):
    names = ("ana", "ben", "cara", "dan", "ella", "fran", "gina", "hugo", "iris", "unknown")
    phones = {name: f"09{17 + index}1234567" for index, name in enumerate(names)}
    regular = workbook_bytes([("OVERALL", [
        ("debtor_email", "LEVEL", "phone_1", "phone_2", "phone_3"),
        ("ana@example.com", level, None, phones["ana"], None),
        ("ben@example.com", level, phones["ben"], None, None),
        ("cara@example.com", level, None, None, phones["cara"]),
        ("dan@example.com", "Level 6", phones["dan"], None, None),
        ("ella@example.com", "Level 6", phones["ella"], None, None),
    ])])
    geo = workbook_bytes([("OVERALL", [
        ("email", "placement", "Phone 1", "Area"),
        ("fran@example.com", "Level 6", phones["fran"], "Region 3"),
        ("gina@example.com", "Level 1", phones["gina"], "R3"),
        ("hugo@example.com", "Level 6", phones["hugo"], "Mindanao"),
        ("iris@example.com", "Level 1", phones["iris"], "Mindanao"),
        (" HUGO@EXAMPLE.COM ", "Level 6", "+63" + phones["hugo"][1:], "Region 3"),
    ])])
    uploads = {"inbound_reply_email_lookup_uploads": [Upload("regular.xlsx", regular), Upload("geo.xlsx", geo)]}
    phone_headers = ("ID", "Client Name", "Phone", "Bank Campaign", "Message Content")
    customers = [(f"EXAMPLE-{index:03d}", name.title(), phones[name], "Original Campaign", "Customer reply")
                 for index, name in enumerate(names, 1)]
    # Repeated messages and equivalent Philippine phone formats count once.
    customers.append(("EXAMPLE-011", "Another name", "+63" + phones["ana"][1:], "Original Campaign", "Second message"))
    for channel, answered in (("sms", ("ana", "dan", "hugo")), ("viber", ("ana", "ben", "dan", "hugo"))):
        agents = [(f"AGENT-{index:03d}", name.title(), "+63" + phones[name][1:], "Original Campaign", "Agent response")
                  for index, name in enumerate(answered, 1)]
        agents.append(("AGENT-ONLY", "Outside", "09991234567", "Agent Only Campaign", "Outbound only"))
        uploads[f"inbound_reply_{channel}_1_upload"] = Upload("agents.csv", csv_bytes(phone_headers, agents))
        uploads[f"inbound_reply_{channel}_2_upload"] = Upload("customers.csv", csv_bytes(phone_headers, customers))
    uploads["inbound_reply_email_2_upload"] = Upload("dc.csv", csv_bytes(("ID", "sender", "Campaign", "Email"), [
        (f"EMAIL-{index:03d}", f"{name}@example.com", "Original Campaign", "Respond Now")
        for index, name in enumerate(names, 1)
    ] + [("EMAIL-DUPLICATE", " ANA@EXAMPLE.COM ", "Original Campaign", "Respond Now")]))
    uploads["inbound_reply_email_1_upload"] = Upload("history.csv", csv_bytes(("email", "sent", "receive"), [
        (f"{name}@example.com", "Agent response", None) for name in ("ana", "ben", "dan", "hugo", "agentonly")
    ] + [("fran@example.com", None, "Customer message")]))
    return uploads


def sample_manual_totals():
    return {campaign_totals_key(REPORT_DATE, "SBC Level 1"): {
        "active_ticket": 5000, "dc_email": 4337, "dc_sms": 3500, "dc_viber": 2500,
    }}


class PhonePopulationTests(unittest.TestCase):
    def test_only_customer_contacts_and_same_campaign_agent_responses_are_counted(self):
        headers = ("Phone", "Campaign")
        customers = read_reply_source("customers.csv", csv_bytes(headers, [
            ("09171234567", "A"), ("+63 917 123 4567", "a"),
            ("09181234567", "A"), ("09191234567", "B"),
        ]))
        agents = read_reply_source("agents.csv", csv_bytes(headers, [
            ("09171234567", "A"), ("09171234567", "A"),
            ("09181234567", "B"), ("09191234567", "A"), ("09201234567", "Agent only"),
        ]))
        report = match_viber_sources(customers, agents, agent_replied_source="campaign_replies", match_mode="phone")
        self.assertEqual([(c.campaign, c.unique_people, c.agent_replied, c.awaiting_reply) for c in report.campaigns],
                         [("A", 2, 1, 1), ("B", 1, 0, 1)])
        self.assertEqual((report.totals.unique_people, report.totals.agent_replied, report.totals.awaiting_reply), (3, 1, 2))
        self.assertNotIn("09201234567", build_people_csv(report).decode("utf-8-sig"))
        empty = read_reply_source("empty.csv", csv_bytes(headers, []))
        empty_report = match_viber_sources(empty, agents, agent_replied_source="campaign_replies", match_mode="phone")
        self.assertEqual(empty_report.totals.unique_people, 0)
        self.assertEqual(empty_report.campaigns, ())

    def test_phone_only_lookup_uses_all_columns_and_keeps_multiple_campaign_memberships(self):
        lookup = build_email_campaign_lookup([("regular.csv", csv_bytes(("LEVEL", "phone_1", "phone_2", "phone_3"), [
            ("Level 1", "09171234567", "+63 918 123 4567", "09191234567"),
            ("Level 6", "09171234567", None, None),
        ]))])
        self.assertEqual(lookup.campaigns_by_email, {})
        self.assertEqual(lookup.campaigns_by_phone, {
            "09171234567": ("SBC Level 1", "SBC Level 6"),
            "09181234567": ("SBC Level 1",), "09191234567": ("SBC Level 1",),
        })
        source = read_reply_source("customers.csv", csv_bytes(("Phone", "Name"), [
            ("09171234567", "Matched"), ("09201234567", "Outside"), ("+63 920 123 4567", "Duplicate"),
            ("invalid phone", "Missing usable phone"),
        ]),
                                   require_campaign=False)
        resolved = apply_phone_campaign_lookup(source, lookup)
        self.assertEqual([row.campaign for row in resolved.records], ["SBC Level 1", "SBC Level 6"])
        self.assertEqual(resolved.unmatched_addresses, 1)
        self.assertEqual(len(source.records), 4)
        self.assertEqual(resolved.excluded_rows, 1)
        self.assertEqual(source.excluded_rows, 0)
        self.assertEqual(resolved.headers, ("Phone", "Name", "Lookup Campaign"))
        self.assertEqual(source.headers, ("Phone", "Name"))


class AllChannelsPageTests(unittest.TestCase):
    def test_saved_databases_are_reused_after_clear_and_in_a_new_session(self):
        from streamlit.testing.v1 import AppTest
        from inbound_campaign_lookup import read_saved_campaign_lookup_files
        from upload_state import render_clear_all_files

        with NamedTemporaryFile(suffix=".zip") as temporary:
            saved_path = Path(temporary.name)
        self.addCleanup(saved_path.unlink, missing_ok=True)
        page = str(ROOT / "pages" / "INBOUND_REPLY.py")
        lookup_key = "inbound_reply_email_lookup_uploads"
        uploads = all_channel_uploads()
        original_files = [(upload.name, upload.getvalue()) for upload in uploads[lookup_key]]

        with patch("branding.render_bank_branding", side_effect=render_clear_all_files), \
                patch("inbound_reply_page.SAVED_CAMPAIGN_LOOKUP_PATH", saved_path), \
                patch("streamlit.file_uploader", side_effect=lambda label, **kw: uploads.get(kw["key"])):
            app = AppTest.from_file(page).run(timeout=20)
            app.button(key="inbound_reply_save_lookup").click().run(timeout=20)
            self.assertFalse(app.error)
            self.assertEqual(read_saved_campaign_lookup_files(saved_path), original_files)

            uploads.pop(lookup_key)
            fresh = AppTest.from_file(page).run(timeout=20)
            self.assertFalse(fresh.exception)
            for channel in ("email", "sms", "viber"):
                report = fresh.session_state[f"inbound_reply_{channel}_result"]["report"]
                self.assertEqual([entry.campaign for entry in report.campaigns],
                                 ["SBC Level 1", "SBC Level 6", "SBC Mindanao"])
            fresh.button(key="clear_all_files").click().run(timeout=20)
            self.assertNotIn("inbound_reply_email_result", fresh.session_state)
            self.assertEqual(read_saved_campaign_lookup_files(saved_path), original_files)
            self.assertTrue(any("Using saved campaign lookup files" in caption.value for caption in fresh.caption))

            uploads[lookup_key] = [Upload("regular.xlsx", b"broken")]
            app.run(timeout=20)
            self.assertTrue(app.error)
            self.assertFalse(any(button.key == "inbound_reply_save_lookup" for button in app.button))
            self.assertEqual(read_saved_campaign_lookup_files(saved_path), original_files)

            uploads[lookup_key] = all_channel_uploads("Level 7")[lookup_key]
            app.run(timeout=20)
            app.button(key="inbound_reply_save_lookup").click().run(timeout=20)
            self.assertFalse(app.error)
            uploads.pop(lookup_key)
            fresh.run(timeout=20)
            lookup = fresh.session_state["inbound_reply_email_lookup"]["lookup"]
            self.assertEqual(lookup.campaigns, ("SBC Level 6", "SBC Level 7", "SBC Mindanao"))

            fresh.button(key="inbound_reply_remove_saved_lookup").click().run(timeout=20)
            self.assertFalse(fresh.exception)
            self.assertFalse(saved_path.exists())
            self.assertNotIn("inbound_reply_email_lookup", fresh.session_state)

    def test_phone_reply_files_can_omit_campaign_in_csv_and_excel_when_lookup_is_uploaded(self):
        from streamlit.testing.v1 import AppTest

        for extension in ("csv", "xlsx"):
            with self.subTest(extension=extension):
                def data(phone):
                    rows = [(phone,)]
                    return (csv_bytes(("Phone",), rows) if extension == "csv" else
                            workbook_bytes([("Messages", [("Phone",), *rows])]))

                uploads = {"inbound_reply_email_lookup_uploads": all_channel_uploads()["inbound_reply_email_lookup_uploads"]}
                for channel in ("sms", "viber"):
                    uploads[f"inbound_reply_{channel}_1_upload"] = Upload(f"agents.{extension}", data("+63 917 123 4567"))
                    uploads[f"inbound_reply_{channel}_2_upload"] = Upload(f"customers.{extension}", data("09171234567"))
                with patch("branding.render_bank_branding"), \
                        patch("streamlit.file_uploader", side_effect=lambda label, **kw: uploads.get(kw["key"])):
                    app = AppTest.from_file(str(ROOT / "pages" / "INBOUND_REPLY.py")).run(timeout=20)
                    self.assertFalse(app.exception)
                    self.assertFalse(app.error)
                    for channel in ("sms", "viber"):
                        report = app.session_state[f"inbound_reply_{channel}_result"]["report"]
                        self.assertEqual([(c.campaign, c.unique_people, c.agent_replied) for c in report.campaigns],
                                         [("SBC Level 1", 1, 1)])

    def test_report_all_channels_lookup_replacements_errors_removal_and_clear(self):
        from streamlit.testing.v1 import AppTest
        from upload_state import render_clear_all_files

        uploads = all_channel_uploads()
        expected_labels = ["SBC Level 1", "SBC Level 6", "SBC Mindanao"]
        with patch("branding.render_bank_branding", side_effect=render_clear_all_files), \
                patch("inbound_reply_page.inbound_report_date", return_value=REPORT_DATE), \
                patch("streamlit.file_uploader", side_effect=lambda label, **kw: uploads.get(kw["key"])), \
                patch("streamlit.html") as html, patch("streamlit.download_button") as downloads:
            app = AppTest.from_file(str(ROOT / "pages" / "INBOUND_REPLY.py"))
            app.session_state["inbound_reply_campaign_totals"] = sample_manual_totals()
            app.run(timeout=20)

            def displayed():
                self.assertFalse(app.exception)
                call = next(call for call in reversed(html.call_args_list) if 'aria-label="Inbound reply summary"' in call.args[0])
                return table_cells(call.args[0])

            rows = displayed()
            self.assertFalse(app.error)
            self.assertEqual([row[0] for row in rows], expected_labels)
            self.assertEqual(rows[0][1:11], ["5,000", "4,337", "86.74%", "3,500", "70.00%", "2,500", "50.00%", "3", "3", "3"])
            self.assertEqual(rows[0][12:], ["2", "66.67%", "1", "33.33%", "2", "66.67%"])
            self.assertEqual(rows[2][8:11], ["4", "4", "4"])
            self.assertEqual(rows[2][12:], ["1", "25.00%"] * 3)
            call = next(call for call in reversed(downloads.call_args_list) if call.kwargs.get("key") == "inbound_reply_summary_download")
            book = load_workbook(BytesIO(call.kwargs["data"]()), data_only=True)
            try:
                self.assertEqual([book["Summary"].cell(r, 1).value for r in range(4, 4 + len(expected_labels))], expected_labels)
                self.assertEqual(book["Summary"].max_row, 4 + len(expected_labels))
                for cell in ("I4", "J4", "K4"):
                    self.assertEqual(book["Summary"][cell].value, 3)
                self.assertEqual([book["Summary"].cell(6, col).value for col in (9, 10, 11, 13, 15, 17)],
                                 [4, 4, 4, 1, 1, 1])
                for cell, expected in (("M4", 2), ("N4", 2 / 3), ("O4", 1), ("P4", 1 / 3), ("Q4", 2), ("R4", 2 / 3)):
                    self.assertAlmostEqual(book["Summary"][cell].value, expected)
                for title, count in (("SMS Unreplied", 6), ("Viber Unreplied", 5), ("Email Unreplied", 5)):
                    sheet = book[title]
                    self.assertEqual(sheet["B1"].value, "Processed")
                    self.assertEqual(sheet["B2"].value, count)
                    self.assertEqual(sheet.max_row - 3, count)
                    self.assertIn("Source: Lookup Campaign", [cell.value for cell in sheet[3]])
                    self.assertTrue(all(sheet.cell(row, 2).value in expected_labels for row in range(4, sheet.max_row + 1)))
                for channel in ("email", "sms", "viber"):
                    report = app.session_state[f"inbound_reply_{channel}_result"]["report"]
                    self.assertEqual(report.totals.unique_people, 9)
            finally:
                book.close()
            call = next(call for call in reversed(downloads.call_args_list)
                        if call.kwargs.get("key") == "inbound_reply_unmatched_download")
            book = load_workbook(BytesIO(call.kwargs["data"]()))
            try:
                self.assertEqual(book.sheetnames, ["SMS Unmatched", "Viber Unmatched", "Email Unmatched"])
                for title in book.sheetnames:
                    headers, *unmatched_rows = book[title].values
                    self.assertEqual(len(unmatched_rows), 1)
                    self.assertEqual(unmatched_rows[0][headers.index("ID")], "EMAIL-010" if title == "Email Unmatched" else "EXAMPLE-010")
                self.assertEqual(book["Email Unmatched"]["A2"].value, "unknown@example.com")
                self.assertEqual(book["SMS Unmatched"]["A2"].value, "09261234567")
                self.assertEqual(book["SMS Unmatched"]["A2"].number_format, "@")
            finally:
                book.close()
            uploads["inbound_reply_email_lookup_uploads"] = all_channel_uploads("Level 7")["inbound_reply_email_lookup_uploads"]
            app.run(timeout=20)
            self.assertEqual([row[0] for row in displayed()], ["SBC Level 6", "SBC Level 7", *expected_labels[2:]])
            self.assertTrue(all(row[1] == "—" for row in displayed()))
            uploads["inbound_reply_email_lookup_uploads"] = uploads["inbound_reply_email_lookup_uploads"][:1]
            app.run(timeout=20)
            self.assertEqual([row[0] for row in displayed()], ["SBC Level 6", "SBC Level 7"])
            for channel in ("email", "sms", "viber"):
                report = app.session_state[f"inbound_reply_{channel}_result"]["report"]
                source = report.dc if channel == "email" else report.message_logs
                self.assertEqual(source.unmatched_addresses, 5)
                self.assertEqual(report.totals.unique_people, 5)
            # Replacing the lookup with no matching contacts must remove every
            # previously prepared customer row from every channel and download.
            uploads["inbound_reply_email_lookup_uploads"] = [Upload("other.csv", csv_bytes(
                ("email", "LEVEL", "phone_1"), [("outside@example.com", "Level 1", "09981234567")],
            ))]
            app.run(timeout=20)
            self.assertEqual([row[0] for row in displayed()], ["SBC Level 1"])
            for channel in ("email", "sms", "viber"):
                report = app.session_state[f"inbound_reply_{channel}_result"]["report"]
                source = report.dc if channel == "email" else report.message_logs
                self.assertEqual(source.records, ())
                self.assertEqual(source.unmatched_addresses, 10)
                self.assertEqual(report.totals.unique_people, 0)
                self.assertEqual(report.totals.agent_replied, 0)
                self.assertEqual(report.totals.awaiting_reply, 0)
                data = (build_email_unreplied_workbook(report) if channel == "email" else
                        build_unreplied_workbook(report))
                book = load_workbook(BytesIO(data))
                try:
                    self.assertEqual(book["Unreplied Details"].max_row, 1)
                    self.assertEqual(book["Summary"].cell(2, 2).value, 0)
                finally:
                    book.close()
            call = next(call for call in reversed(downloads.call_args_list)
                        if call.kwargs.get("key") == "inbound_reply_summary_download")
            book = load_workbook(BytesIO(call.kwargs["data"]()))
            try:
                self.assertEqual(book["Summary"].max_row, 5)
                self.assertEqual(book["Summary"]["A4"].value, "SBC Level 1")
                for title in ("SMS Unreplied", "Viber Unreplied", "Email Unreplied"):
                    self.assertEqual(book[title]["B1"].value, "Processed")
                    self.assertEqual(book[title]["B2"].value, 0)
                    self.assertEqual(book[title].max_row, 3)
            finally:
                book.close()
            uploads["inbound_reply_email_lookup_uploads"] = [Upload("bad.xlsx", b"broken")]
            app.run(timeout=20)
            self.assertFalse(app.exception)
            self.assertTrue(app.error)
            self.assertEqual(len(app.metric), 0)
            for channel in ("email", "sms", "viber"):
                self.assertNotIn(f"inbound_reply_{channel}_result", app.session_state)
            uploads["inbound_reply_email_lookup_uploads"] = []
            app.run(timeout=20)
            self.assertFalse(app.error)
            self.assertEqual([row[0] for row in displayed()], ["Original Campaign"])
            self.assertEqual(displayed()[0][8:11], ["10", "10", "10"])
            app.button(key="clear_all_files").click().run(timeout=20)
            self.assertFalse(app.exception)
            self.assertEqual(len(app.metric), 0)
            self.assertEqual([row[0] for row in displayed()], ["—"])

    def test_unmatched_download_uses_pending_customer_rows_and_refreshes_with_lookup(self):
        from streamlit.testing.v1 import AppTest

        lookup_key = "inbound_reply_email_lookup_uploads"
        download_key = "inbound_reply_unmatched_download"
        lookup_data = csv_bytes(("Email", "LEVEL"), [("known@example.com", 6)])
        uploads = {
            lookup_key: [Upload("regular.csv", lookup_data)],
            "inbound_reply_email_2_upload": Upload("dc.csv", csv_bytes(("ID", "sender", "Campaign", "body"), [
                ("0001", " NEW@example.com ", "Original Campaign", "=Original message"),
                ("0002", "new@example.com", "Original Campaign", "Duplicate"),
                ("0003", "loan@example.com", "PL RECOVERY", "PL Recovery"),
                ("0004", "ignore@example.com", "SBF", "SBF exclusion"),
                ("0005", "invalid address", "Original Campaign", "Invalid email"),
                ("0006", "known@example.com", "Original Campaign", "Matched"),
            ])),
        }
        with patch("branding.render_bank_branding"), \
                patch("streamlit.file_uploader", side_effect=lambda label, **kw: uploads.get(kw["key"])), \
                patch("streamlit.download_button") as downloads:
            app = AppTest.from_file(str(ROOT / "pages" / "INBOUND_REPLY.py")).run(timeout=20)
            self.assertFalse(app.exception)
            self.assertFalse(app.error)
            self.assertEqual([metric.value for metric in app.metric], ["2", "—", "—"])
            call = next(call for call in downloads.call_args_list if call.kwargs.get("key") == download_key)
            book = load_workbook(BytesIO(call.kwargs["data"]()))
            try:
                self.assertEqual(book.sheetnames, ["Email Unmatched"])
                sheet = book.active
                self.assertEqual(list(sheet.values), [
                    ("Normalized Email", "ID", "sender", "Campaign", "body"),
                    ("new@example.com", "0001", " NEW@example.com ", "Original Campaign", "=Original message"),
                ])
                self.assertEqual(sheet["E2"].data_type, "s")
            finally:
                book.close()

            uploads[lookup_key] = [Upload("regular.csv", csv_bytes(("Email", "LEVEL"), [
                ("known@example.com", 6), ("new@example.com", 6),
            ]))]
            downloads.reset_mock()
            app.run(timeout=20)
            self.assertFalse(app.exception)
            self.assertFalse(any(call.kwargs.get("key") == download_key for call in downloads.call_args_list))

            uploads[lookup_key] = [Upload("invalid.xlsx", b"invalid")]
            downloads.reset_mock()
            app.run(timeout=20)
            self.assertTrue(app.error)
            self.assertFalse(any(call.kwargs.get("key") == download_key for call in downloads.call_args_list))

    def test_viber_accepts_identical_fully_replied_files(self):
        from streamlit.testing.v1 import AppTest

        data = csv_bytes(("Phone", "Campaign"), [("09171234567", "A")])
        uploads = {f"inbound_reply_viber_{number}_upload": Upload("replied.csv", data) for number in (1, 2)}
        with patch("branding.render_bank_branding"), \
                patch("streamlit.file_uploader", side_effect=lambda label, **kw: uploads.get(kw["key"])):
            app = AppTest.from_file(str(ROOT / "pages" / "INBOUND_REPLY.py")).run(timeout=20)
            self.assertFalse(app.exception)
            self.assertFalse(app.error)
            self.assertEqual([metric.value for metric in app.metric], ["1", "1", "0", "1"])

    def test_agent_only_uploads_do_not_populate_customer_campaign_rows(self):
        from streamlit.testing.v1 import AppTest

        uploads = {key: value for key, value in all_channel_uploads().items() if key.endswith("_1_upload")}
        with patch("branding.render_bank_branding"), \
                patch("streamlit.file_uploader", side_effect=lambda label, **kw: uploads.get(kw["key"])), \
                patch("streamlit.html") as html:
            app = AppTest.from_file(str(ROOT / "pages" / "INBOUND_REPLY.py")).run(timeout=20)
            self.assertFalse(app.exception)
            call = next(call for call in reversed(html.call_args_list) if 'aria-label="Inbound reply summary"' in call.args[0])
            self.assertEqual([row[0] for row in table_cells(call.args[0])], ["—"])
            self.assertEqual(len(app.metric), 0)


if __name__ == "__main__":
    unittest.main()
