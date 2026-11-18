from __future__ import annotations

import unittest
from datetime import date
from io import BytesIO
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook

from inbound_campaign_lookup import apply_phone_campaign_lookup, build_email_campaign_lookup
from inbound_reply import build_people_csv, match_viber_sources, read_reply_source
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
        ("dan@example.com", "Level 2", phones["dan"], None, None),
        ("ella@example.com", "Level 2", phones["ella"], None, None),
    ])])
    geo = workbook_bytes([("OVERALL", [
        ("email", "placement", "Phone 1", "Area"),
        ("fran@example.com", "Level 6", phones["fran"], "Region 3"),
        ("gina@example.com", "Level 1", phones["gina"], "Region 3"),
        ("hugo@example.com", "Level 6", phones["hugo"], "Mindanao"),
        ("iris@example.com", "Level 1", phones["iris"], "Mindanao"),
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
            ("Level 2", "09171234567", None, None),
        ]))])
        self.assertEqual(lookup.campaigns_by_email, {})
        self.assertEqual(lookup.campaigns_by_phone, {
            "09171234567": ("SBC Level 1", "SBC Level 2"),
            "09181234567": ("SBC Level 1",), "09191234567": ("SBC Level 1",),
        })
        source = read_reply_source("customers.csv", csv_bytes(("Phone",), [("09171234567",), ("09201234567",)]),
                                   require_campaign=False)
        resolved = apply_phone_campaign_lookup(source, lookup)
        self.assertEqual([row.campaign for row in resolved.records], ["SBC Level 1", "SBC Level 2", "Unmatched campaign"])
        self.assertEqual(resolved.unmatched_addresses, 1)
        self.assertEqual(resolved.headers, ("Phone", "Lookup Campaign"))
        self.assertEqual(source.headers, ("Phone",))


class AllChannelsPageTests(unittest.TestCase):
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
        expected_labels = ["SBC Level 1", "SBC Level 2", "SBC Mindanao", "SBC Region 3", "Unmatched campaign"]
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
            self.assertEqual(rows[3][12:], ["0", "0.00%", "0", "0.00%", "0", "0.00%"])
            call = next(call for call in reversed(downloads.call_args_list) if call.kwargs.get("key") == "inbound_reply_summary_download")
            book = load_workbook(BytesIO(call.kwargs["data"]()))
            try:
                self.assertEqual([book["Summary"].cell(r, 1).value for r in range(4, 9)], expected_labels)
                for cell in ("I4", "J4", "K4"):
                    self.assertEqual(book["Summary"][cell].value, 3)
                for cell, expected in (("M4", 2), ("N4", 2 / 3), ("O4", 1), ("P4", 1 / 3), ("Q4", 2), ("R4", 2 / 3)):
                    self.assertAlmostEqual(book["Summary"][cell].value, expected)
                for title, count in (("SMS Unreplied", 7), ("Viber Unreplied", 6), ("Email Unreplied", 6)):
                    sheet = book[title]
                    self.assertEqual(sheet["B1"].value, "Processed")
                    self.assertEqual(sheet["B2"].value, count)
                    self.assertEqual(sheet.max_row - 3, count)
                    self.assertIn("Source: Lookup Campaign", [cell.value for cell in sheet[3]])
            finally:
                book.close()
            uploads["inbound_reply_email_lookup_uploads"] = all_channel_uploads("Level 4")["inbound_reply_email_lookup_uploads"]
            app.run(timeout=20)
            self.assertEqual([row[0] for row in displayed()], ["SBC Level 2", "SBC Level 4", *expected_labels[2:]])
            self.assertTrue(all(row[1] == "—" for row in displayed()))
            uploads["inbound_reply_email_lookup_uploads"] = uploads["inbound_reply_email_lookup_uploads"][:1]
            app.run(timeout=20)
            self.assertEqual([row[0] for row in displayed()], ["SBC Level 2", "SBC Level 4", "Unmatched campaign"])
            for channel in ("email", "sms", "viber"):
                report = app.session_state[f"inbound_reply_{channel}_result"]["report"]
                source = report.dc if channel == "email" else report.message_logs
                self.assertEqual(source.unmatched_addresses, 5)
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
