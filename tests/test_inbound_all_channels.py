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
from inbound_summary import build_summary_workbook, campaign_totals_key
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
    customers = [(f"EXAMPLE-{index:03d}", name.title(), phones[name], "", "Customer reply")
                 for index, name in enumerate(names, 1)]
    # Repeated messages and equivalent Philippine phone formats count once.
    customers.append(("EXAMPLE-011", "Another name", "+63" + phones["ana"][1:], "", "Second message"))
    for channel, answered in (("sms", ("ana", "dan", "hugo")), ("viber", ("ana", "ben", "dan", "hugo"))):
        agents = [(f"AGENT-{index:03d}", name.title(), "+63" + phones[name][1:], "", "Agent response")
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
    def test_lookup_preserves_labeled_phone_memberships_and_fills_blank_campaigns(self):
        headers = ("Phone", "Campaign")
        source = read_reply_source("customers.csv", csv_bytes(headers, [
            ("09171234567", "SBC Cards B2"), ("+63 917 123 4567", "SBC Cards B4"),
            ("09181234567", "SBC Insurance"), ("09191234567", ""),
        ]))
        lookup = build_email_campaign_lookup([("regular.csv", b"Phone,LEVEL\n09171234567,1\n09191234567,6\n")])
        resolved = apply_phone_campaign_lookup(source, lookup)
        self.assertEqual([row.campaign for row in resolved.records],
                         ["SBC Cards B2", "SBC Cards B4", "SBC Insurance", "SBC Level 6"])
        self.assertEqual(resolved.unmatched_addresses, 0)
        agents = read_reply_source("agents.csv", csv_bytes(headers, [("09171234567", "SBC Cards B4")]))
        report = match_viber_sources(resolved, apply_phone_campaign_lookup(agents, lookup),
                                     agent_replied_source="campaign_replies", match_mode="phone")
        self.assertEqual([(row.campaign, row.unique_people, row.agent_replied) for row in report.campaigns], [
            ("SBC Cards B2", 1, 0), ("SBC Cards B4", 1, 1), ("SBC Insurance", 1, 0), ("SBC Level 6", 1, 0),
        ])

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
        self.assertEqual([row.campaign for row in resolved.records], ["SBC Level 1", "SBC Level 6", "sblegal", "sblegal"])
        self.assertEqual(resolved.unmatched_addresses, 1)
        self.assertEqual(len(source.records), 4)
        self.assertEqual(resolved.excluded_rows, 1)
        self.assertEqual(source.excluded_rows, 0)
        self.assertEqual(resolved.headers, ("Phone", "Name", "Lookup Campaign"))
        self.assertEqual(source.headers, ("Phone", "Name"))


class AllChannelsPageTests(unittest.TestCase):
    def test_agent_worked_and_same_day_unreplied_counts_on_page_and_in_excel(self):
        from streamlit.testing.v1 import AppTest

        phone_headers = ("Phone", "Campaign", "Timestamp")
        customers = [("09171234567", "A", "23/11/2026, 8:00am"),
                     ("09171234567", "A", "24/11/2026, 8:00am"),
                     ("09181234567", "A", "24/11/2026, 8:00am"),
                     ("09191234567", "A", "24/11/2026, 8:00am")]
        agents = [("+63 917 123 4567", "A", "2026-11-24T10:00:00Z"),
                  ("09181234567", "A", "2026-11-24T16:01:00Z")]
        uploads = {}
        for channel in ("sms", "viber"):
            uploads[f"inbound_reply_{channel}_1_upload"] = Upload("agents.csv", csv_bytes(phone_headers, agents))
            uploads[f"inbound_reply_{channel}_2_upload"] = Upload("customers.csv", csv_bytes(phone_headers, customers))
        uploads["inbound_reply_email_2_upload"] = Upload("dc.xlsx", workbook_bytes([("Checking", [
            ("sender", "Campaign", "Date", "Date_created", "date"),
            ("ana@example.com", "A", date(2026, 11, 24), date(2026, 11, 24), "Tuesday"),
            ("ben@example.com", "A", date(2026, 11, 24), date(2026, 11, 24), "Tuesday"),
            ("cara@example.com", "A", date(2026, 11, 24), date(2026, 11, 24), "Tuesday"),
            ("ana@example.com", "B", date(2026, 11, 23), date(2026, 11, 23), "Monday"),
        ])]))
        uploads["inbound_reply_email_1_upload"] = Upload("history.csv", csv_bytes(("email", "sent", "Date"), [
            ("ana@example.com", "Agent reply", "2026-11-24 10:00"),
            ("ben@example.com", "Later reply", "2026-11-25 08:00"),
            ("cara@example.com", "", "2026-11-24 15:00"),
        ]))
        with patch("branding.render_bank_branding"), \
                patch("inbound_reply_page.SAVED_CAMPAIGN_LOOKUP_PATH", ROOT / "tests" / "missing-lookup.zip"), \
                patch("streamlit.file_uploader", side_effect=lambda label, **kw: uploads.get(kw["key"])), \
                patch("streamlit.html") as html, \
                patch("streamlit.download_button") as downloads:
            app = AppTest.from_file(str(ROOT / "pages" / "INBOUND_REPLY.py")).run(timeout=20)
            self.assertFalse(app.exception)
            self.assertFalse(app.error)
            summary = next(call.args[0] for call in html.call_args_list if "AGENT WORKED" in call.args[0])
            self.assertEqual(table_cells(summary)[0][12:], ["1", "33.33%", "2", "66.67%", "2", "66.67%"])
            call = next(call for call in downloads.call_args_list
                        if call.kwargs.get("key") == "inbound_reply_summary_download")
            with BytesIO(call.kwargs["data"]()) as data:
                book = load_workbook(data, data_only=True)
                try:
                    self.assertEqual([book["Summary"][f"{column}4"].value for column in "MOQ"], [1, 2, 2])
                    self.assertEqual([book[name]["B2"].value for name in
                                      ("SMS Unreplied", "Viber Unreplied", "Email Unreplied")], [2, 2, 3])
                    self.assertEqual(book["Summary"]["M5"].value, 0)
                finally:
                    book.close()

    def test_geo_and_reco_overall_supply_active_tickets_in_excel_summary(self):
        from streamlit.testing.v1 import AppTest

        geo_headers = ("Account", "Area", "email")
        geo = workbook_bytes([
            ("OVERALL", [geo_headers,
                         ("001", "Region 3", "one@example.com"),
                         ("001", "Mindanao", "alternate@example.com"),
                         ("002", "Mindanao", None),
                         (None, "Mindanao", "no-account@example.com")]),
            ("NEGATIVE", [geo_headers, ("003", "Mindanao", "other-sheet@example.com")]),
        ])
        reco_headers = ("Account No.", "LEVEL", "Product Type", "debtor_email")
        reco = workbook_bytes([
            ("OVERALL", [reco_headers,
                         ("004", 1, "CARDS", "card-one@example.com"),
                         ("004", "Level 2", " cards ", "alternate-card@example.com"),
                         ("005", "Level 5", "CARDS", None),
                         ("006", "L6", "CARDS", "card-six@example.com"),
                         ("007", 6, "CSL", "loan@example.com"),
                         ("008", None, "BEL", None)]),
            ("NEGATIVE", [reco_headers, ("009", 6, "CARDS", "outside-overall@example.com")]),
        ])
        lookup_key = "inbound_reply_email_lookup_uploads"
        uploads = {lookup_key: [Upload("GEO.xlsx", geo), Upload("RECO.xlsx", reco)]}
        with patch("branding.render_bank_branding"), \
                patch("inbound_reply_page.inbound_report_date", return_value=REPORT_DATE), \
                patch("inbound_reply_page.SAVED_CAMPAIGN_LOOKUP_PATH", ROOT / "tests" / "missing-lookup.zip"), \
                patch("streamlit.file_uploader", side_effect=lambda label, **kw: uploads.get(kw["key"])), \
                patch("streamlit.download_button") as downloads:
            app = AppTest.from_file(str(ROOT / "pages" / "INBOUND_REPLY.py"))
            app.session_state["inbound_reply_campaign_totals"] = {
                campaign_totals_key(REPORT_DATE, "SBC Region 3"): {"active_ticket": 100, "dc_email": 1},
            }
            app.run(timeout=20)
            self.assertFalse(app.exception)
            self.assertFalse(app.error)
            call = next(call for call in downloads.call_args_list
                        if call.kwargs.get("key") == "inbound_reply_summary_download")
            book = load_workbook(BytesIO(call.kwargs["data"]()), data_only=True)
            try:
                sheet = book["Summary"]
                rows = {campaign: (ticket, blasts, coverage) for campaign, ticket, blasts, coverage in
                        sheet.iter_rows(min_row=4, max_row=sheet.max_row - 1, max_col=4, values_only=True)}
                self.assertEqual(rows, {
                    "PL RECOVERY": (2, None, None),
                    "SBC Level 1": (2, None, None),
                    "SBC Level 6": (1, None, None),
                    "SBC Mindanao": (2, 1, 0.5),
                })
            finally:
                book.close()

    def test_summary_workbook_totals_survive_database_upload(self):
        from streamlit.testing.v1 import AppTest

        report_date = date(2026, 11, 24)
        labels = ("SBC Insurance", "SBC Mindanao")
        totals = {campaign_totals_key(report_date, label): {"active_ticket": 100, "dc_email": 10,
                                                          "dc_sms": 0, "dc_viber": 20} for label in labels}
        reference = build_summary_workbook("", report_date, "—", campaign_totals=totals, source_campaigns=labels)
        replies = csv_bytes(("Phone", "Campaign"), [("09171234567", "SBC Mindanao")])
        uploads = {
            "inbound_reply_totals_upload": Upload("report.xlsx", reference),
            "inbound_reply_viber_1_upload": Upload("agent.csv", replies),
            "inbound_reply_viber_2_upload": Upload("customer.csv", replies),
        }
        geo = workbook_bytes([("OVERALL", [("Account", "Area", "Phone"), ("001", "Mindanao", "09171234567")])])
        reco = workbook_bytes([("OVERALL", [("Account", "LEVEL", "Product Type", "Phone"), ("002", 6, "CARDS", None)])])

        with patch("branding.render_bank_branding"), \
                patch("inbound_reply_page.inbound_report_date", return_value=REPORT_DATE), \
                patch("inbound_reply_page.SAVED_CAMPAIGN_LOOKUP_PATH", ROOT / "tests" / "missing-lookup.zip"), \
                patch("streamlit.file_uploader", side_effect=lambda label, **kw: uploads.get(kw["key"])), \
                patch("streamlit.download_button") as downloads:
            app = AppTest.from_file(str(ROOT / "pages" / "INBOUND_REPLY.py"))
            for database in ([], [Upload("GEO.xlsx", geo), Upload("RECO.xlsx", reco)]):
                uploads["inbound_reply_email_lookup_uploads"] = database
                downloads.reset_mock()
                app.run(timeout=20)
                self.assertFalse(app.exception)
                self.assertFalse(app.error, [error.value for error in app.error])
                call = next(call for call in downloads.call_args_list
                            if call.kwargs.get("key") == "inbound_reply_summary_download")
                book = load_workbook(BytesIO(call.kwargs["data"]()), data_only=True)
                try:
                    sheet = book["Summary"]
                    self.assertEqual(sheet["A1"].value, "NOVEMBER 24, 2026 DATA")
                    self.assertEqual(tuple(sheet[f"A{row}"].value for row in (4, 5)), labels)
                    self.assertEqual(sheet.max_row, 6)
                    self.assertEqual(tuple(sheet[f"B{row}"].value for row in (4, 5, 6)), (100, 100, 200))
                    self.assertEqual(tuple(sheet[f"K{row}"].value for row in (4, 5, 6)), (0, 1, 1))
                    self.assertEqual(tuple(sheet[f"R{row}"].value for row in (4, 5, 6)), (0, 1, 1))
                    self.assertEqual(sheet["H6"].value, .2)
                finally:
                    book.close()

    def test_sbf_only_email_upload_keeps_existing_campaigns_and_totals(self):
        from streamlit.testing.v1 import AppTest

        phone_file = csv_bytes(("Phone", "Campaign"), [("09171234567", "SBC Level 1")])
        uploads = {
            "inbound_reply_sms_1_upload": [Upload("sms-agent.csv", phone_file)],
            "inbound_reply_sms_2_upload": Upload("sms-client.csv", phone_file),
            "inbound_reply_viber_1_upload": Upload("viber-agent.csv", phone_file),
            "inbound_reply_viber_2_upload": Upload("viber-client.csv", phone_file),
            "inbound_reply_email_lookup_uploads": [Upload("Regular.csv", csv_bytes(
                ("email", "LEVEL"), [("one@example.com", 1), ("outside@example.com", 6)],
            ))],
        }

        with patch("branding.render_bank_branding"), \
                patch("inbound_reply_page.inbound_report_date", return_value=REPORT_DATE), \
                patch("inbound_reply_page.SAVED_CAMPAIGN_LOOKUP_PATH", ROOT / "tests" / "missing-lookup.zip"), \
                patch("streamlit.file_uploader", side_effect=lambda label, **kw: uploads.get(kw["key"])), \
                patch("streamlit.html") as html, patch("streamlit.download_button") as downloads:
            app = AppTest.from_file(str(ROOT / "pages" / "INBOUND_REPLY.py"))
            app.session_state["inbound_reply_campaign_totals"] = {
                campaign_totals_key(REPORT_DATE, "SBC Level 1"): {
                    "active_ticket": 100, "dc_email": 20, "dc_sms": 30, "dc_viber": 40,
                },
            }

            def snapshot():
                html.reset_mock()
                downloads.reset_mock()
                app.run(timeout=20)
                self.assertFalse(app.exception)
                self.assertFalse(app.error)
                markup = next(call.args[0] for call in html.call_args_list
                              if 'aria-label="Inbound reply summary"' in call.args[0])
                call = next(call for call in downloads.call_args_list
                            if call.kwargs.get("key") == "inbound_reply_summary_download")
                book = load_workbook(BytesIO(call.kwargs["data"]()), read_only=True, data_only=True)
                try:
                    return table_cells(markup), list(book["Summary"].values), book["Email Unreplied"]["B1"].value
                finally:
                    book.close()

            baseline_page, baseline_sheet, _ = snapshot()
            self.assertEqual([row[0] for row in baseline_page], ["SBC Level 1"])
            uploads["inbound_reply_email_1_upload"] = Upload(
                "EMAIL AGENT.xlsx", workbook_bytes([("email_history", [
                    ("email", "sent"), ("one@example.com", "Agent reply"),
                ])]),
            )
            uploads["inbound_reply_email_2_upload"] = Upload(
                "EMAIL CLIENT.xlsx", workbook_bytes([("mc4_emails", [
                    ("sender", "Campaign"), ("one@example.com", "MC4 SBF SALAD EMAIL INBOUND"),
                    ("two@example.com", "MC4 SBF SALAD EMAIL INBOUND"),
                ])]),
            )
            page_rows, sheet_rows, status = snapshot()
            self.assertEqual(page_rows, baseline_page)
            self.assertEqual(sheet_rows, baseline_sheet)
            self.assertEqual(status, "No eligible Email replies (2 SBF rows excluded)")
            self.assertTrue(any("The SBF filter excluded 2 rows." in warning.value for warning in app.warning))

            uploads["inbound_reply_email_2_upload"] = Upload(
                "EMAIL CLIENT.xlsx", workbook_bytes([("mc4_emails", [
                    ("sender", "Campaign"), ("one@example.com", "SBC Level 1"),
                ])]),
            )
            page_rows, sheet_rows, status = snapshot()
            self.assertEqual([row[0] for row in page_rows], ["SBC Level 1", "SBC Level 6"])
            self.assertEqual(sheet_rows[3][8:11], (1, 1, 1))
            self.assertEqual(sheet_rows[3][12:18], (1, 1, 1, 1, 1, 1))
            self.assertEqual(status, "Processed")

    def test_blank_campaigns_share_one_summary_row_without_lookup(self):
        from streamlit.testing.v1 import AppTest

        uploads = {
            "inbound_reply_email_2_upload": Upload("email.csv", csv_bytes(("sender", "Campaign"), [
                ("new@example.com", " "), ("NEW@example.com", "Unspecified campaign"),
                ("assigned@example.com", "Assigned"), ("excluded@example.com", "SBF"),
            ])),
            "inbound_reply_sms_1_upload": Upload("sms_agents.csv", csv_bytes(("Phone",), [("+63 917 123 4567",)])),
            "inbound_reply_sms_2_upload": Upload("sms.csv", csv_bytes(("Phone", "Bank Campaign"), [
                ("09171234567", ""), ("+63 917 123 4567", "SBLEGAL"),
            ])),
            "inbound_reply_viber_2_upload": Upload("viber.csv", csv_bytes(("Phone",), [("09181234567",)])),
        }
        with patch("branding.render_bank_branding"), \
                patch("inbound_reply_page.SAVED_CAMPAIGN_LOOKUP_PATH", ROOT / "tests" / "missing-lookup.zip"), \
                patch("streamlit.file_uploader", side_effect=lambda label, **kw: uploads.get(kw["key"])), \
                patch("streamlit.html") as html, patch("streamlit.download_button") as downloads:
            app = AppTest.from_file(str(ROOT / "pages" / "INBOUND_REPLY.py")).run(timeout=20)
            self.assertFalse(app.exception)
            self.assertFalse(app.error)
            call = next(call for call in html.call_args_list if 'aria-label="Inbound reply summary"' in call.args[0])
            rows = table_cells(call.args[0])
            self.assertEqual([row[0] for row in rows], ["Assigned", "sblegal"])
            self.assertEqual(rows[1][8:11], ["1", "1", "—"])

            uploads["inbound_reply_viber_1_upload"] = Upload("viber_agents.csv", csv_bytes(("Phone",), []))
            app.run(timeout=20)
            self.assertFalse(app.exception)
            self.assertFalse(app.error)
            call = next(call for call in reversed(downloads.call_args_list)
                        if call.kwargs.get("key") == "inbound_reply_summary_download")
            book = load_workbook(BytesIO(call.kwargs["data"]()), data_only=True)
            try:
                summary = book["Summary"]
                self.assertEqual([summary.cell(row, 1).value for row in (4, 5)], ["Assigned", "sblegal"])
                self.assertEqual([summary.cell(5, col).value for col in (9, 10, 11, 13, 15, 17)], [1, 1, 1, None, 1, 0])
                self.assertEqual(book["Viber Unreplied"]["B4"].value, "sblegal")
            finally:
                book.close()

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
                                 ["SBC Level 1", "SBC Level 6", "SBC Mindanao", "sblegal"])
            fresh.button(key="clear_all_files").click().run(timeout=20)
            self.assertNotIn("inbound_reply_email_result", fresh.session_state)
            self.assertEqual(read_saved_campaign_lookup_files(saved_path), original_files)
            self.assertTrue(any("Using saved campaign lookup files" in caption.value for caption in fresh.caption))

            uploads[lookup_key] = [Upload("regular.xlsx", b"broken")]
            app.run(timeout=20)
            self.assertTrue(app.error)
            self.assertFalse(any(button.key == "inbound_reply_save_lookup" for button in app.button))
            self.assertEqual(app.session_state["inbound_reply_email_result"]["report"].totals.unique_people, 10)
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

    def test_uploaded_databases_only_fill_missing_saved_matches_in_summary(self):
        from streamlit.testing.v1 import AppTest
        from inbound_campaign_lookup import read_saved_campaign_lookup_files, save_campaign_lookup_files

        with NamedTemporaryFile(suffix=".zip") as temporary:
            saved_path = Path(temporary.name)
        self.addCleanup(saved_path.unlink, missing_ok=True)
        uploads = all_channel_uploads()
        lookup_key = "inbound_reply_email_lookup_uploads"
        saved_files = [(upload.name, upload.getvalue()) for upload in uploads.pop(lookup_key)]
        save_campaign_lookup_files(saved_path, saved_files)
        uploads[lookup_key] = [
            Upload("regular.csv", csv_bytes(("email", "LEVEL", "Product Type", "Phone"), [
                (" ANA@EXAMPLE.COM ", 6, "PL", "+63 917 123 4567"),
                ("unknown@example.com", 6, "PL", "09261234567"),
            ])),
            Upload("B2.csv", csv_bytes(("Email (01)", "Phone"), [("ana@example.com", "09171234567")])),
        ]

        with patch("branding.render_bank_branding"), \
                patch("inbound_reply_page.SAVED_CAMPAIGN_LOOKUP_PATH", saved_path), \
                patch("streamlit.file_uploader", side_effect=lambda label, **kw: uploads.get(kw["key"])), \
                patch("streamlit.download_button") as downloads:
            app = AppTest.from_file(str(ROOT / "pages" / "INBOUND_REPLY.py")).run(timeout=20)

            def summary_counts():
                self.assertFalse(app.exception)
                self.assertFalse(app.error)
                call = next(call for call in reversed(downloads.call_args_list)
                            if call.kwargs.get("key") == "inbound_reply_summary_download")
                book = load_workbook(BytesIO(call.kwargs["data"]()), data_only=True)
                try:
                    sheet = book["Summary"]
                    return {row[0]: row[8:11] for row in sheet.iter_rows(min_row=4, max_row=sheet.max_row - 1,
                                                                      values_only=True)}
                finally:
                    book.close()

            primary_counts = {"SBC Level 1": (3, 3, 3), "SBC Level 6": (2, 2, 2), "SBC Mindanao": (4, 4, 4)}
            self.assertEqual(summary_counts(), {**primary_counts, "PL RECOVERY": (1, 1, 1)})
            self.assertEqual(app.session_state["inbound_reply_email_result"]["report"].dc.unmatched_addresses, 0)
            self.assertEqual(read_saved_campaign_lookup_files(saved_path), saved_files)

            uploads.pop(lookup_key)
            app.run(timeout=20)
            self.assertEqual(summary_counts(), {**primary_counts, "sblegal": (1, 1, 1)})
            self.assertEqual(app.session_state["inbound_reply_email_result"]["report"].dc.unmatched_addresses, 1)

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
        expected_labels = ["SBC Level 1", "SBC Level 6", "SBC Mindanao", "sblegal"]
        with patch("branding.render_bank_branding", side_effect=render_clear_all_files), \
                patch("inbound_reply_page.inbound_report_date", return_value=REPORT_DATE), \
                patch("inbound_reply_page.SAVED_CAMPAIGN_LOOKUP_PATH", ROOT / "tests" / "missing-lookup.zip"), \
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
            self.assertEqual(rows[3][8:11], ["1", "1", "1"])
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
                self.assertEqual([book["Summary"].cell(7, col).value for col in (9, 10, 11)], [1, 1, 1])
                for title, count in (("SMS Unreplied", 7), ("Viber Unreplied", 6), ("Email Unreplied", 6)):
                    sheet = book[title]
                    self.assertEqual(sheet["B1"].value, "Processed")
                    self.assertEqual(sheet["B2"].value, count)
                    self.assertEqual(sheet.max_row - 3, count)
                    self.assertIn("Source: Lookup Campaign", [cell.value for cell in sheet[3]])
                    self.assertTrue(all(sheet.cell(row, 2).value in expected_labels for row in range(4, sheet.max_row + 1)))
                    self.assertEqual(sheet.cell(sheet.max_row, 2).value, "sblegal")
                for channel in ("email", "sms", "viber"):
                    report = app.session_state[f"inbound_reply_{channel}_result"]["report"]
                    self.assertEqual(report.totals.unique_people, 10)
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
            self.assertEqual([row[0] for row in displayed()], ["SBC Level 6", "SBC Level 7", "sblegal"])
            for channel in ("email", "sms", "viber"):
                report = app.session_state[f"inbound_reply_{channel}_result"]["report"]
                source = report.dc if channel == "email" else report.message_logs
                self.assertEqual(source.unmatched_addresses, 5)
                self.assertEqual(report.totals.unique_people, 10)
            # A replacement with no matching contacts moves all replies to sblegal.
            uploads["inbound_reply_email_lookup_uploads"] = [Upload("other.csv", csv_bytes(
                ("email", "LEVEL", "phone_1"), [("outside@example.com", "Level 1", "09981234567")],
            ))]
            app.run(timeout=20)
            self.assertEqual([row[0] for row in displayed()], ["SBC Level 1", "sblegal"])
            self.assertEqual(displayed()[1][8:11], ["10", "10", "10"])
            for channel in ("email", "sms", "viber"):
                report = app.session_state[f"inbound_reply_{channel}_result"]["report"]
                source = report.dc if channel == "email" else report.message_logs
                self.assertEqual({row.campaign for row in source.records}, {"sblegal"})
                self.assertEqual(source.unmatched_addresses, 10)
                answered = 3 if channel == "sms" else 4
                self.assertEqual(report.totals.unique_people, 10)
                self.assertEqual(report.totals.agent_replied, answered)
                self.assertEqual(report.totals.awaiting_reply, 10 - answered)
                data = (build_email_unreplied_workbook(report) if channel == "email" else
                        build_unreplied_workbook(report))
                book = load_workbook(BytesIO(data))
                try:
                    self.assertEqual(book["Unreplied Details"].max_row, 11 - answered)
                    self.assertEqual(book["Summary"].cell(2, 2).value, 10 - answered)
                finally:
                    book.close()
            call = next(call for call in reversed(downloads.call_args_list)
                        if call.kwargs.get("key") == "inbound_reply_summary_download")
            book = load_workbook(BytesIO(call.kwargs["data"]()))
            try:
                self.assertEqual(book["Summary"].max_row, 6)
                self.assertEqual(book["Summary"]["A4"].value, "SBC Level 1")
                self.assertEqual(book["Summary"]["A5"].value, "sblegal")
                for title, count in (("SMS Unreplied", 7), ("Viber Unreplied", 6), ("Email Unreplied", 6)):
                    self.assertEqual(book[title]["B1"].value, "Processed")
                    self.assertEqual(book[title]["B2"].value, count)
                    self.assertEqual(book[title].max_row, 3 + count)
            finally:
                book.close()
            before_invalid = displayed()
            previous_reports = {
                channel: app.session_state[f"inbound_reply_{channel}_result"]["report"]
                for channel in ("email", "sms", "viber")
            }
            uploads["inbound_reply_email_lookup_uploads"] = [Upload("bad.xlsx", b"broken")]
            app.run(timeout=20)
            self.assertFalse(app.exception)
            self.assertTrue(app.error)
            self.assertEqual(displayed(), before_invalid)
            for channel, report in previous_reports.items():
                self.assertIs(app.session_state[f"inbound_reply_{channel}_result"]["report"], report)
            uploads["inbound_reply_email_lookup_uploads"] = []
            app.run(timeout=20)
            self.assertFalse(app.error)
            self.assertEqual([row[0] for row in displayed()], ["Original Campaign", "sblegal"])
            self.assertEqual([row[8:11] for row in displayed()], [["10", "—", "—"], ["—", "10", "10"]])
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
                patch("inbound_reply_page.SAVED_CAMPAIGN_LOOKUP_PATH", ROOT / "tests" / "missing-lookup.zip"), \
                patch("streamlit.file_uploader", side_effect=lambda label, **kw: uploads.get(kw["key"])), \
                patch("streamlit.download_button") as downloads:
            app = AppTest.from_file(str(ROOT / "pages" / "INBOUND_REPLY.py")).run(timeout=20)
            self.assertFalse(app.exception)
            self.assertFalse(app.error)
            self.assertEqual([metric.value for metric in app.metric], ["3", "—", "—"])
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
