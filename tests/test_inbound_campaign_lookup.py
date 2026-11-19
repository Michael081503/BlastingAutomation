from __future__ import annotations

import csv
import unittest
from datetime import date
from io import BytesIO, StringIO
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from inbound_campaign_lookup import (
    apply_email_campaign_lookup, apply_phone_campaign_lookup,
    build_email_campaign_lookup, read_email_campaign_source,
)
from inbound_email import build_email_unreplied_workbook, match_email_sources, read_email_source
from inbound_reply import InboundReplyError, match_viber_sources, read_reply_source
from inbound_summary import build_summary_workbook
from test_inbound_display import table_cells
from test_inbound_email import csv_bytes


ROOT = Path(__file__).resolve().parents[1]


def workbook_bytes(sheets):
    book = Workbook()
    book.remove(book.active)
    for name, rows in sheets:
        sheet = book.create_sheet(name)
        for row in rows:
            sheet.append(row)
    output = BytesIO()
    book.save(output)
    book.close()
    return output.getvalue()


def regular_file(level="Level 1"):
    return workbook_bytes([("OVERALL", [
        ("Name", "debtor_email", "LEVEL"),
        ("Client", "client@example.com", level),
    ])])


def geo_file():
    return workbook_bytes([("OVERALL", [
        ("name", "email", "placement", "Area"),
        ("North", "north@example.com", "Level 6", "Region 3"),
        ("South", "south@example.com", "Level 1", "Mindanao"),
    ])])


class CampaignLookupTests(unittest.TestCase):
    def test_levels_two_to_five_merge_into_level_one_without_duplicating_replies(self):
        data = workbook_bytes([("OVERALL", [
            ("debtor_email", "LEVEL", "phone_1"),
            *(("shared@example.com", f"Level {level}", "09171234567") for level in range(1, 6)),
            ("six@example.com", "Level 6", "09181234567"),
        ])])
        lookup = build_email_campaign_lookup([("Regular.xlsx", data)])
        self.assertEqual(lookup.campaigns, ("SBC Level 1", "SBC Level 6"))
        self.assertEqual(lookup.campaigns_by_email, {
            "shared@example.com": ("SBC Level 1",), "six@example.com": ("SBC Level 6",),
        })
        self.assertEqual(lookup.campaigns_by_phone, {
            "09171234567": ("SBC Level 1",), "09181234567": ("SBC Level 6",),
        })
        dc = read_email_source("dc.csv", csv_bytes(("sender",), [
            ("SHARED@example.com",), ("shared@example.com",), ("six@example.com",),
        ]), kind="dc", require_campaign=False)
        history = read_email_source("history.csv", csv_bytes(("email", "sent"), [
            ("shared@example.com", "Agent response"), ("shared@example.com", "Second response"),
        ]), kind="history")
        email = match_email_sources(apply_email_campaign_lookup(dc, lookup), history)
        self.assertEqual([(c.campaign, c.unique_people, c.agent_replied) for c in email.campaigns],
                         [("SBC Level 1", 1, 1), ("SBC Level 6", 1, 0)])
        customers = read_reply_source("customers.csv", csv_bytes(("Phone",), [
            ("09171234567",), ("+63 917 123 4567",), ("09181234567",),
        ]), require_campaign=False)
        agents = read_reply_source("agents.csv", csv_bytes(("Phone",), [
            ("09171234567",), ("09171234567",),
        ]), require_campaign=False)
        phone = match_viber_sources(apply_phone_campaign_lookup(customers, lookup),
                                   apply_phone_campaign_lookup(agents, lookup),
                                   agent_replied_source="campaign_replies", match_mode="phone")
        self.assertEqual([(c.campaign, c.unique_people, c.agent_replied) for c in phone.campaigns],
                         [("SBC Level 1", 1, 1), ("SBC Level 6", 1, 0)])
        output = build_summary_workbook("", date(2026, 11, 19), "", phone, email, sms_report=phone)
        book = load_workbook(BytesIO(output))
        try:
            self.assertEqual(book.sheetnames, ["Summary", "SMS Unreplied", "Viber Unreplied", "Email Unreplied"])
            self.assertEqual([book["Summary"].cell(row, 1).value for row in (4, 5)], ["SBC Level 1", "SBC Level 6"])
            self.assertEqual([book["Summary"].cell(4, col).value for col in (9, 10, 11, 13, 15, 17)], [1] * 6)
            for name in ("SMS Unreplied", "Viber Unreplied", "Email Unreplied"):
                self.assertEqual(book[name]["B4"].value, "SBC Level 6")
        finally:
            book.close()

    def test_real_header_shapes_use_area_before_placement_and_level_without_filename_or_order(self):
        files = [("first.xlsx", geo_file()), ("second.xlsx", regular_file("LEVEL 2"))]
        expected = {
            "client@example.com": ("SBC Level 1",),
            "north@example.com": ("SBC Region 3",),
            "south@example.com": ("SBC Mindanao",),
        }
        self.assertEqual(build_email_campaign_lookup(files).campaigns_by_email, expected)
        self.assertEqual(build_email_campaign_lookup(list(reversed(files))).campaigns_by_email, expected)

    def test_overall_is_used_without_status_sheet_duplicates_or_stale_memberships(self):
        data = workbook_bytes([
            ("NEGATIVE", [("email", "placement"), ("client@example.com", "Level 6")]),
            ("OVERALL", [("email", "placement"), (" CLIENT@EXAMPLE.COM ", "LEVEL 1"),
                         ("client@example.com", "Level 1"), ("client@example.com", "Level 7")]),
        ])
        source = read_email_campaign_source("active.xlsx", data)
        self.assertEqual(source.sheets, ("OVERALL",))
        self.assertEqual(source.campaigns_by_email, {"client@example.com": ("SBC Level 1", "SBC Level 7")})

    def test_fallback_sheets_csv_blank_levels_and_invalid_addresses(self):
        data = workbook_bytes([
            ("OVERALL", [("Unrelated",), ("Cover page",)]),
            ("First", [("debtor_email", "LEVEL"), ("client@example.com", "level 2"),
                       ("missing@example.com", None), ("invalid address", "Level 1")]),
            ("Second", [("debtor_email", "LEVEL"), ("client@example.com", "Level 2")]),
        ])
        source = read_email_campaign_source("active.xlsx", data)
        self.assertEqual(source.sheets, ("First", "Second"))
        self.assertEqual(source.campaigns_by_email, {"client@example.com": ("SBC Level 1",)})
        self.assertEqual((source.invalid_email_rows, source.missing_campaign_rows), (1, 1))
        csv_source = read_email_campaign_source("active.csv", csv_bytes(("email", "Area", "placement"), [
            ("north@example.com", "REGION 3", "Level 6"),
            ("south@example.com", "mindanao", "Level 1"),
        ]))
        self.assertEqual(csv_source.campaigns_by_email, {
            "north@example.com": ("SBC Region 3",), "south@example.com": ("SBC Mindanao",),
        })

    def test_bad_lookup_is_rejected_instead_of_using_a_partial_index(self):
        for name, data in (("bad.xlsx", b"broken"), ("wrong.csv", b"email,Unrelated\nclient@example.com,test\n")):
            with self.subTest(name=name), self.assertRaises(InboundReplyError):
                build_email_campaign_lookup([("good.xlsx", regular_file()), (name, data)])

    def test_semicolon_csv_with_multiline_notes_works_for_lookup_and_email_uploads(self):
        def data(headers, rows):
            output = StringIO(newline="")
            writer = csv.writer(output, delimiter=";")
            writer.writerow(headers)
            writer.writerows(rows)
            return output.getvalue().encode("utf-8-sig")

        lookup = build_email_campaign_lookup([("active.csv", data(("debtor_email", "LEVEL", "Notes"), [
            ("client@example.com", "Level 1", "first line\nsecond line"),
        ]))])
        dc = read_email_source("dc.csv", data(("sender", "Notes"), [
            ("client@example.com", "first line\nsecond line"),
        ]), kind="dc", require_campaign=False)
        history = read_email_source("history.csv", data(("email", "sent"), [
            ("client@example.com", "first line\nsecond line"),
        ]), kind="history")
        report = match_email_sources(apply_email_campaign_lookup(dc, lookup), history)
        self.assertEqual([(c.campaign, c.agent_replied) for c in report.campaigns], [("SBC Level 1", 1)])

    def test_matching_and_both_downloads_exclude_unmatched_accounts_and_keep_original_fields(self):
        files = [
            ("regular.xlsx", regular_file()),
            ("extra.csv", csv_bytes(("email", "Area"), [("client@example.com", "Mindanao")])),
        ]
        dc = read_email_source("dc.csv", csv_bytes(("ID", "sender", "Campaign", "Email"), [
            ("001", " CLIENT@example.com ", "Generic campaign", "Respond Now"),
            ("002", "client@example.com", "Generic campaign", "Duplicate"),
            ("003", "unknown@example.com", "Generic campaign", "Respond Now"),
            ("003-DUPLICATE", " UNKNOWN@example.com ", "Generic campaign", "Duplicate"),
            ("004", "sent@example.com", "Generic campaign", "Respond Now"),
            ("005", "sbf@example.com", "MC4 SBF PL", "Respond Now"),
        ]), kind="dc")
        resolved = apply_email_campaign_lookup(dc, build_email_campaign_lookup(files))
        history = read_email_source("history.csv", csv_bytes(("email", "sent", "receive"), [
            ("sent@example.com", "Agent reply", None),
            ("client@example.com", None, "Customer message"),
        ]), kind="history")
        report = match_email_sources(resolved, history)
        self.assertEqual([(c.campaign, c.unique_people, c.agent_replied, c.awaiting_reply) for c in report.campaigns], [
            ("SBC Level 1", 1, 0, 1), ("SBC Mindanao", 1, 0, 1),
        ])
        self.assertEqual(report.totals.unique_people, 1)
        self.assertEqual(report.totals.agent_replied, 0)
        self.assertEqual({row.email for row in resolved.records}, {"client@example.com"})
        self.assertEqual(resolved.unmatched_addresses, 2)
        self.assertEqual(len(dc.records), 5)
        self.assertEqual(dc.headers, ("ID", "sender", "Campaign", "Email"))
        book = load_workbook(BytesIO(build_email_unreplied_workbook(report)))
        try:
            values = list(book["Unreplied Details"].values)
            self.assertEqual(values[0], (*dc.headers, "Lookup Campaign", "Matched Email", "Reply Status"))
            self.assertEqual([row[4] for row in values[1:]], ["SBC Level 1", "SBC Mindanao"])
            self.assertEqual(values[1][:4], ("001", " CLIENT@example.com ", "Generic campaign", "Respond Now"))
            self.assertEqual(book["Summary"].cell(4, 2).value, 2)
        finally:
            book.close()
        book = load_workbook(BytesIO(build_summary_workbook("", date(2026, 11, 18), "—", email_report=report)))
        try:
            self.assertEqual([book["Summary"].cell(i, 1).value for i in (4, 5)],
                             ["SBC Level 1", "SBC Mindanao"])
            self.assertEqual(book["Summary"].max_row, 6)
            values = list(book["Email Unreplied"].values)
            self.assertEqual(len(values[3:]), 2)
            self.assertEqual(book["Email Unreplied"]["B2"].value, 2)
            self.assertIn("Source: Lookup Campaign", values[2])
            self.assertEqual([row[1] for row in values[3:]], ["SBC Level 1", "SBC Mindanao"])
        finally:
            book.close()


class CampaignLookupPageTests(unittest.TestCase):
    def test_preview_processing_same_name_replacement_errors_removal_and_clear_update_results(self):
        from streamlit.testing.v1 import AppTest
        from upload_state import render_clear_all_files

        class Upload(BytesIO):
            def __init__(self, name, data):
                super().__init__(data)
                self.name = name

        uploads = {
            "inbound_reply_email_2_upload": Upload("dc.csv", csv_bytes(("ID", "sender", "Email"), [
                ("001", "client@example.com", "Respond Now"),
                ("002", "south@example.com", "Respond Now"),
            ])),
            "inbound_reply_email_lookup_uploads": [Upload("active.xlsx", regular_file()), Upload("other.xlsx", geo_file())],
        }

        def rows(html):
            call = next(call for call in reversed(html.call_args_list)
                        if 'aria-label="Inbound reply summary"' in call.args[0])
            return table_cells(call.args[0])

        with patch("branding.render_bank_branding", side_effect=render_clear_all_files), \
                patch("streamlit.file_uploader", side_effect=lambda label, **kwargs: uploads.get(kwargs["key"])), \
                patch("streamlit.html") as html, patch("streamlit.download_button") as downloads:
            app = AppTest.from_file(str(ROOT / "pages" / "INBOUND_REPLY.py")).run(timeout=20)
            self.assertFalse(app.exception)
            self.assertFalse(app.error)
            self.assertEqual([row[0] for row in rows(html)], ["SBC Level 1", "SBC Mindanao", "SBC Region 3"])
            self.assertTrue(all(row[8] == "—" for row in rows(html)))
            uploads["inbound_reply_email_1_upload"] = Upload("history.csv", csv_bytes(("email", "sent"), [
                ("client@example.com", "Agent reply"),
            ]))
            app.run(timeout=20)
            self.assertFalse(app.exception)
            self.assertFalse(app.error)
            self.assertEqual([metric.value for metric in app.metric], ["2", "1", "1"])
            self.assertEqual([row[12] for row in rows(html)], ["1", "0", "—"])
            call = next(call for call in reversed(downloads.call_args_list)
                        if call.kwargs.get("key") == "inbound_reply_summary_download")
            book = load_workbook(BytesIO(call.kwargs["data"]()))
            try:
                self.assertEqual(book["Email Unreplied"].cell(4, 2).value, "SBC Mindanao")
            finally:
                book.close()
            uploads["inbound_reply_email_lookup_uploads"][0] = Upload("active.xlsx", regular_file("Level 7"))
            app.run(timeout=20)
            self.assertFalse(app.exception)
            self.assertEqual([row[0] for row in rows(html)], ["SBC Level 7", "SBC Mindanao", "SBC Region 3"])
            uploads["inbound_reply_email_lookup_uploads"] = [Upload("active.xlsx", regular_file("Level 7"))]
            app.run(timeout=20)
            self.assertFalse(app.exception)
            self.assertEqual([row[0] for row in rows(html)], ["SBC Level 7"])
            self.assertEqual([metric.value for metric in app.metric], ["1", "1", "0"])
            self.assertTrue(any("Excluded 1 email addresses without a campaign match" in caption.value for caption in app.caption))
            uploads["inbound_reply_email_lookup_uploads"] = [Upload("bad.xlsx", b"broken")]
            app.run(timeout=20)
            self.assertFalse(app.exception)
            self.assertTrue(app.error)
            self.assertNotIn("inbound_reply_email_result", app.session_state)
            self.assertNotIn("inbound_reply_campaign_labels", app.session_state)
            self.assertEqual(len(app.metric), 0)
            uploads["inbound_reply_email_lookup_uploads"] = []
            app.run(timeout=20)
            self.assertFalse(app.exception)
            self.assertTrue(app.error)  # Campaign is required again when no lookup files are present.
            self.assertNotIn("inbound_reply_email_lookup", app.session_state)
            uploads["inbound_reply_email_lookup_uploads"] = [Upload("active.xlsx", regular_file())]
            app.run(timeout=20)
            self.assertIn("report", app.session_state["inbound_reply_email_result"])
            app.button(key="clear_all_files").click().run(timeout=20)
            self.assertFalse(app.exception)
            for key in ("inbound_reply_email_lookup", "inbound_reply_email_result", "inbound_reply_campaign_labels"):
                self.assertNotIn(key, app.session_state)
            self.assertEqual(len(app.metric), 0)


if __name__ == "__main__":
    unittest.main()
