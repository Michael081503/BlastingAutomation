from __future__ import annotations

import csv
import unittest
from datetime import date
from io import BytesIO, StringIO
from pathlib import Path
from unittest.mock import patch

from inbound_email import (
    build_email_unreplied_workbook, match_email_sources, normalized_email, read_email_source,
)
from inbound_reply import InboundReplyError, match_viber_sources, read_reply_source


ROOT = Path(__file__).resolve().parents[1]
DC_HEADERS = ("Id", "sender", "Email", "Campaign", "Agent Lark", "AI Suggestion")
HISTORY_HEADERS = ("Id", "email", "receive", "sent", "Date", "Person")


def csv_bytes(headers, rows):
    output = StringIO(newline="")
    writer = csv.writer(output)
    writer.writerow(headers)
    writer.writerows(rows)
    return output.getvalue().encode("utf-8-sig")


def make_report(dc_rows, history_rows):
    return match_email_sources(
        read_email_source("DC.csv", csv_bytes(DC_HEADERS, dc_rows), kind="dc"),
        read_email_source("History.csv", csv_bytes(HISTORY_HEADERS, history_rows), kind="history"),
    )


class EmailMatchingTests(unittest.TestCase):
    def test_dc_population_deduplicates_sender_and_only_sent_history_marks_reply(self):
        report = make_report([
            ("01", " ANA@example.com ", "Respond Now", "B2", "Assigned Agent", "AI suggestion"),
            ("02", "ana@EXAMPLE.com", "Respond Now", "b2", "Assigned Agent", "AI suggestion"),
            ("03", "ben@example.com", "Respond Now", "B2", "Assigned Agent", "AI suggestion"),
            ("04", "cara@example.com", "Respond Now", "B2", "Assigned Agent", "AI suggestion"),
        ], [
            ("H1", "ana@example.com", None, "Agent reply", "2026-11-16", "Agent One"),
            ("H2", "ana@example.com", None, "Second reply", "2026-11-16", "Agent Two"),
            ("H3", "ben@example.com", "Received message", " ", "2026-11-16", "Agent Three"),
            ("H4", "outside@example.com", None, "Unrelated sent email", "2026-11-16", "Agent Four"),
        ])
        self.assertEqual(report.totals.unique_people, 3)
        self.assertEqual(report.totals.agent_replied, 1)
        self.assertEqual(report.totals.awaiting_reply, 2)
        self.assertEqual(report.totals.matched_both_files, 2)
        self.assertEqual(report.history_only_addresses, 1)
        self.assertEqual(len(report.campaigns), 1)
        self.assertAlmostEqual(report.campaigns[0].agent_percent, 1 / 3)
        self.assertEqual(len(report.accounts), 3)

    def test_different_addresses_with_same_agent_or_name_remain_separate(self):
        report = make_report([
            ("1", "ana+one@example.com", "Respond Now", "B2", "Same Name", "Draft only"),
            ("2", "ana+two@example.com", "Respond Now", "B2", "Same Name", "Draft only"),
        ], [("H1", "ana+one@example.com", None, "Actual sent reply", "2026-11-16", "Same Name")])
        self.assertEqual(report.totals.unique_people, 2)
        self.assertEqual(report.totals.agent_replied, 1)
        self.assertEqual(report.totals.awaiting_reply, 1)

    def test_campaigns_come_from_dc_and_sent_history_applies_to_matching_address(self):
        report = make_report([
            ("1", "ana@example.com", "Respond Now", "B2", None, None),
            ("2", "ana@example.com", "Respond Now", "B4", None, None),
        ], [("H1", "ana@example.com", None, "Sent reply", "2026-11-16", "Agent One")])
        self.assertEqual(report.totals.unique_people, 1)
        self.assertEqual([(row.campaign, row.agent_replied) for row in report.campaigns], [("B2", 1), ("B4", 1)])

    def test_empty_history_leaves_all_dc_accounts_awaiting_reply(self):
        report = make_report([("1", "ana@example.com", "Respond Now", "B2", None, None)], [])
        self.assertEqual(report.totals.awaiting_reply, 1)
        self.assertEqual(report.totals.agent_replied, 0)
        empty = make_report([], [("H1", "outside@example.com", None, "Sent", "2026-11-16", "Agent")])
        self.assertEqual(empty.totals.unique_people, 0)
        self.assertEqual(empty.campaigns, ())


class EmailReaderAndWorkbookTests(unittest.TestCase):
    def test_dc_email_workflow_label_is_not_used_as_address(self):
        source = read_email_source("DC.csv", csv_bytes(DC_HEADERS, [
            ("1", "client@example.com", "Respond Now", "B2", "Agent", "Suggestion"),
        ]), kind="dc")
        self.assertEqual(source.records[0].email, "client@example.com")
        self.assertFalse(source.records[0].sent)
        self.assertEqual(source.records[0].source_values[2], "Respond Now")

    def test_invalid_addresses_are_counted_and_missing_headers_are_rejected(self):
        source = read_email_source("DC.csv", csv_bytes(DC_HEADERS, [
            ("1", "not an address", "Respond Now", "B2", None, None),
            ("2", "good@example.com", "Respond Now", "B2", None, None),
        ]), kind="dc")
        self.assertEqual(source.excluded_rows, 1)
        self.assertEqual(len(source.records), 1)
        with self.assertRaises(InboundReplyError):
            read_email_source("wrong.csv", csv_bytes(HISTORY_HEADERS, []), kind="dc")
        self.assertEqual(normalized_email("  Client@EXAMPLE.COM "), "client@example.com")
        self.assertEqual(normalized_email("client @example.com"), "")

    def test_xlsx_ignores_incorrect_dimensions_and_reads_all_data_rows(self):
        from openpyxl import Workbook
        from zipfile import ZipFile
        import re

        book = Workbook()
        book.active.title = "mc4_emails"
        book.active.append(DC_HEADERS)
        book.active.append(("1", "ana@example.com", "Respond Now", "B2", None, None))
        book.active.append(("2", "ben@example.com", "Respond Now", "B2", None, None))
        original = BytesIO()
        book.save(original)
        book.close()
        damaged = BytesIO()
        with ZipFile(BytesIO(original.getvalue())) as source, ZipFile(damaged, "w") as target:
            for item in source.infolist():
                data = source.read(item.filename)
                if item.filename == "xl/worksheets/sheet1.xml":
                    data = re.sub(rb'<dimension ref="[^"]+"', b'<dimension ref="A1:A1"', data)
                target.writestr(item, data)
        result = read_email_source("DC.xlsx", damaged.getvalue(), kind="dc")
        self.assertEqual(len(result.records), 2)
        self.assertEqual(result.sheet, "mc4_emails")

    def test_unreplied_workbook_preserves_dc_fields_and_matches_campaign_counts(self):
        from openpyxl import load_workbook

        report = make_report([
            ("000000000000000001", " ANA@example.com ", "Respond Now", "B2", None, "=SUM(1,1)"),
            ("2", "ana@example.com", "Respond Now", "B2", None, "Duplicate"),
            ("3", "ana@example.com", "Respond Now", "B4", None, "Another campaign"),
            ("4", "ben@example.com", "Respond Now", "B2", None, None),
        ], [("H1", "ben@example.com", None, "Sent reply", "2026-11-16", "Agent")])
        book = load_workbook(BytesIO(build_email_unreplied_workbook(report)))
        try:
            self.assertEqual(book.sheetnames, ["Summary", "Unreplied Details"])
            self.assertEqual(list(book["Summary"].values), [
                ("Campaign", "Not replied"), ("B2", 1), ("B4", 1), ("TOTAL", 2),
            ])
            sheet = book["Unreplied Details"]
            values = list(sheet.values)
            self.assertEqual(values[0], (*DC_HEADERS, "Matched Email", "Reply Status"))
            self.assertEqual(len(values) - 1, 2)
            self.assertEqual(values[1][0], "000000000000000001")
            self.assertEqual(values[1][1], " ANA@example.com ")
            self.assertEqual(values[1][-2:], ("ana@example.com", "Not replied"))
            self.assertEqual(sheet["F2"].value, "=SUM(1,1)")
            self.assertEqual(sheet["F2"].data_type, "s")
        finally:
            book.close()


class EmailPageTests(unittest.TestCase):
    def test_email_and_viber_columns_join_by_campaign_without_filling_missing_channels(self):
        from inbound_reply_page import build_report_html

        email = make_report([("1", "ana@example.com", "Respond Now", "B2", None, None)], [])
        viber_source = csv_bytes(("Client Name", "Phone", "Bank Campaign"), [("Ana", "09171234567", "b2")])
        empty_viber = csv_bytes(("Client Name", "Phone", "Bank Campaign"), [])
        viber = match_viber_sources(read_reply_source("logs.csv", viber_source),
                                    read_reply_source("replies.csv", empty_viber),
                                    agent_replied_source="campaign_replies", match_mode="phone")
        html = build_report_html("", date(2026, 11, 18), "YES", viber, email)
        body = html.split("<tbody>")[1].split("</tbody>")[0]
        self.assertEqual(body.count("<tr>"), 1)
        values = [cell.split("</td>")[0] for cell in body.split("<td>")[1:]]
        self.assertEqual(values[8], "1")
        self.assertEqual(values[10], "1")
        self.assertEqual(values[12:14], ["0", "0.00%"])
        self.assertEqual(values[16:18], ["0", "0.00%"])
        self.assertEqual(values[9], "—")

    def test_email_only_download_replacements_and_clear_leave_no_stale_results(self):
        from openpyxl import load_workbook
        from streamlit.testing.v1 import AppTest
        from upload_state import render_clear_all_files

        class Upload(BytesIO):
            def __init__(self, name, data):
                super().__init__(data)
                self.name = name

        uploads = {
            "inbound_reply_email_1_upload": Upload("history.csv", csv_bytes(HISTORY_HEADERS, [
                ("H1", "ana@example.com", None, "Sent reply", "2026-11-16", "Agent"),
            ])),
            "inbound_reply_email_2_upload": Upload("DC.csv", csv_bytes(DC_HEADERS, [
                ("1", "ana@example.com", "Respond Now", "B2", None, None),
                ("2", "ben@example.com", "Respond Now", "B2", None, None),
            ])),
        }
        with patch("branding.render_bank_branding", side_effect=render_clear_all_files), \
                patch("streamlit.file_uploader", side_effect=lambda label, **kwargs: uploads.get(kwargs["key"])), \
                patch("streamlit.download_button") as downloads:
            app = AppTest.from_file(str(ROOT / "pages" / "INBOUND_REPLY.py")).run(timeout=20)
            self.assertFalse(app.exception)
            self.assertFalse(app.error)
            self.assertEqual([metric.value for metric in app.metric], ["2", "1", "1"])
            self.assertNotIn("inbound_reply_viber_result", app.session_state)
            call = next(call for call in downloads.call_args_list
                        if call.kwargs.get("key") == "inbound_reply_email_unreplied_download")
            book = load_workbook(BytesIO(call.kwargs["data"]()))
            try:
                self.assertEqual(book["Summary"]["B2"].value, 1)
                self.assertEqual(book["Unreplied Details"].cell(2, 2).value, "ben@example.com")
            finally:
                book.close()
            uploads["inbound_reply_email_1_upload"] = Upload("history.csv", csv_bytes(HISTORY_HEADERS, []))
            app.run(timeout=20)
            self.assertFalse(app.exception)
            self.assertEqual([metric.value for metric in app.metric], ["2", "0", "2"])
            uploads["inbound_reply_email_1_upload"] = Upload("history.csv", b"Wrong,Headers\n1,2\n")
            app.run(timeout=20)
            self.assertFalse(app.exception)
            self.assertTrue(app.error)
            self.assertNotIn("report", app.session_state["inbound_reply_email_result"])
            self.assertEqual(len(app.metric), 0)
            del uploads["inbound_reply_email_1_upload"]
            app.run(timeout=20)
            self.assertNotIn("inbound_reply_email_result", app.session_state)
            uploads["inbound_reply_email_1_upload"] = Upload("history.csv", csv_bytes(HISTORY_HEADERS, []))
            app.run(timeout=20)
            self.assertIn("report", app.session_state["inbound_reply_email_result"])
            app.button(key="clear_all_files").click().run(timeout=20)
            self.assertFalse(app.exception)
            self.assertNotIn("inbound_reply_email_result", app.session_state)
            self.assertEqual(len(app.metric), 0)


if __name__ == "__main__":
    unittest.main()
