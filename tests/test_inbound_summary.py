from __future__ import annotations

import unittest
from datetime import date
from io import BytesIO
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from inbound_email import match_email_sources, read_email_source
from inbound_reply import CampaignReplyCounts, match_viber_sources, read_reply_source
from inbound_summary import build_summary_workbook, campaign_totals_key, display_cell, summary_rows
from test_inbound_email import DC_HEADERS, HISTORY_HEADERS, csv_bytes


ROOT = Path(__file__).resolve().parents[1]


def reports():
    dc = read_email_source("DC.csv", csv_bytes(DC_HEADERS, [
        ("000001", "pending@example.com", "Respond Now", "A", None, "=SUM(1,1)"),
        ("000002", "replied@example.com", "Respond Now", "A", None, None),
    ]), kind="dc")
    history = read_email_source("History.csv", csv_bytes(HISTORY_HEADERS, [
        ("H1", "replied@example.com", None, "Agent response", "2026-11-16", "Agent"),
    ]), kind="history")
    headers = ("ID", "Client Name", "Phone", "Bank Campaign", "Message Content")
    logs = read_reply_source("logs.csv", csv_bytes(headers, [
        ("000003", "Replied", "09171234567", "A", "Source message"),
        ("000004", "Pending", "09181234567", "B", "Keep original text"),
    ]))
    replies = read_reply_source("replies.csv", csv_bytes(headers, [
        ("000005", "Replied", "+63 917 123 4567", "A", "Reply message"),
    ]))
    return (match_viber_sources(logs, replies, agent_replied_source="campaign_replies", match_mode="phone"),
            match_email_sources(dc, history))


class GroupedSummaryTests(unittest.TestCase):
    def test_all_eighteen_sample_columns_match_the_reference_order_and_percentages(self):
        from inbound_reply_page import build_report_html
        from test_inbound_display import table_cells

        label = "SAMPLE CAMPAIGN"
        report_date = date(2026, 11, 12)
        email = SimpleNamespace(campaigns=(CampaignReplyCounts(label, 51, 20, 31, 20),))
        sms = SimpleNamespace(campaigns=(CampaignReplyCounts(label, 512, 500, 12, 500),))
        viber = SimpleNamespace(campaigns=(CampaignReplyCounts(label, 321, 10, 311, 10),))
        totals = {campaign_totals_key(report_date, label): {
            "active_ticket": 58431, "dc_email": 46104, "dc_sms": 58412, "dc_viber": 1000,
        }}
        expected = [label, "58,431", "46,104", "78.90%", "58,412", "99.97%", "1,000", "1.71%",
                    "51", "512", "321", "YES", "20", "39.22%", "500", "97.66%", "10", "3.12%"]
        rows = summary_rows(label, report_date, "YES", viber, email, totals, sms)
        self.assertEqual([display_cell(value, column) for column, value in enumerate(rows[0])], expected)
        html = build_report_html(label, report_date, "YES", viber, email, totals, sms_report=sms)
        self.assertEqual(table_cells(html), [expected])

    def test_excel_summary_has_reference_merges_and_second_sheet_has_all_unreplied_details(self):
        from openpyxl import load_workbook

        viber, email = reports()
        report_date = date(2026, 11, 12)
        totals = {campaign_totals_key(report_date, "A"): {
            "active_ticket": 100, "dc_email": 50, "dc_viber": 60,
        }}
        book = load_workbook(BytesIO(build_summary_workbook("", report_date, "YES", viber, email, totals)))
        try:
            self.assertEqual(book.sheetnames, ["Summary", "Unreplied Details"])
            sheet = book["Summary"]
            self.assertEqual({str(area) for area in sheet.merged_cells.ranges},
                             {"A1:K1", "L1:L3", "M1:R2", "A2:A3", "B2:B3", "C2:H2", "I2:K2"})
            self.assertEqual(sheet.max_column, 18)
            self.assertEqual(sheet["A1"].value, "NOVEMBER 12, 2026 DATA")
            self.assertEqual(sheet["L1"].value, "SENT TO OPS")
            self.assertEqual(sheet["M1"].value, "AGENT WORKED")
            self.assertEqual(sheet["B4"].value, 100)
            self.assertEqual(sheet["B4"].data_type, "n")
            self.assertEqual(sheet["C4"].value, 50)
            self.assertEqual(sheet["D4"].value, .5)
            self.assertEqual(sheet["D4"].number_format, "0.00%")
            self.assertEqual(sheet["I4"].value, 2)
            self.assertEqual(sheet["M4"].value, 1)
            self.assertEqual(sheet["N4"].value, .5)
            self.assertIsNone(sheet["J4"].value)
            self.assertIsNone(sheet["O4"].value)
            self.assertEqual(sheet.freeze_panes, "C4")
            details = book["Unreplied Details"]
            values = list(details.values)
            rows = [dict(zip(values[0], row)) for row in values[1:]]
            self.assertEqual(len(rows), 2)
            self.assertEqual({(row["Channel"], row["Campaign"]) for row in rows}, {("EMAIL", "A"), ("VIBER", "B")})
            email_row = next(row for row in rows if row["Channel"] == "EMAIL")
            viber_row = next(row for row in rows if row["Channel"] == "VIBER")
            self.assertEqual(email_row["Email"], "pending@example.com")
            self.assertEqual(email_row["Source: Id"], "000001")
            self.assertEqual(email_row["Source: AI Suggestion"], "=SUM(1,1)")
            self.assertEqual(viber_row["Phone"], "09181234567")
            self.assertEqual(viber_row["Source: ID"], "000004")
            self.assertEqual(viber_row["Source: Message Content"], "Keep original text")
            self.assertTrue(all(row["Reply Status"] == "Not replied" for row in rows))
            formula_column = list(values[0]).index("Source: AI Suggestion") + 1
            self.assertEqual(details.cell(2, formula_column).data_type, "s")
        finally:
            book.close()

    def test_manual_only_report_keeps_missing_channels_blank(self):
        from openpyxl import load_workbook

        report_date = date(2026, 11, 12)
        totals = {campaign_totals_key(report_date, "Manual"): {"active_ticket": 0, "dc_email": 10}}
        book = load_workbook(BytesIO(build_summary_workbook("Manual", report_date, "NO", campaign_totals=totals)))
        try:
            self.assertEqual(book["Summary"]["B4"].value, 0)
            self.assertEqual(book["Summary"]["C4"].value, 10)
            self.assertIsNone(book["Summary"]["D4"].value)
            self.assertIsNone(book["Summary"]["I4"].value)
            self.assertEqual(book["Unreplied Details"].max_row, 1)
        finally:
            book.close()

    def test_main_download_uses_grouped_summary_and_controls_are_collapsed(self):
        from openpyxl import load_workbook
        from streamlit.testing.v1 import AppTest
        from upload_state import render_clear_all_files

        viber, email = reports()
        with patch("branding.render_bank_branding", side_effect=render_clear_all_files), \
                patch("streamlit.file_uploader", return_value=None), \
                patch("inbound_reply_page.prepare_viber_report", return_value=viber), \
                patch("inbound_reply_page.prepare_email_report", return_value=email), \
                patch("streamlit.download_button") as downloads:
            app = AppTest.from_file(str(ROOT / "pages" / "INBOUND_REPLY.py")).run(timeout=20)
            self.assertFalse(app.exception)
            call = next(call for call in downloads.call_args_list if call.kwargs.get("key") == "inbound_reply_summary_download")
            book = load_workbook(BytesIO(call.kwargs["data"]()))
            try:
                self.assertEqual(book["Summary"].max_column, 18)
                self.assertEqual(book.sheetnames, ["Summary", "Unreplied Details"])
            finally:
                book.close()
            for label in ("Report settings", "Upload files", "Edit campaign totals", "Email account details", "Viber account details"):
                panel = next(panel for panel in app.expander if panel.label == label)
                self.assertFalse(panel.proto.expanded)


if __name__ == "__main__":
    unittest.main()
