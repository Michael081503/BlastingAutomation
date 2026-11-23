from __future__ import annotations

import unittest
from datetime import date
from io import BytesIO
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from inbound_email import build_email_unreplied_workbook, match_email_sources, read_email_source
from inbound_reply import CampaignReplyCounts, match_viber_sources, read_reply_source
from inbound_summary import (
    OVERALL_BLAST_LABEL, build_summary_workbook, campaign_totals_key, display_cell,
    overall_blast_row, summary_rows,
)
from test_inbound_email import DC_HEADERS, HISTORY_HEADERS, csv_bytes


ROOT = Path(__file__).resolve().parents[1]
OUTPUT_SHEETS = ["Summary", "SMS Unreplied", "Viber Unreplied", "Email Unreplied"]


def detail_rows(sheet):
    values = list(sheet.iter_rows(min_row=3, values_only=True))
    return [dict(zip(values[0], row)) for row in values[1:]]


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


class OverallBlastCoverageTests(unittest.TestCase):
    report_date = date(2026, 11, 18)

    def totals(self):
        return {campaign_totals_key(self.report_date, label): values for label, values in (
            ("Small", {"active_ticket": 100, "dc_email": 50, "dc_sms": 100, "dc_viber": 50}),
            ("Large", {"active_ticket": 900, "dc_email": 90, "dc_sms": 100, "dc_viber": 450}),
            ("Removed campaign", {"active_ticket": 9999, "dc_email": 9999, "dc_sms": 9999, "dc_viber": 9999}),
        )}

    def rows(self, totals=None, *, campaigns=("Small", "Large"), report_date=None):
        return summary_rows("", report_date or self.report_date, "",
                            campaign_totals=self.totals() if totals is None else totals, source_campaigns=campaigns)

    def test_overall_coverage_divides_summed_counts_and_excludes_other_campaigns_and_dates(self):
        overall = overall_blast_row(self.rows())
        self.assertEqual(overall[:8], [OVERALL_BLAST_LABEL, 1000, 140, .14, 200, .2, 500, .5])
        self.assertEqual(overall[8:], [None] * 10)
        self.assertEqual([display_cell(overall[index], index) for index in (3, 5, 7)],
                         ["14.00%", "20.00%", "50.00%"])
        self.assertEqual(overall_blast_row(self.rows(campaigns=("Small",)))[:4], [OVERALL_BLAST_LABEL, 100, 50, .5])
        self.assertEqual(overall_blast_row(self.rows(report_date=date(2026, 11, 19)))[1:8], [None] * 7)

    def test_missing_counts_zero_blasts_and_zero_active_do_not_create_false_percentages(self):
        totals = self.totals()
        small = totals[campaign_totals_key(self.report_date, "Small")]
        for value in (None, -1, True, "50"):
            with self.subTest(missing_blast=value):
                small["dc_email"] = value
                result = overall_blast_row(self.rows(totals))
                self.assertEqual(result[2], 90)
                self.assertIsNone(result[3])
                self.assertEqual(result[5], .2)
        small["dc_email"] = 0
        result = overall_blast_row(self.rows(totals))
        self.assertEqual(result[3], .09)
        small["active_ticket"] = None
        self.assertTrue(all(overall_blast_row(self.rows(totals))[index] is None for index in (3, 5, 7)))
        for values in totals.values():
            values.update(active_ticket=0, dc_email=0, dc_sms=0, dc_viber=0)
        self.assertEqual(overall_blast_row(self.rows(totals))[1:8], [0, 0, None, 0, None, 0, None])
        small["active_ticket"] = 100
        self.assertEqual(overall_blast_row(self.rows(totals))[1:8], [100, 0, 0, 0, 0, 0, 0])
        self.assertEqual(overall_blast_row([])[1:], [None] * 17)

    def test_html_footer_keeps_campaign_rows_and_shows_overall_channel_percentages(self):
        from inbound_reply_page import build_report_html
        from test_inbound_display import table_cells

        html = build_report_html("", self.report_date, "", campaign_totals=self.totals(),
                                 source_campaigns=("Small", "Large"))
        self.assertEqual([row[0] for row in table_cells(html)], ["Large", "Small"])
        footer = html.split('<tfoot aria-label="Overall blast coverage">')[1].split('</tfoot>')[0]
        cells = [cell.split('</td>')[0] for cell in footer.split('<td>')[1:]]
        self.assertEqual(cells[:8], [OVERALL_BLAST_LABEL, "1,000", "140", "14.00%", "200", "20.00%", "500", "50.00%"])

    def test_excel_footer_recalculates_from_campaign_cells_and_keeps_four_output_sheets(self):
        from openpyxl import load_workbook

        book = load_workbook(BytesIO(build_summary_workbook("", self.report_date, "", campaign_totals=self.totals(),
                                                          source_campaigns=("Small", "Large"))))
        try:
            self.assertEqual(book.sheetnames, OUTPUT_SHEETS)
            sheet = book["Summary"]
            self.assertEqual(sheet["A6"].value, OVERALL_BLAST_LABEL)
            self.assertEqual(sheet["B6"].value, '=IF(COUNT(B4:B5)=0,"",SUM(B4:B5))')
            self.assertEqual(sheet["C6"].value, '=IF(COUNT(C4:C5)=0,"",SUM(C4:C5))')
            for column, percent_column in (("C", "D"), ("E", "F"), ("G", "H")):
                cell = sheet[f"{percent_column}6"]
                self.assertEqual(cell.data_type, "f")
                self.assertIn(f'SUM({column}$4:{column}$5)/SUM($B$4:$B$5)', cell.value)
                self.assertIn(f'COUNT({column}$4:{column}$5)=ROWS({column}$4:{column}$5)', cell.value)
                self.assertIn('COUNT($B$4:$B$5)=ROWS($B$4:$B$5)', cell.value)
                self.assertEqual(cell.number_format, "0.00%")
                self.assertTrue(cell.font.bold)
            self.assertEqual(sheet.max_row, 6)
            self.assertEqual(sheet.print_area, "'Summary'!$A$1:$R$6")
            self.assertEqual(book.calculation.calcMode, "auto")
            self.assertTrue(book.calculation.fullCalcOnLoad)
            self.assertIsNone(sheet["I6"].value)
        finally:
            book.close()

    def test_page_edits_refresh_overall_percentages_and_downloaded_counts(self):
        from hashlib import sha256
        from openpyxl import load_workbook
        from streamlit.testing.v1 import AppTest

        with patch("branding.render_bank_branding"), patch("streamlit.file_uploader", return_value=None), \
                patch("inbound_reply_page.inbound_report_date", return_value=self.report_date), \
                patch("inbound_reply_page.prepare_campaign_labels", return_value=("Small", "Large")), \
                patch("streamlit.html") as html, patch("streamlit.download_button") as downloads:
            app = AppTest.from_file(str(ROOT / "pages" / "INBOUND_REPLY.py"))
            app.session_state["inbound_reply_campaign_totals"] = self.totals()
            app.run(timeout=20)

            def footer():
                self.assertFalse(app.exception)
                call = next(call for call in reversed(html.call_args_list)
                            if 'aria-label="Inbound reply summary"' in call.args[0])
                return call.args[0].split('<tfoot aria-label="Overall blast coverage">')[1].split('</tfoot>')[0]

            self.assertIn("14.00%", footer())
            prefix = "inbound_reply_total_" + sha256(campaign_totals_key(self.report_date, "Large").encode()).hexdigest()[:16]
            app.number_input(key=prefix + "_dc_email").set_value(450).run(timeout=20)
            self.assertNotIn("14.00%", footer())
            self.assertIn("50.00%", footer())
            call = next(call for call in reversed(downloads.call_args_list)
                        if call.kwargs.get("key") == "inbound_reply_summary_download")
            book = load_workbook(BytesIO(call.kwargs["data"]()))
            try:
                self.assertEqual(book["Summary"]["C4"].value, 450)
                self.assertEqual(book["Summary"]["A6"].value, OVERALL_BLAST_LABEL)
                self.assertEqual(book["Summary"]["D6"].data_type, "f")
            finally:
                book.close()


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

    def test_excel_summary_has_reference_merges_and_separate_channel_details(self):
        from openpyxl import load_workbook

        viber, email = reports()
        report_date = date(2026, 11, 12)
        totals = {campaign_totals_key(report_date, "A"): {
            "active_ticket": 100, "dc_email": 50, "dc_viber": 60,
        }}
        book = load_workbook(BytesIO(build_summary_workbook("", report_date, "YES", viber, email, totals)))
        try:
            self.assertEqual(book.sheetnames, OUTPUT_SHEETS)
            sheet = book["Summary"]
            self.assertEqual({str(area) for area in sheet.merged_cells.ranges},
                             {"A1:K1", "L1:L3", "M1:R2", "A2:A3", "B2:B3", "C2:H2", "I2:K2"})
            self.assertEqual(sheet.max_column, 18)
            self.assertEqual(sheet["A1"].value, "NOVEMBER 12, 2026 DATA")
            self.assertEqual(sheet["L1"].value, "SENT TO OPS")
            self.assertEqual(sheet["M1"].value, "AGENT WORKED")
            self.assertTrue(sheet["I3"].alignment.wrap_text)
            self.assertGreaterEqual(sheet.row_dimensions[3].height, 30)
            self.assertEqual(sheet["B4"].value, 100)
            self.assertEqual(sheet["B4"].data_type, "n")
            self.assertEqual(sheet["C4"].value, 50)
            self.assertEqual(sheet["D4"].value,
                             '=IF(AND(ISNUMBER($B4),$B4>0,ISNUMBER(C4),C4>=0),C4/$B4,"")')
            self.assertEqual(sheet["D4"].data_type, "f")
            self.assertEqual(sheet["D4"].number_format, "0.00%")
            self.assertEqual(sheet["I4"].value, 2)
            self.assertEqual(sheet["M4"].value, 1)
            for percentage, worked, replies in (("N4", "M4", "I4"), ("P4", "O4", "J4"), ("R4", "Q4", "K4")):
                self.assertEqual(sheet[percentage].value,
                                 f'=IF(AND(ISNUMBER({replies}),{replies}>0,ISNUMBER({worked}),{worked}>=0),'
                                 f'{worked}/{replies},"")')
                self.assertEqual(sheet[percentage].number_format, "0.00%")
            self.assertIsNone(sheet["J4"].value)
            self.assertIsNone(sheet["O4"].value)
            self.assertEqual(sheet.freeze_panes, "C4")
            rows = detail_rows(book["Email Unreplied"]) + detail_rows(book["Viber Unreplied"])
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
            details = book["Email Unreplied"]
            formula_column = [cell.value for cell in details[3]].index("Source: AI Suggestion") + 1
            self.assertEqual(details.cell(4, formula_column).data_type, "s")
            self.assertEqual(book["SMS Unreplied"]["B1"].value, "Not uploaded")
            self.assertIsNone(book["SMS Unreplied"]["B2"].value)
            self.assertEqual(detail_rows(book["SMS Unreplied"]), [])
        finally:
            book.close()

    def test_unreplied_message_columns_wrap_and_fit_multiline_text_in_both_downloads(self):
        from openpyxl import load_workbook

        body = ("This sample message needs more than one line in the worksheet. " * 4).strip() + "\nSecond paragraph."
        dc = read_email_source("DC.csv", csv_bytes(("sender", "Campaign", "body"), [
            ("pending@example.com", "SBC B2", body),
        ]), kind="dc")
        history = read_email_source("History.csv", csv_bytes(("email", "sent"), []), kind="history")
        report = match_email_sources(dc, history)
        exports = (
            (build_summary_workbook("", date(2026, 11, 19), "", email_report=report), "Email Unreplied", 3, "Source: body"),
            (build_email_unreplied_workbook(report), "Unreplied Details", 1, "body"),
        )
        for data, title, header_row, header in exports:
            with self.subTest(sheet=title):
                workbook = load_workbook(BytesIO(data))
                try:
                    sheet = workbook[title]
                    column = next(cell.column for cell in sheet[header_row] if cell.value == header)
                    message = sheet.cell(header_row + 1, column)
                    self.assertEqual(message.value, body)
                    self.assertTrue(message.alignment.wrap_text)
                    self.assertGreater(sheet.column_dimensions[message.column_letter].width, 75)
                    self.assertGreater(sheet.row_dimensions[message.row].height, 30)
                finally:
                    workbook.close()

    def test_email_unreplied_links_to_the_complete_text_of_overflowing_messages(self):
        from openpyxl import load_workbook

        body = "=Long message\n" + "A complete paragraph to display in the worksheet.\n" * 100
        dc = read_email_source("DC.csv", csv_bytes(("sender", "Campaign", "body"), [
            ("pending@example.com", "SBC B2", body),
        ]), kind="dc")
        history = read_email_source("History.csv", csv_bytes(("email", "sent"), []), kind="history")
        report = match_email_sources(dc, history)
        workbook = load_workbook(BytesIO(build_summary_workbook("", date(2026, 11, 19), "", email_report=report)))
        try:
            sheet = workbook["Email Unreplied"]
            source = sheet["I4"]
            self.assertEqual(sheet["B2"].value, 1)
            self.assertEqual(sheet.auto_filter.ref, "A3:I4")
            self.assertEqual(source.value, body)
            self.assertEqual(source.hyperlink.location, "'Email Unreplied'!A7")
            parts = [sheet.cell(row, 2) for row in range(7, sheet.max_row + 1)]
            self.assertGreater(len(parts), 1)
            self.assertEqual("".join(part.value for part in parts), body)
            for part in parts:
                self.assertTrue(part.alignment.wrap_text)
                self.assertEqual(part.data_type, "s")
                self.assertLessEqual(sheet.row_dimensions[part.row].height, 409)
                self.assertEqual(sheet.cell(part.row, 1).hyperlink.location, "'Email Unreplied'!I4")
        finally:
            workbook.close()

    def test_manual_only_report_keeps_missing_channels_blank(self):
        from openpyxl import load_workbook

        report_date = date(2026, 11, 12)
        totals = {campaign_totals_key(report_date, "Manual"): {"active_ticket": 0, "dc_email": 10}}
        book = load_workbook(BytesIO(build_summary_workbook("Manual", report_date, "NO", campaign_totals=totals)))
        try:
            self.assertEqual(book["Summary"]["B4"].value, 0)
            self.assertEqual(book["Summary"]["C4"].value, 10)
            self.assertEqual(book["Summary"]["D4"].data_type, "f")
            self.assertIsNone(book["Summary"]["I4"].value)
            for name in OUTPUT_SHEETS[1:]:
                self.assertEqual(book[name].max_row, 3)
                self.assertEqual(book[name]["B1"].value, "Not uploaded")
                self.assertIsNone(book[name]["B2"].value)
        finally:
            book.close()

    def test_blast_coverage_uses_active_total_and_excel_keeps_editable_formulas_for_every_row(self):
        from openpyxl import load_workbook

        report_date = date(2026, 11, 18)
        totals = {campaign_totals_key(report_date, "SBC Level 1"): {
            "active_ticket": 5000, "dc_email": 8, "dc_sms": 2500, "dc_viber": 5000,
        }}
        rows = summary_rows("", report_date, "", campaign_totals=totals, source_campaigns=("SBC Level 1", "SBC Level 2"))
        self.assertEqual([display_cell(rows[0][i], i) for i in (3, 5, 7)], ["0.16%", "50.00%", "100.00%"])
        book = load_workbook(BytesIO(build_summary_workbook("", report_date, "", campaign_totals=totals,
                                                          source_campaigns=("SBC Level 1", "SBC Level 2"))))
        try:
            self.assertEqual(book.calculation.calcMode, "auto")
            self.assertTrue(book.calculation.fullCalcOnLoad)
            for row in (4, 5):
                for numerator, percentage in (("C", "D"), ("E", "F"), ("G", "H")):
                    cell = book["Summary"][f"{percentage}{row}"]
                    self.assertEqual(cell.data_type, "f")
                    self.assertIn(f"{numerator}{row}/$B{row}", cell.value)
                    self.assertEqual(cell.number_format, "0.00%")
            self.assertIsNone(book["Summary"]["B5"].value)
            self.assertIsNone(book["Summary"]["C5"].value)
            self.assertEqual(book["Summary"]["B4"].value, 5000)
            self.assertEqual(book["Summary"]["C4"].value, 8)
        finally:
            book.close()

    def test_main_download_uses_grouped_summary_and_controls_are_visible(self):
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
                self.assertEqual(book.sheetnames, OUTPUT_SHEETS)
            finally:
                book.close()
            headings = {heading.value for heading in app.subheader}
            for label in ("Upload files", "Edit campaign totals", "Email account details", "Viber account details"):
                self.assertIn(label, headings)
            self.assertNotIn("Report settings", headings)
            self.assertEqual(len(app.text_input), 0)
            self.assertEqual(len(app.date_input), 0)
            self.assertEqual(len(app.selectbox), 0)
            self.assertFalse(any("CAMPAIGN is filled automatically" in item.value for item in app.caption))
            self.assertEqual(len(app.expander), 0)


if __name__ == "__main__":
    unittest.main()
