from __future__ import annotations

import unittest
from datetime import date, timedelta
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from inbound_reply import CampaignReplyCounts
from inbound_reply_page import build_report_html
from inbound_summary import campaign_totals_key


ROOT = Path(__file__).resolve().parents[1]


def table_cells(html):
    body = html.split("<tbody>")[1].split("</tbody>")[0]
    return [[cell.split("</td>")[0] for cell in row.split("<td>")[1:]]
            for row in body.split("<tr>")[1:]]


class ManualCampaignTotalsTests(unittest.TestCase):
    def test_entered_counts_use_active_ticket_for_dc_percentages_and_replies_for_agents(self):
        report_date = date(2026, 11, 12)
        email = SimpleNamespace(campaigns=(CampaignReplyCounts("Sample Campaign", 51, 20, 31, 20),))
        viber = SimpleNamespace(campaigns=(CampaignReplyCounts("Sample Campaign", 321, 10, 311, 10),))
        totals = {campaign_totals_key(report_date, "sample  campaign"): {
            "active_ticket": 58431, "dc_email": 46104, "dc_sms": 58412, "dc_viber": 1000,
        }}
        html = build_report_html("", report_date, "YES", viber, email, totals)
        cells = table_cells(html)[0]
        self.assertEqual(cells[1:8], ["58,431", "46,104", "78.90%", "58,412", "99.97%", "1,000", "1.71%"])
        self.assertEqual(cells[8:11], ["51", "—", "321"])
        self.assertEqual(cells[12:14], ["20", "39.22%"])
        self.assertEqual(cells[16:18], ["10", "3.12%"])
        self.assertIn('colspan="6" scope="colgroup">DC WORKED', html)
        totals[campaign_totals_key(report_date, "Sample Campaign")]["active_ticket"] *= 2
        cells = table_cells(build_report_html("", report_date, "YES", viber, email, totals))[0]
        self.assertEqual(cells[3], "39.45%")
        self.assertEqual(cells[13], "39.22%")

    def test_blank_or_zero_active_ticket_keeps_dc_percentages_blank(self):
        report_date = date(2026, 11, 12)
        for active, display in ((None, "—"), (0, "0")):
            with self.subTest(active=active):
                totals = {campaign_totals_key(report_date, "Sample"): {"active_ticket": active, "dc_email": 10}}
                cells = table_cells(build_report_html("Sample", report_date, "NO", campaign_totals=totals))[0]
                self.assertEqual(cells[1:4], [display, "10", "—"])
                self.assertEqual(cells[4:8], ["—", "—", "—", "—"])
        totals = {campaign_totals_key(report_date, "Sample"): {"active_ticket": 100, "dc_email": 0}}
        cells = table_cells(build_report_html("Sample", report_date, "NO", campaign_totals=totals))[0]
        self.assertEqual(cells[2:4], ["0", "0.00%"])

    def test_manual_counts_do_not_leak_to_a_different_date_or_campaign(self):
        report_date = date(2026, 11, 12)
        totals = {campaign_totals_key(report_date, "A"): {"active_ticket": 100, "dc_viber": 20}}
        cells = table_cells(build_report_html("B", report_date, "NO", campaign_totals=totals))[0]
        self.assertEqual(cells[1], "—")
        cells = table_cells(build_report_html("A", report_date + timedelta(days=1), "NO", campaign_totals=totals))[0]
        self.assertEqual(cells[1], "—")
        cells = table_cells(build_report_html(" a ", report_date, "NO", campaign_totals=totals))[0]
        self.assertEqual(cells[1], "100")
        self.assertEqual(cells[6:8], ["20", "20.00%"])


class ManualTotalsPageTests(unittest.TestCase):
    def test_page_uses_saved_totals_without_showing_input_controls(self):
        from streamlit.testing.v1 import AppTest
        from upload_state import render_clear_all_files

        report_date = date(2026, 11, 18)
        totals = {campaign_totals_key(report_date, "Sample Campaign"): {
            "active_ticket": 100, "dc_email": 25,
        }}
        with patch("branding.render_bank_branding", side_effect=render_clear_all_files), \
                patch("streamlit.file_uploader", return_value=None), patch("streamlit.html") as html, \
                patch("inbound_reply_page.inbound_report_date", return_value=report_date), \
                patch("inbound_reply_page.prepare_campaign_inputs", return_value=(("Sample Campaign",), {})):
            app = AppTest.from_file(str(ROOT / "pages" / "INBOUND_REPLY.py"))
            app.session_state["inbound_reply_campaign_totals"] = totals
            app.run(timeout=20)
            self.assertFalse(app.exception)
            self.assertEqual(len(app.text_input), 0)
            self.assertEqual(len(app.date_input), 0)
            self.assertEqual(len(app.number_input), 0)
            self.assertEqual(table_cells(html.call_args.args[0])[0][1:4], ["100", "25", "25.00%"])
            app.button(key="clear_all_files").click().run(timeout=20)
            self.assertEqual(table_cells(html.call_args.args[0])[0][1:4], ["100", "25", "25.00%"])
            self.assertEqual(len(app.number_input), 0)


if __name__ == "__main__":
    unittest.main()
