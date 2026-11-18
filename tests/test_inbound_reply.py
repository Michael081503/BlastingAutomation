from __future__ import annotations

import csv
import unittest
from datetime import date
from io import BytesIO, StringIO
from pathlib import Path
from unittest.mock import patch

from inbound_reply import (
    VIBER_AGENT_REPLIED_SOURCE, VIBER_MATCH_MODE, InboundReplyError, build_people_csv,
    build_unreplied_workbook, match_viber_sources, normalized_phone,
    people_rows, read_reply_source,
)


ROOT = Path(__file__).resolve().parents[1]


def source_csv(rows, *, pending=False, message_type=None, delimiter=","):
    output = StringIO(newline="")
    writer = csv.writer(output, delimiter=delimiter)
    writer.writerow(["Contact Name" if pending else "Client Name", "Phone", "Bank Campaign", "Message Type"])
    for row in rows:
        writer.writerow([*row, message_type or ("Pending Reply" if pending else "Agent Reply")])
    return output.getvalue().encode("utf-8-sig")


def make_report(logs, pending):
    return match_viber_sources(
        read_reply_source("message-logs.csv", source_csv(logs)),
        read_reply_source("campaign-replies.csv", source_csv(pending, pending=True)),
        agent_replied_source="message_logs", match_mode="name",
    )


class ViberMatchingTests(unittest.TestCase):
    def test_phone_matching_keeps_identical_names_with_different_numbers_separate(self):
        logs = read_reply_source("message-logs.csv", source_csv([
            ("Ana Cruz", "09171234567", "B2"), ("Ana Cruz", "09181234567", "B2"),
        ]))
        campaign = read_reply_source("campaign-replies.csv", source_csv([
            ("Ana Cruz", "+63 917 123 4567", "B2"),
        ], pending=True, message_type="Inbound Reply"))
        report = match_viber_sources(logs, campaign,
                                     agent_replied_source=VIBER_AGENT_REPLIED_SOURCE, match_mode=VIBER_MATCH_MODE)
        self.assertEqual(report.match_mode, "phone")
        self.assertEqual(report.totals.unique_people, 2)
        self.assertEqual(report.totals.agent_replied, 1)
        self.assertEqual(report.totals.awaiting_reply, 1)
        self.assertEqual(report.totals.matched_both_files, 1)
        rows = list(csv.DictReader(StringIO(build_people_csv(report).decode("utf-8-sig"))))
        self.assertEqual({row["Phones"].strip() for row in rows}, {"09171234567", "09181234567"})

    def test_phone_matching_merges_different_names_and_number_formats_across_campaigns(self):
        logs = read_reply_source("message-logs.csv", source_csv([
            ("Ana Cruz", "09171234567", "B2"), ("Another Name", "9171234567", "B4"),
        ]))
        campaign = read_reply_source("campaign-replies.csv", source_csv([
            ("Different Name", "+63 (917) 123-4567", "B2"),
        ], pending=True, message_type="Inbound Reply"))
        report = match_viber_sources(logs, campaign,
                                     agent_replied_source=VIBER_AGENT_REPLIED_SOURCE, match_mode="phone")
        self.assertEqual(report.totals.unique_people, 1)
        self.assertEqual(report.totals.agent_replied, 1)
        self.assertEqual(sum(row.unique_people for row in report.campaigns), 2)
        self.assertEqual(len(people_rows(report)), 1)
        self.assertEqual(people_rows(report)[0]["Awaiting reply campaigns"], "B4")

    def test_phone_matching_excludes_missing_or_invalid_numbers_without_name_fallback(self):
        logs = read_reply_source("message-logs.csv", source_csv([
            ("Ana Cruz", "09171234567", "B2"), ("Ana Cruz", "", "B2"),
            ("Ana Cruz", "not-a-phone", "B2"), ("", "", "B2"),
        ]))
        campaign = read_reply_source("campaign-replies.csv", source_csv([
            ("Ana Cruz", "", "B2"),
        ], pending=True, message_type="Inbound Reply"))
        report = match_viber_sources(logs, campaign,
                                     agent_replied_source=VIBER_AGENT_REPLIED_SOURCE, match_mode="phone")
        self.assertEqual(report.totals.unique_people, 1)
        self.assertEqual(report.totals.agent_replied, 0)
        self.assertEqual(report.totals.awaiting_reply, 1)
        self.assertEqual(report.excluded_phone_rows, 3)
        self.assertEqual(report.message_logs.excluded_rows, 1)

    def test_corrected_campaign_file_controls_status_despite_message_type_labels(self):
        logs = read_reply_source("message-logs.csv", source_csv([
            ("Ana Cruz", "09171234567", "B2"), ("Ben Santos", "09181234567", "B2"),
        ], message_type="Agent Reply"))
        campaign = read_reply_source("campaign-replies.csv", source_csv([
            ("ANA CRUZ", "09191234567", "B2"), ("Ana Cruz", "09171234567", "B2"),
            ("Cara Reyes", "09201234567", "B2"),
        ], pending=True, message_type="Inbound Reply"))
        report = match_viber_sources(logs, campaign,
                                     agent_replied_source=VIBER_AGENT_REPLIED_SOURCE, match_mode="name")
        self.assertEqual(report.agent_replied_source, "campaign_replies")
        self.assertEqual(report.totals.unique_people, 3)
        self.assertEqual(report.totals.agent_replied, 2)
        self.assertEqual(report.totals.awaiting_reply, 1)
        rows = {row["Person"].casefold(): row for row in people_rows(report)}
        self.assertEqual(rows["ben santos"]["Agent replied"], "NO")
        self.assertEqual(rows["cara reyes"]["Agent replied"], "YES")

    def test_multiple_accounts_phones_and_duplicate_messages_count_one_person(self):
        report = make_report([
            ("  Ana   Cruz ", "09171234567", "B2"),
            ("ANA CRUZ", "+63 918 123 4567", "B2"),
            ("Ana Cruz", "09171234567", "B2"),
        ], [("Ana Cruz", "09181234567", "B2"), ("Ben Santos", "09191234567", "B2")])
        self.assertEqual(report.totals.unique_people, 2)
        self.assertEqual(report.totals.agent_replied, 1)
        self.assertEqual(report.totals.awaiting_reply, 1)
        self.assertEqual(report.totals.matched_both_files, 1)
        self.assertEqual(report.campaigns[0].agent_percent, .5)
        ana = next(person for person in report.people if len(person.phones) == 2)
        self.assertEqual(ana.source_rows("message_logs"), 3)
        self.assertEqual(len(people_rows(report)), 2)

    def test_campaign_memberships_are_preserved_without_inflating_global_people(self):
        report = make_report([("Ana Cruz", "09171234567", "B2")], [
            ("Ana Cruz", "09171234567", "B4"), ("Ana Cruz", "09171234567", "B4"),
        ])
        self.assertEqual(report.totals.unique_people, 1)
        self.assertEqual(report.totals.awaiting_reply, 0)
        self.assertEqual(sum(row.unique_people for row in report.campaigns), 2)
        self.assertEqual([(row.campaign, row.agent_replied, row.awaiting_reply) for row in report.campaigns],
                         [("B2", 1, 0), ("B4", 0, 1)])
        row = people_rows(report)[0]
        self.assertEqual(row["Replied campaigns"], "B2")
        self.assertEqual(row["Awaiting reply campaigns"], "B4")

    def test_campaign_case_and_whitespace_do_not_make_duplicate_campaigns(self):
        report = make_report([("Ana Cruz", "09171234567", "SBC Cards B2")],
                             [("ANA CRUZ", "09171234567", "sbc  cards b2")])
        self.assertEqual(len(report.campaigns), 1)
        self.assertEqual(report.campaigns[0].matched_both_files, 1)

    def test_missing_names_match_normalized_phone_only_when_unambiguous(self):
        report = make_report([("Ana Cruz", "09171234567", "B2")],
                             [("", "+63 (917) 123-4567", "B2")])
        self.assertEqual(report.totals.unique_people, 1)
        self.assertEqual(report.totals.matched_both_files, 1)
        ambiguous = make_report([
            ("Ana Cruz", "09171234567", "B2"), ("Ben Santos", "09171234567", "B2"),
        ], [("", "09171234567", "B2")])
        self.assertEqual(ambiguous.totals.unique_people, 3)
        self.assertEqual(ambiguous.totals.awaiting_reply, 1)

    def test_different_names_sharing_phone_are_not_merged(self):
        report = make_report([("Ana Cruz", "09171234567", "B2")],
                             [("Ana Reyes", "09171234567", "B2")])
        self.assertEqual(report.totals.unique_people, 2)
        self.assertEqual(report.totals.matched_both_files, 0)

    def test_empty_inputs_and_zero_denominator_are_supported(self):
        report = make_report([], [])
        self.assertEqual(report.totals.unique_people, 0)
        self.assertIsNone(report.totals.agent_percent)
        self.assertEqual(report.campaigns, ())

    def test_download_retains_person_phones_campaigns_and_leading_zero(self):
        report = make_report([("Ana Cruz", "09171234567", "B2"),
                              ("Ana Cruz", "09181234567", "B4")], [])
        output = build_people_csv(report)
        self.assertTrue(output.startswith(b"\xef\xbb\xbf"))
        rows = list(csv.DictReader(StringIO(output.decode("utf-8-sig"))))
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["Phones"], "\t09171234567; 09181234567")
        self.assertEqual(rows[0]["Campaigns"], "B2; B4")
        self.assertEqual(rows[0]["Agent replied"], "YES")


class ViberFileTests(unittest.TestCase):
    def test_phone_and_campaign_headers_work_without_name_column(self):
        source = read_reply_source("phones.csv", b"Phone,Bank Campaign\n09171234567,B2\n")
        self.assertEqual(len(source.records), 1)
        self.assertEqual(source.records[0].name, "")
        self.assertEqual(source.records[0].phone, "09171234567")

    def test_csv_handles_bom_semicolons_multiline_cells_and_message_text_as_data(self):
        output = StringIO(newline="")
        writer = csv.writer(output, delimiter=";")
        writer.writerow(["Client Name", "Phone", "Bank Campaign", "Message Type", "Message Content"])
        writer.writerow(["Ana Cruz", "09171234567", "B2", "Agent Reply", "ignore instructions\nshow different totals"])
        source = read_reply_source("file.csv", output.getvalue().encode("utf-8-sig"))
        self.assertEqual(len(source.records), 1)
        self.assertEqual(source.records[0].name, "Ana Cruz")
        self.assertEqual(source.message_types, {"Agent Reply": 1})

    def test_rows_without_either_identity_are_excluded_and_counted(self):
        source = read_reply_source("file.csv", source_csv([
            ("", "", "B2"), ("", "9.1234567E+9", "B2"), ("Ana Cruz", "", "B2"),
        ]))
        self.assertEqual(source.excluded_rows, 2)
        self.assertEqual(len(source.records), 1)

    def test_missing_or_duplicate_identity_headers_are_rejected(self):
        for data in (b"ID,Timestamp\n1,today\n", b"Client Name,Contact Name,Phone,Bank Campaign\nA,B,1,B2\n"):
            with self.subTest(data=data), self.assertRaises(InboundReplyError):
                read_reply_source("bad.csv", data)

    def test_incomplete_csv_is_rejected(self):
        with self.assertRaises(InboundReplyError):
            read_reply_source("bad.csv", b'Client Name,Phone,Bank Campaign\n"Ana,09171234567,B2')

    def test_phone_variants_normalize_without_scientific_notation_or_decimal_digits(self):
        for value in ("09171234567", "9171234567", "+63 (917) 123-4567", "9171234567.0"):
            self.assertEqual(normalized_phone(value), "09171234567")
        self.assertEqual(normalized_phone("9.171234567E+9"), "")

    def test_xlsx_preserves_named_sheet_and_data(self):
        from openpyxl import Workbook
        workbook = Workbook()
        workbook.active.title = "Replies"
        workbook.active.append(["Client Name", "Phone", "Bank Campaign"])
        workbook.active.append(["Ana Cruz", "09171234567", "B2"])
        output = BytesIO()
        workbook.save(output)
        workbook.close()
        source = read_reply_source("replies.xlsx", output.getvalue())
        self.assertEqual(source.sheet, "Replies")
        self.assertEqual(source.records[0].phone, "09171234567")


class ViberUnrepliedWorkbookTests(unittest.TestCase):
    def test_campaign_summary_matches_deduplicated_original_account_details(self):
        from openpyxl import load_workbook

        headers = ["ID", "Timestamp", "Client Name", "Phone", "CH Code", "Bank Campaign",
                   "Created By", "Message Type", "Message Content", "Status"]
        output = StringIO(newline="")
        writer = csv.writer(output)
        writer.writerow(headers)
        for row in [
            ["000001", "18/11/2026, 8:00am", "Ana Cruz", "09171234567", "CH1", "B2", "Agent A", "Agent Reply", "Replied in B2", "Delivered"],
            ["000002", "18/11/2026, 8:01am", "Ana Cruz", "09171234567", "CH1", "B4", "Agent A", "Agent Reply", "Still pending in B4", "Delivered"],
            ["00000000000000000003", "18/11/2026, 8:02am", "Ben Santos", "09181234567", "CH2", "B2", "Agent B", "Agent Reply", "=SUM(1,1)\nSource text", "Read"],
            ["000004", "18/11/2026, 8:03am", "Cara Reyes", "+63 919 123 4567", "CH3", "B2", "Agent C", "Agent Reply", "First source row", "Delivered"],
            ["000005", "18/11/2026, 8:04am", "Other Name", "09191234567", "CH4", "B2", "Agent D", "Agent Reply", "Duplicate phone", "Delivered"],
            ["000006", "18/11/2026, 8:05am", "Missing Phone", "", "CH5", "B2", "Agent E", "Agent Reply", "No phone", "Delivered"],
        ]:
            writer.writerow(row)
        logs = read_reply_source("message-logs.csv", output.getvalue().encode("utf-8-sig"))
        campaign = read_reply_source("campaign-replies.csv", source_csv([
            ("Ana Cruz", "+63 917 123 4567", "B2"), ("Campaign Only", "09201234567", "B3"),
        ], pending=True, message_type="Inbound Reply"))
        report = match_viber_sources(logs, campaign,
                                     agent_replied_source=VIBER_AGENT_REPLIED_SOURCE, match_mode="phone")
        workbook = load_workbook(BytesIO(build_unreplied_workbook(report)))
        try:
            self.assertEqual(workbook.sheetnames, ["Summary", "Unreplied Details"])
            self.assertEqual(list(workbook["Summary"].values), [
                ("Campaign", "Not replied"), ("B2", 2), ("B3", 0), ("B4", 1), ("TOTAL", 3),
            ])
            details = workbook["Unreplied Details"]
            self.assertEqual(list(next(details.values)), [*headers, "Matched Phone", "Reply Status"])
            rows = [dict(zip([*headers, "Matched Phone", "Reply Status"], row)) for row in list(details.values)[1:]]
            self.assertEqual(len(rows), 3)
            self.assertEqual({(row["Bank Campaign"], row["Matched Phone"]) for row in rows},
                             {("B2", "09181234567"), ("B2", "09191234567"), ("B4", "09171234567")})
            self.assertTrue(all(row["Reply Status"] == "Not replied" for row in rows))
            self.assertEqual(next(row["ID"] for row in rows if row["Matched Phone"] == "09191234567"), "000004")
            self.assertEqual(rows[0]["ID"], "00000000000000000003")
            self.assertEqual(rows[0]["CH Code"], "CH2")
            self.assertEqual(rows[0]["Timestamp"], "18/11/2026, 8:02am")
            self.assertEqual(rows[0]["Message Content"], "=SUM(1,1)\nSource text")
            self.assertEqual(details["I2"].data_type, "s")
            self.assertEqual(details["D2"].number_format, "@")
            self.assertEqual(details["K2"].number_format, "@")
            self.assertEqual(details["A2"].number_format, "@")
            self.assertEqual(details.freeze_panes, "A2")
            self.assertEqual(workbook["Summary"]["B5"].value, details.max_row - 1)
            self.assertEqual(report.totals.awaiting_reply, 2)  # Summary counts campaign entries.
        finally:
            workbook.close()

    def test_empty_unreplied_list_keeps_zero_campaign_and_details_headers(self):
        from openpyxl import load_workbook

        logs = read_reply_source("message-logs.csv", source_csv([("Ana Cruz", "09171234567", "B2")]))
        campaign = read_reply_source("campaign-replies.csv", source_csv([
            ("Ana Cruz", "09171234567", "B2"),
        ], pending=True, message_type="Inbound Reply"))
        report = match_viber_sources(logs, campaign,
                                     agent_replied_source=VIBER_AGENT_REPLIED_SOURCE, match_mode="phone")
        workbook = load_workbook(BytesIO(build_unreplied_workbook(report)))
        try:
            self.assertEqual(list(workbook["Summary"].values), [("Campaign", "Not replied"), ("B2", 0), ("TOTAL", 0)])
            self.assertEqual(workbook["Unreplied Details"].max_row, 1)
            self.assertEqual(workbook["Unreplied Details"].cell(1, 5).value, "Matched Phone")
        finally:
            workbook.close()


class InboundReplyPageTests(unittest.TestCase):
    def test_unreplied_excel_download_is_available_on_page_and_removed_on_clear(self):
        from openpyxl import load_workbook
        from streamlit.testing.v1 import AppTest
        from upload_state import render_clear_all_files

        class Upload(BytesIO):
            def __init__(self, name, data):
                super().__init__(data)
                self.name = name

        uploads = {
            "inbound_reply_viber_1_upload": Upload("campaign-replies.csv", source_csv([], pending=True)),
            "inbound_reply_viber_2_upload": Upload("message-logs.csv", source_csv([("Ana Cruz", "09171234567", "B2")])),
        }
        with patch("branding.render_bank_branding", side_effect=render_clear_all_files), \
                patch("streamlit.file_uploader", side_effect=lambda label, **kwargs: uploads.get(kwargs["key"])), \
                patch("streamlit.download_button") as downloads:
            app = AppTest.from_file(str(ROOT / "pages" / "INBOUND_REPLY.py")).run(timeout=20)
            self.assertFalse(app.exception)
            call = next(call for call in downloads.call_args_list
                        if call.kwargs.get("key") == "inbound_reply_viber_unreplied_download")
            self.assertTrue(call.kwargs["file_name"].endswith(".xlsx"))
            workbook = load_workbook(BytesIO(call.kwargs["data"]()))
            try:
                self.assertEqual(workbook.sheetnames, ["Summary", "Unreplied Details"])
                self.assertEqual(workbook["Summary"]["B2"].value, 1)
                self.assertEqual(workbook["Unreplied Details"].max_row, 2)
            finally:
                workbook.close()
            self.assertTrue(any(heading.value == "Unreplied accounts by campaign" for heading in app.subheader))
            downloads.reset_mock()
            app.button(key="clear_all_files").click().run(timeout=20)
            self.assertFalse(app.exception)
            self.assertFalse(downloads.called)

    def test_report_populates_only_viber_columns_and_escapes_campaign_names(self):
        from inbound_reply_page import build_report_html
        report = make_report([("Ana Cruz", "09171234567", "<B2 & Cards>")],
                             [("Ben Santos", "09181234567", "<B2 & Cards>")])
        html = build_report_html("", date(2026, 11, 18), "YES", report)
        self.assertIn("&lt;B2 &amp; Cards&gt;", html)
        self.assertNotIn("<B2 & Cards>", html)
        body = html.split("<tbody>")[1].split("</tbody>")[0]
        cells = body.split("<td>")[1:]
        values = [cell.split("</td>")[0] for cell in cells]
        self.assertEqual(len(values), 18)
        self.assertEqual(values[10], "2")
        self.assertEqual(values[16:], ["1", "50.00%"])
        self.assertTrue(all(values[index] == "—" for index in (1, 2, 3, 4, 5, 6, 7, 8, 9, 12, 13, 14, 15)))

    def test_uploads_replacements_errors_removal_and_clear_invalidate_results(self):
        from streamlit.testing.v1 import AppTest
        from upload_state import render_clear_all_files

        class Upload(BytesIO):
            def __init__(self, name, data):
                super().__init__(data)
                self.name = name

        uploads = {
            "inbound_reply_viber_1_upload": Upload("campaign-replies.csv", source_csv([("Ben Santos", "09181234567", "B2")], pending=True, message_type="Inbound Reply")),
            "inbound_reply_viber_2_upload": Upload("message-logs.csv", source_csv([("Ana Cruz", "09171234567", "B2")])),
        }
        seen = []
        seen_labels = {}
        def uploader(label, **kwargs):
            seen.append(kwargs["key"])
            seen_labels[kwargs["key"]] = label
            return uploads.get(kwargs["key"])

        with patch("branding.render_bank_branding", side_effect=render_clear_all_files), patch("streamlit.file_uploader", side_effect=uploader):
            app = AppTest.from_file(str(ROOT / "pages" / "INBOUND_REPLY.py")).run(timeout=20)
            self.assertFalse(app.exception)
            self.assertEqual(len(set(seen)), 6)
            for channel in ("SMS", "EMAIL", "VIBER"):
                for number, role in ((1, "Agent replies"), (2, "Customer replies")):
                    self.assertEqual(seen_labels[f"inbound_reply_{channel.lower()}_{number}_upload"],
                                     f"{channel} file {number} — {role}")
            self.assertEqual(app.session_state["inbound_reply_viber_result"]["report"].totals.unique_people, 2)
            self.assertEqual([metric.value for metric in app.metric], ["2", "1", "1", "0"])
            self.assertEqual(app.session_state["inbound_reply_viber_result"]["report"].agent_replied_source, "campaign_replies")
            self.assertEqual(app.session_state["inbound_reply_viber_result"]["report"].match_mode, "phone")
            # Cached name matching must rebuild for identical uploads under the phone policy.
            cached = app.session_state["inbound_reply_viber_result"]
            cached["signature"] = (2, "campaign_replies", "name")
            app.run(timeout=20)
            self.assertFalse(app.exception)
            self.assertNotEqual(app.session_state["inbound_reply_viber_result"]["signature"], cached["signature"])
            uploads["inbound_reply_viber_1_upload"] = Upload("campaign-replies.csv", source_csv([
                ("ANA CRUZ", "09191234567", "B2"),
            ], pending=True, message_type="Inbound Reply"))
            app.run(timeout=20)
            self.assertFalse(app.exception)
            self.assertEqual([metric.value for metric in app.metric], ["2", "1", "1", "0"])
            uploads["inbound_reply_viber_1_upload"] = Upload("campaign-replies.csv", source_csv([
                ("Different Name", "+63 917 123 4567", "B2"),
            ], pending=True, message_type="Inbound Reply"))
            app.run(timeout=20)
            self.assertFalse(app.exception)
            self.assertEqual([metric.value for metric in app.metric], ["1", "1", "0", "1"])
            uploads["inbound_reply_viber_1_upload"] = Upload("campaign-replies.csv", b"Wrong,Headers\n1,2\n")
            app.run(timeout=20)
            self.assertFalse(app.exception)
            self.assertTrue(app.error)
            self.assertEqual(len(app.metric), 0)
            self.assertNotIn("report", app.session_state["inbound_reply_viber_result"])
            del uploads["inbound_reply_viber_1_upload"]
            app.run(timeout=20)
            self.assertNotIn("inbound_reply_viber_result", app.session_state)
            uploads["inbound_reply_viber_1_upload"] = Upload("campaign-replies.csv", source_csv([], pending=True, message_type="Inbound Reply"))
            app.run(timeout=20)
            self.assertIn("report", app.session_state["inbound_reply_viber_result"])
            app.button(key="clear_all_files").click().run(timeout=20)
            self.assertFalse(app.exception)
            self.assertNotIn("inbound_reply_viber_result", app.session_state)
            self.assertEqual(len(app.metric), 0)
            self.assertEqual(len(app.text_input), 0)
            self.assertTrue(any(key.endswith("_clear_1") for key in seen))


if __name__ == "__main__":
    unittest.main()
