from __future__ import annotations

import csv
import unittest
from io import BytesIO, StringIO
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook

from inbound_campaign_lookup import build_email_campaign_lookup, read_email_campaign_source
from test_inbound_all_channels import Upload, all_channel_uploads
from test_inbound_campaign_lookup import workbook_bytes
from test_inbound_display import table_cells
from test_inbound_email import csv_bytes


ROOT = Path(__file__).resolve().parents[1]
BUCKET_HEADERS = ("Email (01)", "Email (02)", "Email (03)", "Email (04)", "Email (05)", "Bucket Code", "Area", "Mobile 1")


def bucket_files():
    return [
        ("B2.xlsx", workbook_bytes([
            ("SBC", [BUCKET_HEADERS,
                     ("b2@example.com", "secondary@example.com", None, None, None, 2, "Mindanao", "09301234567"),
                     (" SHARED@example.com ", None, None, None, None, 2, "NCR", "09311234567")]),
            ("SBC Lite", [("Account Key", "Name"), ("0001", "Example")]),
        ])),
        ("B4.xlsx", workbook_bytes([
            ("Sheet1", [BUCKET_HEADERS,
                        ("b4@example.com", None, None, None, None, 4, "Mindanao", "09321234567"),
                        ("six@example.com", None, None, None, None, 6, "Region 3", "09331234567"),
                        ("seventy@example.com", None, None, None, "ignored@example.com", 70, "NCR", "09341234567"),
                        ("shared@example.com", None, None, None, None, 4, "NCR", "+63 931 123 4567")]),
        ])),
    ]


def append_csv(upload, extra_rows):
    rows = list(csv.reader(StringIO(upload.getvalue().decode("utf-8-sig"), newline="")))
    return Upload(upload.name, csv_bytes(rows[0], [*rows[1:], *extra_rows]))


class BucketLookupTests(unittest.TestCase):
    def test_bucket_account_tables_are_preferred_over_reference_sheets(self):
        data = workbook_bytes([
            ("SBC", [("Account Key", "Email (01)", "Bucket Code", "Area"),
                     ("0001", "customer@example.com", 2, "NCR")]),
            ("IDENTIFIER", [(None, None), ("Email (01)", "Field reference"),
                            ("example@example.com", "Example value")]),
            ("SBC Lite", [("Account Key", "Name"), ("0001", "Example")]),
        ])
        source = read_email_campaign_source("B2.xlsx", data)
        self.assertEqual(source.sheets, ("SBC",))
        self.assertEqual(source.campaigns_by_email, {"customer@example.com": ("SBC B2",)})

    def test_file_membership_overrides_area_and_mixed_buckets_with_b2_all_emails_and_b4_primary_only(self):
        files = bucket_files()
        b2 = read_email_campaign_source(*files[0])
        b4 = read_email_campaign_source(*files[1])
        self.assertEqual(b2.sheets, ("SBC",))
        self.assertEqual(b4.sheets, ("Sheet1",))
        self.assertEqual(b2.campaign_label, "SBC B2")
        self.assertEqual(b4.campaign_label, "SBC B4")
        self.assertEqual(b2.campaigns_by_email, {
            "b2@example.com": ("SBC B2",), "secondary@example.com": ("SBC B2",), "shared@example.com": ("SBC B2",),
        })
        self.assertEqual(b4.campaigns_by_email, {address: ("SBC B4",) for address in (
            "b4@example.com", "six@example.com", "seventy@example.com", "shared@example.com",
        )})
        self.assertEqual(b4.missing_campaign_rows, 0)
        self.assertNotIn("ignored@example.com", b4.campaigns_by_email)

    def test_b4_reads_only_email_01_without_falling_back_for_blank_or_invalid_primary_cells(self):
        headers = ("Email (01)", "Email (02)", "Email (03)", "Email (04)", "Email (05)", "Mobile 1")
        rows = [
            (" PRIMARY@example.com ", "second@example.com", "third@example.com", "fourth@example.com", "fifth@example.com", "09301234567"),
            (None, "secondaryonly@example.com", None, None, None, "09311234567"),
            ("invalid", "fallback@example.com", None, None, None, "09321234567"),
        ]
        for extension in ("csv", "xlsx"):
            with self.subTest(extension=extension):
                data = csv_bytes(headers, rows) if extension == "csv" else workbook_bytes([("Sheet1", [headers, *rows])])
                source = read_email_campaign_source(f"B4.{extension}", data)
                self.assertEqual(source.campaigns_by_email, {"primary@example.com": ("SBC B4",)})
                self.assertEqual(source.invalid_email_rows, 2)
                self.assertEqual(source.campaigns_by_phone, {phone: ("SBC B4",) for phone in (
                    "09301234567", "09311234567", "09321234567",
                )})

    def test_renamed_bucket_file_uses_primary_email_for_b4_and_retains_b2_secondary_emails(self):
        source = read_email_campaign_source("renamed.csv", csv_bytes(("Email (01)", "Email (02)", "Bucket Code"), [
            ("b2@example.com", "b2second@example.com", 2),
            ("b4@example.com", "b4second@example.com", 4),
        ]))
        self.assertEqual(source.campaigns_by_email, {
            "b2@example.com": ("SBC B2",), "b2second@example.com": ("SBC B2",), "b4@example.com": ("SBC B4",),
        })

    def test_b2_precedes_b4_independent_of_upload_order_and_preserves_other_campaign_memberships(self):
        files = bucket_files() + [("regular.csv", csv_bytes(("email", "LEVEL"), [("shared@example.com", "Level 1")]))]
        forward = build_email_campaign_lookup(files)
        reverse = build_email_campaign_lookup(list(reversed(files)))
        self.assertEqual(forward.campaigns_by_email, reverse.campaigns_by_email)
        self.assertEqual(forward.campaigns_by_phone, reverse.campaigns_by_phone)
        self.assertEqual(forward.campaigns_by_email["shared@example.com"], ("SBC B2", "SBC Level 1"))
        self.assertEqual(forward.campaigns_by_phone["09311234567"], ("SBC B2",))
        self.assertEqual(forward.campaigns_by_email["six@example.com"], ("SBC B4",))
        b4_only = build_email_campaign_lookup([bucket_files()[1]])
        self.assertEqual(b4_only.campaigns_by_email["shared@example.com"], ("SBC B4",))
        copied = read_email_campaign_source("SBC-B4 (1).xlsx", bucket_files()[1][1])
        self.assertEqual(copied.campaigns_by_email, b4_only.sources[0].campaigns_by_email)

    def test_bucket_code_identifies_renamed_files_without_using_area(self):
        source = read_email_campaign_source("renamed.csv", csv_bytes(("Email (01)", "Bucket Code", "Area"), [
            ("b2@example.com", 2, "Mindanao"), ("b4@example.com", "B4", "Region 3"),
            ("unknown@example.com", 6, "NCR"),
        ]))
        self.assertEqual(source.campaigns_by_email, {"b2@example.com": ("SBC B2",), "b4@example.com": ("SBC B4",)})
        self.assertEqual(source.missing_campaign_rows, 1)


class BucketLookupPageTests(unittest.TestCase):
    def test_all_four_lookup_files_populate_campaigns_counts_exports_and_refresh_when_b2_is_removed(self):
        from streamlit.testing.v1 import AppTest

        uploads = all_channel_uploads()
        uploads["inbound_reply_email_lookup_uploads"].extend(Upload(name, data) for name, data in bucket_files())
        addresses = ("b2", "secondary", "shared", "b4", "six", "seventy")
        uploads["inbound_reply_email_2_upload"] = append_csv(uploads["inbound_reply_email_2_upload"], [
            (f"BUCKET-{i}", f"{name}@example.com", "Original Campaign", "Respond Now") for i, name in enumerate(addresses)
        ])
        history = append_csv(uploads.pop("inbound_reply_email_1_upload"), [
            ("b2@example.com", "Agent response", None), ("b4@example.com", "Agent response", None),
        ])
        for channel in ("sms", "viber"):
            key = f"inbound_reply_{channel}_2_upload"
            uploads[key] = append_csv(uploads[key], [
                (f"BUCKET-{i}", "Customer", f"093{i}1234567", "Original Campaign", "Customer message") for i in range(5)
            ])
            key = f"inbound_reply_{channel}_1_upload"
            uploads[key] = append_csv(uploads[key], [
                ("BUCKET-0", "Customer", "09301234567", "Original Campaign", "Agent response"),
                ("BUCKET-2", "Customer", "09321234567", "Original Campaign", "Agent response"),
            ])
        with patch("branding.render_bank_branding"), \
                patch("inbound_reply_page.SAVED_CAMPAIGN_LOOKUP_PATH", ROOT / "tests" / "missing-lookup.zip"), \
                patch("streamlit.file_uploader", side_effect=lambda label, **kw: uploads.get(kw["key"])), \
                patch("streamlit.html") as html, patch("streamlit.download_button") as downloads:
            app = AppTest.from_file(str(ROOT / "pages" / "INBOUND_REPLY.py")).run(timeout=20)

            def displayed():
                self.assertFalse(app.exception)
                call = next(call for call in reversed(html.call_args_list) if 'aria-label="Inbound reply summary"' in call.args[0])
                return table_cells(call.args[0])

            self.assertEqual([row[0] for row in displayed()][:2], ["SBC B2", "SBC B4"])
            self.assertEqual(displayed()[0][8:11], ["3", "2", "2"])
            self.assertEqual(displayed()[0][12:14], ["—", "—"])
            uploads["inbound_reply_email_1_upload"] = history
            app.run(timeout=20)
            self.assertFalse(app.error)
            self.assertEqual(displayed()[0][8:11], ["3", "2", "2"])
            self.assertEqual(displayed()[0][12:], ["1", "33.33%", "1", "50.00%", "1", "50.00%"])
            self.assertEqual(displayed()[1][8:11], ["3", "3", "3"])
            call = next(call for call in reversed(downloads.call_args_list) if call.kwargs.get("key") == "inbound_reply_summary_download")
            book = load_workbook(BytesIO(call.kwargs["data"]()))
            try:
                self.assertEqual(book["Summary"]["A4"].value, "SBC B2")
                self.assertEqual(book["Summary"]["A5"].value, "SBC B4")
                self.assertEqual(book["Summary"]["I4"].value, 3)
                for title, count in (("Email Unreplied", 10), ("SMS Unreplied", 10), ("Viber Unreplied", 9)):
                    self.assertEqual(book[title]["B2"].value, count)
                    self.assertEqual(book[title].max_row - 3, count)
                    self.assertTrue(all(book[title].cell(row, 2).value in [row[0] for row in displayed()]
                                        for row in range(4, book[title].max_row + 1)))
                rows = list(book["Email Unreplied"].iter_rows(min_row=4, values_only=True))
                self.assertEqual({row[1] for row in rows if row[4] == "shared@example.com"}, {"SBC B2"})
                self.assertEqual({row[1] for row in rows if row[4] in {"six@example.com", "seventy@example.com"}}, {"SBC B4"})
            finally:
                book.close()
            uploads["inbound_reply_email_lookup_uploads"] = [u for u in uploads["inbound_reply_email_lookup_uploads"] if u.name != "B2.xlsx"]
            app.run(timeout=20)
            self.assertFalse(app.error)
            self.assertNotIn("SBC B2", [row[0] for row in displayed()])
            b4 = next(row for row in displayed() if row[0] == "SBC B4")
            self.assertEqual(b4[8:11], ["4", "4", "4"])


if __name__ == "__main__":
    unittest.main()
