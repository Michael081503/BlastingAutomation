from __future__ import annotations

import unittest
from io import BytesIO
from pathlib import Path
from tempfile import NamedTemporaryFile
from unittest.mock import patch

from openpyxl import load_workbook

from inbound_campaign_lookup import save_campaign_lookup_files
from test_inbound_all_channels import Upload
from test_inbound_campaign_lookup import workbook_bytes
from test_inbound_display import table_cells
from test_inbound_email import csv_bytes


ROOT = Path(__file__).resolve().parents[1]
CAMPAIGNS = ["PL RECOVERY", "SBC B2", "SBC B4", "SBC Level 1", "SBC Level 6", "SBC Mindanao"]


def recovery_uploads():
    databases = [
        ("Regular.xlsx", [
            ("debtor_email", "LEVEL", "phone_1", "phone_2"),
            ("level1@example.com", 1, None, "09171234567"),
            ("level6@example.com", 6, "09181234567", None),
        ]),
        ("GEO.xlsx", [
            ("email", "placement", "Area", "Phone 1"),
            ("region3@example.com", 6, "Region 3", "09191234567"),
            ("mindanao@example.com", 1, "Mindanao", "09201234567"),
        ]),
        ("B2.xlsx", [
            ("Email (01)", "Email (02)", "Mobile 1"),
            ("primary@example.com", "b2@example.com", "09211234567"),
            ("shared@example.com", None, "09231234567"),
        ]),
        ("B4.xlsx", [
            ("Email (01)", "Email (02)", "Mobile 1"),
            ("b4@example.com", "ignored@example.com", "09221234567"),
            ("shared@example.com", None, "+63 923 123 4567"),
        ]),
    ]
    contacts = [
        ("level1", "09171234567"), ("level6", "09181234567"),
        ("region3", "09191234567"), ("mindanao", "09201234567"),
        ("b2", "09211234567"), ("b4", "09221234567"),
        ("shared", "09231234567"), ("unknown", "09241234567"),
    ]
    email_rows = [(str(i), f"{name}@example.com", "PL RECOVERY") for i, (name, _) in enumerate(contacts)]
    phone_rows = [(str(i), phone, "PL RECOVERY") for i, (_, phone) in enumerate(contacts)]
    phone_agents = csv_bytes(("Phone", "Bank Campaign"), [
        ("+63 917 123 4567", "SBC Level 1"), ("09211234567", "SBC_PL_RECOVERY"),
    ])
    phone_customers = csv_bytes(("ID", "Phone", "Bank Campaign"), [
        *phone_rows, ("duplicate", "+63 917 123 4567", "PL-Recovery"),
    ])
    return {
        "inbound_reply_email_lookup_uploads": [
            Upload(name, workbook_bytes([("OVERALL", rows)])) for name, rows in databases
        ],
        "inbound_reply_email_1_upload": Upload("history.csv", csv_bytes(("email", "sent"), [
            ("LEVEL1@example.com", "Agent response"), ("b2@example.com", "Agent response"),
        ])),
        "inbound_reply_email_2_upload": Upload("dc.csv", csv_bytes(("ID", "sender", "Campaign"), [
            *email_rows, ("duplicate", " LEVEL1@EXAMPLE.COM ", "PL RECOVERY"),
        ])),
        "inbound_reply_viber_1_upload": Upload("agents.csv", phone_agents),
        "inbound_reply_viber_2_upload": Upload("customers.csv", phone_customers),
        "inbound_reply_sms_1_upload": Upload("sms_agents.csv", phone_agents),
        "inbound_reply_sms_2_upload": Upload("sms_customers.csv", phone_customers),
    }


class RecoveryLookupPageTests(unittest.TestCase):
    def test_pending_recovery_uploads_show_one_campaign(self):
        from streamlit.testing.v1 import AppTest

        for label in ("PL RECOVERY", " pl recovery ", "SBC_PL_RECOVERY", "SBC - PL-Recovery"):
            with self.subTest(label=label):
                uploads = {
                    "inbound_reply_email_2_upload": Upload("dc.csv", csv_bytes(
                        ("sender", "Campaign"), [("level1@example.com", label)],
                    )),
                    "inbound_reply_viber_2_upload": Upload("customers.csv", csv_bytes(
                        ("Phone", "Bank Campaign"), [("09171234567", label)],
                    )),
                    "inbound_reply_sms_2_upload": Upload("sms_customers.csv", csv_bytes(
                        ("Phone", "Campaign"), [("09171234567", label)],
                    )),
                }
                with patch("branding.render_bank_branding"), \
                        patch("streamlit.file_uploader", side_effect=lambda title, **kw: uploads.get(kw["key"])), \
                        patch("streamlit.html") as html:
                    app = AppTest.from_file(str(ROOT / "pages" / "INBOUND_REPLY.py")).run(timeout=20)
                    self.assertFalse(app.exception)
                    self.assertEqual(app.session_state["inbound_reply_campaign_labels"]["labels"], ("PL RECOVERY",))
                    uploads["inbound_reply_email_lookup_uploads"] = recovery_uploads()["inbound_reply_email_lookup_uploads"]
                    app.run(timeout=20)
                    self.assertFalse(app.exception)
                    self.assertFalse(app.error)
                    call = next(call for call in reversed(html.call_args_list)
                                if 'aria-label="Inbound reply summary"' in call.args[0])
                    self.assertEqual([row[0] for row in table_cells(call.args[0])], CAMPAIGNS)

    def test_recovery_stays_in_its_campaign_in_summary_and_downloads(self):
        from streamlit.testing.v1 import AppTest

        uploads = recovery_uploads()
        with patch("branding.render_bank_branding"), \
                patch("streamlit.file_uploader", side_effect=lambda label, **kw: uploads.get(kw["key"])), \
                patch("streamlit.html") as html, patch("streamlit.download_button") as downloads:
            app = AppTest.from_file(str(ROOT / "pages" / "INBOUND_REPLY.py")).run(timeout=20)
            for reverse in (False, True):
                if reverse:
                    uploads["inbound_reply_email_lookup_uploads"].reverse()
                    app.run(timeout=20)
                self.assertFalse(app.exception)
                self.assertFalse(app.error)
                for channel, answered in (("email", 2), ("sms", 1), ("viber", 1)):
                    report = app.session_state[f"inbound_reply_{channel}_result"]["report"]
                    self.assertEqual([(c.campaign, c.unique_people, c.agent_replied) for c in report.campaigns],
                                     [("PL RECOVERY", 8, answered)])
                    source = report.dc if channel == "email" else report.message_logs
                    self.assertEqual(source.unmatched_addresses, 0)
                call = next(call for call in reversed(html.call_args_list)
                            if 'aria-label="Inbound reply summary"' in call.args[0])
                self.assertEqual([row[0] for row in table_cells(call.args[0])], CAMPAIGNS)
                for key, sheets in (
                    ("inbound_reply_summary_download", {"Email Unreplied": 6, "SMS Unreplied": 7, "Viber Unreplied": 7}),
                    ("inbound_reply_email_unreplied_download", {"Unreplied Details": 6}),
                    ("inbound_reply_viber_unreplied_download", {"Unreplied Details": 7}),
                ):
                    call = next(call for call in reversed(downloads.call_args_list) if call.kwargs.get("key") == key)
                    data = call.kwargs["data"]
                    book = load_workbook(BytesIO(data() if callable(data) else data))
                    try:
                        first_row = 4 if key == "inbound_reply_summary_download" else 2
                        for title, count in sheets.items():
                            rows = list(book[title].iter_rows(min_row=first_row, values_only=True))
                            self.assertEqual(len(rows), count)
                            headers = [cell.value for cell in book[title][first_row - 1]]
                            lookup_column = next(i for i, header in enumerate(headers) if header.endswith("Lookup Campaign"))
                            original_headers = (("Source: Campaign", "Source: Bank Campaign") if first_row == 4
                                                else ("Campaign", "Bank Campaign"))
                            original_column = next(i for i, header in enumerate(headers) if header in original_headers)
                            self.assertTrue(all(row[lookup_column] == "PL RECOVERY" for row in rows))
                            self.assertTrue(all(row[original_column] == "PL RECOVERY" for row in rows))
                    finally:
                        book.close()

    def test_recovery_survives_saved_lookup_changes_and_other_accounts_use_lookup(self):
        from streamlit.testing.v1 import AppTest

        uploads = recovery_uploads()
        files = [(upload.name, upload.getvalue()) for upload in uploads.pop("inbound_reply_email_lookup_uploads")]
        uploads["inbound_reply_email_2_upload"] = Upload("dc.csv", csv_bytes(("sender", "Campaign"), [
            ("level1@example.com", "PL RECOVERY"),
            ("level6@example.com", "Original campaign"),
            ("unknown@example.com", "Original campaign"),
        ]))
        with NamedTemporaryFile(suffix=".zip") as temporary:
            saved_path = Path(temporary.name)
        self.addCleanup(saved_path.unlink, missing_ok=True)
        with patch("branding.render_bank_branding"), \
                patch("inbound_reply_page.SAVED_CAMPAIGN_LOOKUP_PATH", saved_path), \
                patch("streamlit.file_uploader", side_effect=lambda label, **kw: uploads.get(kw["key"])):
            app = AppTest.from_file(str(ROOT / "pages" / "INBOUND_REPLY.py")).run(timeout=20)
            for with_lookup in (False, True, False):
                if with_lookup:
                    save_campaign_lookup_files(saved_path, files)
                else:
                    saved_path.unlink(missing_ok=True)
                app.run(timeout=20)
                self.assertFalse(app.exception)
                self.assertFalse(app.error)
                email = app.session_state["inbound_reply_email_result"]["report"]
                expected = [("PL RECOVERY", 1), ("SBC Level 6", 1)] if with_lookup else [
                    ("Original campaign", 2), ("PL RECOVERY", 1),
                ]
                self.assertEqual([(c.campaign, c.unique_people) for c in email.campaigns], expected)
                self.assertEqual(email.dc.unmatched_addresses, int(with_lookup))
                for channel in ("sms", "viber"):
                    report = app.session_state[f"inbound_reply_{channel}_result"]["report"]
                    self.assertEqual([(c.campaign, c.unique_people, c.agent_replied) for c in report.campaigns],
                                     [("PL RECOVERY", 8, 1)])


if __name__ == "__main__":
    unittest.main()
