from __future__ import annotations

import csv
import unittest
from datetime import datetime
from io import BytesIO, StringIO
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from streamlit.testing.v1 import AppTest

from automation import DEFAULT_TEMPLATE_HEADERS, FIELD_MAPPING
from export_config import CHANNEL_SUFFIXES, GEO_OUTPUT_FILENAMES, OUTPUT_FILENAMES


ROOT = Path(__file__).resolve().parents[1]


class Upload(BytesIO):
    def __init__(self, data: bytes):
        super().__init__(data)
        self.name = "mixed-workflows.xlsx"


def mixed_workbook(*, geo_negative=False) -> bytes:
    workbook = Workbook()
    grid = workbook.active
    grid.title = "DataGrid"
    grid.append(["No", "Name", "New I.C.", "Account No.", "Email", "Financing/Card No.", "Status Code"])
    grid.append([1, "DataGrid account", None, "GRID-001", "grid@example.com", "GRID-CARD-001", "ACTIVE"])

    def add_accounts(name: str, cases: list[tuple], *, geo: bool = False) -> None:
        sheet = workbook.create_sheet(name)
        headers = [*FIELD_MAPPING, *(["Area"] if geo else [])]
        sheet.append(headers)
        for account, product, level, status, area in cases:
            record = {
                "debtor_email": f"{account.lower()}@example.com",
                "OLD I.C": f"IC-{account}",
                "Account no.": account,
                "financing/Card no.": f"CARD-{account}",
                "Name": f"Account {account}",
                "endo date": datetime(2026, 10, 28),
                "level": level,
                "product": product,
                "OB (Masterlist)": 1234.5,
                "phone_1": "09171234567",
                "Christmas Amnesty": status,
                "Area": area,
            }
            sheet.append([record.get(header) for header in headers])

    add_accounts("NEGATIVE", [
        ("RECO001", "MC", 1, None, None),
        ("RECO006", " mc ", 6, None, None),
        ("RECOBEL", "BEL", 4, None, None),
        ("RECOLOAN", "PL", 3, None, None),
    ])
    add_accounts("GEO", [
        ("GEOM001", "MC", 1, "QUALIFIED", "Mindanao"),
        ("GEOM006", "mc", 6, "QUALIFIED", "Mindanao"),
        ("GEORBEL", "BEL", 3, None, "Region 3"),
    ], geo=True)
    if geo_negative:
        workbook["NEGATIVE"].title = "OTHER RECO"
        workbook["GEO"].title = "NEGATIVE"
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def csv_rows(file) -> list[dict[str, str]]:
    return list(csv.DictReader(StringIO(file.data.decode("cp1252"))))


class WorkflowPageTests(unittest.TestCase):
    def assert_no_embedded_b2(self, app) -> None:
        self.assertFalse(any(button.key == "b2_prepare" for button in app.button))
        self.assertFalse(any(widget.label == "SBC database sheet" for widget in app.selectbox))
        self.assertFalse(any("match accounts and contacts" in heading.value for heading in app.subheader))
        self.assertFalse(any(heading.value in ("B2 files", "B2 activity logs") for heading in app.subheader))

    def test_negative_page_only_exposes_b2_and_b4_matching_uploads(self):
        with patch("streamlit.file_uploader", return_value=None) as uploads:
            app = AppTest.from_file(str(ROOT / "pages" / "NEGATIVE_B2.py")).run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual({call.kwargs["key"] for call in uploads.call_args_list},
                             {"b2_database_upload", "b2_grid_upload", "b4_database_upload", "b4_grid_upload"})
            self.assertFalse(any("AutoStat" in panel.label for panel in app.expander))
            self.assertFalse(any(heading.value == "B2 files" for heading in app.subheader))

    def test_autostat_bucket_modes_compile_multiple_files_and_deliver_all_channels(self):
        for bucket in (2, 4):
            with self.subTest(bucket=bucket):
                prefix = f"autostat_b{bucket}_"
                csv_upload = Upload(
                    b"cardnumber,primary_key,FullName,email,phone,Terms,OB\n"
                    b"CARD1,000000000000000001,CSV Person,csv@example.com,+639171234567,COL1,9\n"
                    b"CARD2,000000000000000002,Email only,only@example.com,123,COL2,1000\n"
                )
                csv_upload.name = f"NEGATIVE B{bucket} - CARDS.csv"
                workbook = Workbook()
                workbook.active.append(["primary_key", "cardnumber", "FullName", "email", "phone", "Terms", "OB"])
                workbook.active.append(["000000000000000003", "CARD3", "Excel Person", None, "09181234567", "COL3", 20000])
                workbook.active.append(["000000000000000001", "OTHER-CARD", "Duplicate", "duplicate@example.com", "09191234567", "OTHER", 99999])
                output = BytesIO()
                workbook.save(output)
                workbook.close()
                excel_upload = Upload(output.getvalue())
                excel_upload.name = f"NEGATIVE B{bucket} - BEL.xlsx"
                files = [csv_upload, excel_upload]

                def uploader(label, **kwargs):
                    return files if kwargs.get("key") == prefix + "uploads" else None

                with patch("streamlit.file_uploader", side_effect=uploader) as upload, patch("streamlit.download_button") as download:
                    app = AppTest.from_file(str(ROOT / "pages" / "AUTOSTAT.py")).run(timeout=20)
                    app.radio(key="autostat_upload_type").set_value(f"B{bucket}").run(timeout=20)
                    self.assertEqual(len(app.exception), 0)
                    self.assertEqual(len(app.error), 0)
                    call = next(call for call in upload.call_args_list if call.kwargs["key"] == prefix + "uploads")
                    self.assertTrue(call.kwargs["accept_multiple_files"])
                    self.assertEqual(set(call.kwargs["type"]), {"csv", "xlsx"})
                    result = app.session_state[prefix + "result"]
                    self.assertEqual([file.count for file in result["bundle"].files], [2, 2, 2])
                    email = csv_rows(result["bundle"].files[0])
                    self.assertEqual([row["Account No."] for row in email], ["000000000000000002", "000000000000000001"])
                    self.assertEqual([row["Collector"] for row in email], ["COL2", "COL1"])
                    self.assertEqual(email[1]["REMARKS"], "csv@example.com | EMAIL SENT | COL1")
                    for file in result["bundle"].files[1:]:
                        self.assertEqual([row["Account number"] for row in csv_rows(file)], ["000000000000000003", "000000000000000001"])
                        self.assertEqual([row["Mobile Number"] for row in csv_rows(file)], ["09181234567", "09171234567"])
                    app.run(timeout=20)
                    self.assertEqual(app.session_state[prefix + "result"]["timestamp"], result["timestamp"])
                    for excel in (True, False):
                        if not excel:
                            app.selectbox(key=prefix + "output_format").select("CSV (.csv)").run(timeout=20)
                        calls = {call.kwargs.get("key"): call for call in download.call_args_list}
                        zip_call = calls[prefix + "zip"]
                        self.assertTrue(zip_call.args[2].endswith("_Excel.zip" if excel else "_CSV.zip"))
                        data = zip_call.args[1]()
                        with ZipFile(BytesIO(data)) as archive:
                            expected_names = [f"AUTOSTAT_B{bucket}_{channel}.{'xlsx' if excel else 'csv'}" for channel in ("Email", "SMS", "Viber")]
                            self.assertEqual(archive.namelist(), expected_names)
                            for index, name in enumerate(expected_names):
                                single = calls[prefix + f"download_{index}"]
                                self.assertEqual(single.args[2], name)
                                if excel:
                                    book = load_workbook(BytesIO(archive.read(name)))
                                    account_column = "B" if index == 0 else "A"
                                    self.assertTrue(all(cell.value.startswith("00000000000000000") and cell.data_type == "s"
                                                        and cell.number_format == "@" for cell in book.active[account_column][1:]))
                                    book.close()
                                else:
                                    self.assertEqual(archive.read(name), single.args[1]())
                                    rows = list(csv.DictReader(StringIO(archive.read(name).decode('cp1252'))))
                                    original = csv_rows(result['bundle'].files[index])
                                    for row, source in zip(rows, original, strict=True):
                                        for header, value in source.items():
                                            expected = '\t' + value if value and header in ('Account No.', 'Account number', 'Mobile Number') else value
                                            self.assertEqual(row[header], expected)
                    replacement = Upload(b"cardnumber,primary_key,email,phone\nCARD9,000009,updated@example.com,09191234567\n")
                    replacement.name = csv_upload.name
                    files[0] = replacement
                    app.run(timeout=20)
                    self.assertEqual(len(app.error), 0)
                    self.assertIn("000009", {row["Account No."] for row in csv_rows(app.session_state[prefix + "result"]["bundle"].files[0])})
                    files.clear()
                    app.run(timeout=20)
                    self.assertNotIn(prefix + "result", app.session_state)
                    self.assertEqual(len(app.get("download_button")), 0)

    def test_mixed_workbook_keeps_geo_and_reco_grouped_exports_on_main_page(self):
        upload = Upload(mixed_workbook())
        with patch("streamlit.file_uploader", side_effect=lambda label, **kw: upload if kw.get("key") == "source_upload" else None) as uploader:
            app = AppTest.from_file(str(ROOT / "app.py")).run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(app.title[0].value, "GEO & RECO")
            self.assertEqual({call.kwargs["key"] for call in uploader.call_args_list}, {"source_upload", "template_upload"})
            source = next(widget for widget in app.selectbox if widget.label == "Source sheet")
            self.assertFalse(source.disabled)
            self.assertEqual(source.value, "NEGATIVE")
            self.assert_no_embedded_b2(app)

            upload = Upload(mixed_workbook(geo_negative=True))
            app.run(timeout=20)
            self.assertEqual(app.radio[0].value, "Separate Area and Christmas Amnesty groups")
            app.run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(len(app.error), 0)
            result = app.session_state["result"]
            self.assertTrue(result["geo"])
            self.assertTrue(result["grouped"])
            for file in app.session_state["geo_autostat_result"]["bundle"].files:
                self.assertEqual({row["Account No."] for row in csv_rows(file)},
                                 {"CARD-GEOM001", "CARD-GEOM006", "CARD-GEORBEL"})
            self.assertEqual(result["group_counts"]["mindanao_qualified_cards"], 2)
            self.assertEqual(result["group_counts"]["r3_blank_bel"], 1)
            self.assertEqual({file.filename for file in result["files"]}, {
                f"{GEO_OUTPUT_FILENAMES[group]}{CHANNEL_SUFFIXES[channel]}.csv"
                for group in ("mindanao_qualified_cards", "r3_blank_bel")
                for channel in ("email", "viber")
            })
            for file in result["files"]:
                rows = csv_rows(file)
                self.assertEqual(tuple(rows[0]), DEFAULT_TEMPLATE_HEADERS)
                self.assertTrue(all(row["primary_key"] in {"GEOM001", "GEOM006", "GEORBEL"} for row in rows))
                self.assertTrue(all(row["cardnumber"] == "CARD-" + row["primary_key"] for row in rows))
                self.assertTrue(all(row["Product"] == ("BEL" if row["primary_key"] == "GEORBEL" else "CARDS") for row in rows))
                expected_class = "SBC GEO Mindanao" if file.group_key.startswith("mindanao_") else "SBC GEO R3"
                self.assertTrue(all(row["doclass"] == expected_class for row in rows))
            self.assert_no_embedded_b2(app)

            upload = Upload(mixed_workbook())
            app.run(timeout=20)
            self.assertFalse(app.session_state["result"]["geo"])
            self.assertNotEqual(app.session_state["result"]["signature"], result["signature"])
            self.assertEqual(app.radio[0].value, "Separate Christmas Amnesty groups")
            app.run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(len(app.error), 0)
            result = app.session_state["result"]
            self.assertFalse(result["geo"])
            self.assertTrue(result["grouped"])
            for file in app.session_state["geo_autostat_result"]["bundle"].files:
                self.assertEqual({row["Account No."] for row in csv_rows(file)},
                                 {"CARD-RECO001", "CARD-RECO006", "CARD-RECOBEL", "CARD-RECOLOAN"})
            groups = ("blank_cards_level_1", "blank_cards_level_6", "blank_bel", "blank_other_products")
            self.assertTrue(all(result["group_counts"][group] == 1 for group in groups))
            self.assertEqual({file.filename for file in result["files"]}, {
                f"{OUTPUT_FILENAMES[group]}{CHANNEL_SUFFIXES[channel]}.csv"
                for group in groups for channel in ("email", "viber")
            })
            self.assertEqual({row["primary_key"] for file in result["files"] for row in csv_rows(file)},
                             {"RECO001", "RECO006", "RECOBEL", "RECOLOAN"})
            self.assertTrue(all(row["cardnumber"] == "CARD-" + row["primary_key"]
                                for file in result["files"] for row in csv_rows(file)))
            self.assertTrue(all(row["Product"] == "CARDS" for file in result["files"] for row in csv_rows(file)
                                if row["primary_key"] in {"RECO001", "RECO006"}))
            with ZipFile(BytesIO(result["zip_data"])) as archive:
                self.assertEqual(set(archive.namelist()), {file.filename for file in result["files"]})
                for file in result["files"]:
                    self.assertEqual(archive.read(file.filename), file.data)
            self.assert_no_embedded_b2(app)

    def test_remembered_b2_workflow_does_not_replace_autostat(self):
        with patch("streamlit.file_uploader", return_value=None):
            app = AppTest.from_file(str(ROOT / "pages" / "AUTOSTAT.py"))
            app.session_state["autostat_workflow"] = "B2: Match DataGrid with SBC database"
            app.run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(app.title[0].value, "AUTOSTAT")
            self.assertTrue(any(widget.key == "autostat_remark_by" for widget in app.text_input))
            self.assertNotIn("autostat_result", app.session_state)
            self.assertFalse(any(widget.key == "autostat_workflow" for widget in app.radio))
            self.assert_no_embedded_b2(app)


if __name__ == "__main__":
    unittest.main()
