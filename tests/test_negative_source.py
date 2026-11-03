from __future__ import annotations

import csv
import hashlib
import unittest
from dataclasses import replace
from io import BytesIO, StringIO
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from streamlit.testing.v1 import AppTest

from automation import AutomationError, SheetInfo, create_default_csv_template, negative_source_sheet
from exports import (
    build_csv_zip, build_excel_zip, build_geo_exports, build_negative_exports,
    source_account_counts, validate_negative_exports,
)

ROOT = Path(__file__).resolve().parents[1]


class Upload(BytesIO):
    name = "same-workbook.xlsx"


def workbook(*, geo=False, account="00001", negative=True, empty=False):
    book = Workbook()
    book.remove(book.active)
    headers = ["primary_key", "cardnumber", "indexoneuid", "email", "phone", "OB", "Product", "Level", "ChristmasAmnesty"]
    if geo:
        headers.append("Area")
    for name, key in [("OVERALL", "OTHER01"), ("NEGATIVE", account), ("RPC NO PTP", "OTHER02")]:
        if name == "NEGATIVE" and not negative:
            continue
        sheet = book.create_sheet(name)
        sheet.append(headers)
        if name != "NEGATIVE" or not empty:
            values = [key, "CARD-" + key, "REF-" + key, "one@example.com", "09171234567", 100, "MC", 6, "QUALIFIED"]
            sheet.append([*values, "Mindanao"] if geo else values)
    data = BytesIO()
    book.save(data)
    book.close()
    return Upload(data.getvalue())


class NegativeSourcePolicyTests(unittest.TestCase):
    def test_requires_one_negative_sheet_without_fallback(self):
        negative = SheetInfo(" Negative ", 1, 0)
        self.assertIs(negative_source_sheet([SheetInfo("OVERALL", 1, 20), negative]), negative)
        with self.assertRaisesRegex(AutomationError, "no NEGATIVE"):
            negative_source_sheet([SheetInfo("OVERALL", 1, 20)])
        with self.assertRaisesRegex(AutomationError, "More than one"):
            negative_source_sheet([negative, SheetInfo("NEGATIVE", 1, 20)])

    def test_geo_and_reco_ignore_old_sheet_choices_and_reject_other_sheet_accounts(self):
        for geo in (False, True):
            with self.subTest(geo=geo):
                uploads = [workbook(geo=geo)]
                digest = hashlib.sha256(uploads[0].getvalue()).hexdigest()
                with patch("streamlit.file_uploader", side_effect=lambda label, **kw: uploads[0] if kw["key"] == "source_upload" else None):
                    app = AppTest.from_file(str(ROOT / "app.py"))
                    app.session_state[f"source_sheet_{digest}"] = "OVERALL"
                    app.run(timeout=20)
                    self.assertEqual(len(app.exception), 0)
                    self.assertEqual(len(app.error), 0)
                    result = app.session_state["result"]
                    self.assertEqual(result["source_sheet"], "NEGATIVE")
                    self.assertEqual(result["source_sha256"], digest)
                    self.assertEqual(result["geo"], geo)
                    self.assertFalse(any(widget.label == "Source sheet" for widget in app.selectbox))
                    self.assertEqual({record["primary_key"] for file in result["files"] for record in file.records}, {"00001"})
                    for file in app.session_state["geo_autostat_result"]["bundle"].files:
                        rows = csv.DictReader(StringIO(file.data.decode("cp1252")))
                        self.assertEqual({row["Account No."] for row in rows}, {"CARD-00001"})
                    with ZipFile(BytesIO(build_csv_zip(result["files"]))) as archive:
                        for name in archive.namelist():
                            rows = csv.DictReader(StringIO(archive.read(name).decode("cp1252")))
                            self.assertEqual({row["primary_key"].strip() for row in rows}, {"00001"})
                    with ZipFile(BytesIO(build_excel_zip(result["files"]))) as archive:
                        for name in archive.namelist():
                            book = load_workbook(BytesIO(archive.read(name)))
                            self.assertEqual({row[1].value for row in book.active.iter_rows(min_row=2)}, {"00001"})
                            book.close()
                    uploads[0] = workbook(geo=geo, account="00002")
                    app.run(timeout=20)
                    self.assertNotEqual(app.session_state["result"]["source_sha256"], digest)
                    self.assertEqual({record["primary_key"] for file in app.session_state["result"]["files"] for record in file.records}, {"00002"})

    def test_missing_or_empty_negative_never_falls_back_or_keeps_previous_downloads(self):
        for geo in (False, True):
            for missing in (False, True):
                with self.subTest(geo=geo, missing=missing):
                    uploads = [workbook(geo=geo)]
                    with patch("streamlit.file_uploader", side_effect=lambda label, **kw: uploads[0] if kw["key"] == "source_upload" else None):
                        app = AppTest.from_file(str(ROOT / "app.py")).run(timeout=20)
                        self.assertIn("result", app.session_state)
                        uploads[0] = workbook(geo=geo, negative=not missing, empty=not missing)
                        app.run(timeout=20)
                        self.assertEqual(len(app.exception), 0)
                        if missing:
                            self.assertNotIn("result", app.session_state)
                            self.assertNotIn("geo_autostat_result", app.session_state)
                        else:
                            self.assertEqual(app.session_state["result"]["count"], 0)
                            self.assertFalse(app.session_state["result"].get("files"))
                            if "geo_autostat_result" in app.session_state:
                                activity = app.session_state["geo_autostat_result"]
                                self.assertTrue("error" in activity or all(file.count == 0 for file in activity["bundle"].files))
                        self.assertFalse(any(item.key in {"csv_zip", "excel_zip", "geo_autostat_excel"}
                                             for item in app.get("download_button")))
                        if missing:
                            self.assertTrue(any("no NEGATIVE" in item.value for item in app.error))

    def test_alien_cached_output_blocks_all_campaign_and_activity_downloads(self):
        for geo in (False, True):
            upload = workbook(geo=geo)
            with self.subTest(geo=geo), patch("streamlit.file_uploader", side_effect=lambda label, **kw: upload if kw["key"] == "source_upload" else None):
                app = AppTest.from_file(str(ROOT / "app.py")).run(timeout=20)
                result = app.session_state["result"]
                file = result["files"][0]
                text = file.data.decode("cp1252").replace("00001", "ALIEN")
                result["files"][0] = replace(file, data=text.encode("cp1252"),
                                               records=[{**record, "primary_key": "ALIEN", "cardnumber": "CARD-ALIEN", "indexoneuid": "REF-ALIEN"}
                                                        for record in file.records])
                app.run(timeout=20)
                self.assertEqual(len(app.exception), 0)
                self.assertTrue(any("Downloads blocked" in item.value for item in app.error))
                self.assertNotIn("files", app.session_state["result"])
                self.assertNotIn("geo_autostat_result", app.session_state)
                self.assertFalse(any(item.key in {"csv_zip", "excel_zip", "geo_autostat_excel"}
                                     for item in app.get("download_button")))


class NegativeMembershipTests(unittest.TestCase):
    def test_checks_source_identities_counts_csv_and_excel_records_for_both_workflows(self):
        records = [{"primary_key": "00001", "cardnumber": "CARD-1", "indexoneuid": "REF-1", "email": "one@example.com",
                    "phone": "09171234567", "product": "MC", "level": 6, "ChristmasAmnesty": "QUALIFIED", "AREA": "Mindanao"}]
        source = source_account_counts(records)
        for builder in (build_geo_exports, build_negative_exports):
            for encoding in ("cp1252", "utf-8-sig"):
                with self.subTest(builder=builder.__name__, encoding=encoding):
                    files = builder(create_default_csv_template(), records, "", 1, ".csv", encoding=encoding).files
                    validate_negative_exports(files, source, source_sheet="NEGATIVE")
                    with self.assertRaisesRegex(AutomationError, "NEGATIVE worksheet only"):
                        validate_negative_exports(files, source, source_sheet="OVERALL")
                    alien = [{**records[0], "cardnumber": "OTHER-CARD"}]
                    alien_files = builder(create_default_csv_template(), alien, "", 1, ".csv", encoding=encoding).files
                    with self.assertRaisesRegex(AutomationError, "not verified"):
                        validate_negative_exports(alien_files, source, source_sheet="NEGATIVE")
                    with self.assertRaisesRegex(AutomationError, "not verified"):
                        validate_negative_exports(files + files, source, source_sheet="NEGATIVE")
                    corrupt = replace(files[0], data=files[0].data.replace(b"CARD-1", b"OTHER-CARD"))
                    with self.assertRaisesRegex(AutomationError, "inconsistent account data"):
                        validate_negative_exports([corrupt], source, source_sheet="NEGATIVE")


if __name__ == "__main__":
    unittest.main()
