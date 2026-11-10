from __future__ import annotations

import csv
import unittest
from datetime import datetime
from io import BytesIO, StringIO
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

from openpyxl import Workbook
from streamlit.testing.v1 import AppTest

from automation import AutomationError
from autostat import SourceSelection, inspect_contact_source
from b2_matching import build_b2_exports, build_b4_exports
from b2_page import no_matching_rows_message
from export_config import COLLECTION_AGENTS
from exports import build_csv_zip


ROOT = Path(__file__).resolve().parents[1]
HEADERS = ["Remarks", "chcode", "Account Key", "Contract Number", "Product Category Code", "Bucket Code",
           "Email (01)", "Mobile 1", "Name", "Total Outstanding", "Assignment Date"]
RECORDS = [
    ["Active", "SPM-1", "0001", "000000000000000001", "MC", 4,
     "one@example.com", "09171234567", "First", 100, datetime(2026, 11, 9)],
    ["Active", "SPM-2", "BEL-2", "BEL-002", "BEL", "B4",
     "invalid", "09181234567", "Second", 500, datetime(2026, 11, 9)],
    ["Active", None, "0003", "000003", "MC", 4,
     "three@example.com", "123", "Third", 300, datetime(2026, 11, 9)],
    ["Active", "SPM-4", "0004", "000004", "MC", 6,
     "four@example.com", "09171234567", "Other bucket", 100, None],
    ["Active", "SPM-5", "0005", "000005", "BEL", 70,
     "five@example.com", "09171234567", "Other bucket", 100, None],
    ["Active", "SPM-6", "0006", "000006", "MC", 4, "invalid", "123", "No contacts", 100, None],
    ["Active", "SPM-7", None, "000007", "MC", 4,
     "seven@example.com", "09171234567", "Missing key", 100, None],
    ["Active", "SPM-8", "0008", "000008", "PL", 4,
     "eight@example.com", "09171234567", "Other product", 100, None],
]


def workbook_data(headers=HEADERS, records=RECORDS, *, masks=None):
    workbook = Workbook()
    workbook.active.title = "Sheet1"
    workbook.active.append(headers)
    for record in records:
        workbook.active.append(record)
    for coordinate, number_format in (masks or {}).items():
        workbook.active[coordinate].number_format = number_format
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def selection(data, name="SPM ENDORSEMENT B4.xlsx"):
    return SourceSelection(name, data, inspect_contact_source(data, ".xlsx")[0])


def rows(file):
    return list(csv.DictReader(StringIO(file.data.decode("cp1252"))))


class SpmEndorsementTests(unittest.TestCase):
    def test_database_alone_filters_bucket_and_contacts_and_keeps_reference_separate_from_agent(self):
        result = build_b4_exports(selection(workbook_data()))
        cards, bel = result.bundle.files
        self.assertEqual([file.count for file in result.bundle.files], [2, 1])
        self.assertEqual([row["primary_key"] for row in rows(cards)], ["0003", "0001"])
        self.assertEqual([row["indexoneuid"] for row in rows(cards)], ["", "SPM-1"])
        self.assertEqual([row["CHCode"] for row in rows(cards)], ["AMRANCELI", "KPALIMA"])
        self.assertEqual(rows(cards)[1]["Terms"], "000000000000000001")
        self.assertEqual(rows(bel)[0]["indexoneuid"], "SPM-2")
        self.assertEqual(rows(bel)[0]["Promo"], "BEL")
        self.assertEqual(rows(bel)[0]["email"], "")
        self.assertEqual(rows(bel)[0]["phone"], "9181234567")
        self.assertEqual(rows(cards)[1]["EndoDate"], "2026-11-09 00:00:00")
        self.assertEqual(result.stats["Database accounts checked"], 8)
        self.assertEqual(result.stats["Other or missing bucket"], 2)
        self.assertEqual(result.stats["Missing account numbers"], 1)
        self.assertEqual(result.stats["Unsupported product"], 1)
        self.assertEqual(sum(summary["No valid contact"] for summary in result.bundle.summary), 1)
        self.assertTrue(all("Database row" in entry for entry in result.audit))
        with ZipFile(BytesIO(build_csv_zip(result.bundle.files))) as archive:
            exported = list(csv.DictReader(StringIO(archive.read(cards.filename).decode("cp1252"))))
            self.assertEqual(exported[1]["Terms"], "\t000000000000000001")
            self.assertEqual(exported[1]["indexoneuid"], "SPM-1")

    def test_database_reference_precedes_account_list_reference_and_blank_reference_can_fall_back(self):
        database = selection(workbook_data())
        accounts = selection(workbook_data(["Account No.", "Financing/Card No.", "Old I.C."], [
            ["0001", "000000000000000001", "OLDER-REFERENCE"], ["0003", "000003", "LIST-REFERENCE"],
        ]), "list.xlsx")
        result = build_b4_exports(database, accounts)
        output = {row["primary_key"]: row for file in result.bundle.files for row in rows(file)}
        self.assertEqual(set(output), {"0001", "0003"})
        self.assertEqual(output["0001"]["indexoneuid"], "SPM-1")
        self.assertEqual(output["0003"]["indexoneuid"], "LIST-REFERENCE")
        self.assertEqual(result.stats["DataGrid rows"], 2)

    def test_zero_masks_and_long_text_identifiers_survive_database_only_generation(self):
        data = workbook_data(records=[[
            "Active", 37, 1, "000012345678901234567890", "MC", 4,
            "one@example.com", "09171234567", "First", 100, None,
        ]], masks={"B2": "00000", "C2": "0000"})
        output = rows(build_b4_exports(selection(data)).bundle.files[0])[0]
        self.assertEqual((output["primary_key"], output["indexoneuid"], output["Terms"]),
                         ("0001", "00037", "000012345678901234567890"))

    def test_ambiguous_database_identities_do_not_export_and_b2_still_requires_account_list(self):
        source = selection(workbook_data(records=[RECORDS[0], RECORDS[0]]))
        result = build_b4_exports(source)
        self.assertEqual([file.count for file in result.bundle.files], [0, 0])
        self.assertEqual(result.stats["Conflicting or ambiguous keys"], 2)
        self.assertIn("both Account Key and Contract Number", no_matching_rows_message(result, bucket=4, database_only=True))
        with self.assertRaisesRegex(AutomationError, "B2 needs both"):
            build_b2_exports(source, None)


class Upload(BytesIO):
    def __init__(self, data):
        super().__init__(data)
        self.name = "SPM ENDORSEMENT B4.xlsx"


class SpmEndorsementPageTests(unittest.TestCase):
    def test_single_database_prepares_downloads_and_switching_modes_replaces_results(self):
        uploads = {"b4_database_upload": Upload(workbook_data())}
        with patch("streamlit.file_uploader", side_effect=lambda label, **kwargs: uploads.get(kwargs["key"])), \
                patch("b2_page.render_agent_availability", return_value=COLLECTION_AGENTS):
            app = AppTest.from_file(str(ROOT / "pages" / "NEGATIVE_B2.py")).run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertNotIn("b4_result", app.session_state)
            app.checkbox(key="b4_use_database_accounts").check().run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(len(app.error), 0)
            original = app.session_state["b4_result"]
            self.assertEqual([file.count for file in original["matched"].bundle.files], [2, 1])
            self.assertEqual(len(app.get("download_button")), 3)
            app.run(timeout=20)
            self.assertEqual(app.session_state["b4_result"]["signature"], original["signature"])
            uploads["b4_database_upload"] = Upload(workbook_data(records=[RECORDS[1]]))
            app.run(timeout=20)
            self.assertNotEqual(app.session_state["b4_result"]["signature"], original["signature"])
            self.assertEqual([file.count for file in app.session_state["b4_result"]["matched"].bundle.files], [0, 1])
            app.checkbox(key="b4_use_database_accounts").uncheck().run(timeout=20)
            self.assertNotIn("b4_result", app.session_state)
            self.assertEqual(len(app.get("download_button")), 0)
            app.checkbox(key="b4_use_database_accounts").check().run(timeout=20)
            app.button(key="clear_all_files").click().run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertNotIn("b4_result", app.session_state)
            self.assertTrue(app.checkbox(key="b4_use_database_accounts").value)
            self.assertEqual(len(app.get("download_button")), 0)
