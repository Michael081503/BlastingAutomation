from __future__ import annotations

import csv
import unittest
from datetime import datetime
from hashlib import sha256
from io import BytesIO, StringIO
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook
from streamlit.testing.v1 import AppTest

from automation import (
    DEFAULT_TEMPLATE_HEADERS, AutomationError, assignment_scheme_level,
    create_default_csv_template, inspect_workbook, read_source,
)
from export_config import GEO_RECO_AGENTS
from exports import (
    build_export_excel, build_negative_exports, separate_negative,
    source_account_counts, validate_source_exports,
)


ROOT = Path(__file__).resolve().parents[1]
HEADERS = (
    "ACA", "ACCOUNT STATUS", "EFS STATUS", "DESIGNATION", "Product",
    "Card Number/ Short PN", "Account Code", "Assignment Scheme", "Customer Name",
    "Currency", "Assignment Date", "Pullout date", "PRIN", "AIR", "OB",
    "Product Description", "DOB", "PHONE 1", "PHONE 2", "EMAIL ADD",
    "ADDRESS 1", "ADDRESS 2", "Account Type", "WO date", "PORTFOLIO GROUP",
)


def source_data(*, scheme="Recovery (L1 Unsecured)", extra_headers=(), rows=None):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "SPM List"
    headers = (*HEADERS, *extra_headers)
    sheet.append(headers)
    rows = rows if rows is not None else [{"Assignment Scheme": scheme}]
    for index, updates in enumerate(rows, 1):
        record = {
            "Product": "MC", "Card Number/ Short PN": f"0000123456789012345{index}",
            "Account Code": f"000{index}", "Assignment Scheme": scheme,
            "Customer Name": f"Customer {index}", "Assignment Date": datetime(2026, 11, 18),
            "OB": index * 100, "PHONE 1": "09171234567", "EMAIL ADD": f"person{index}@example.com",
        }
        record.update(updates)
        sheet.append([record.get(header) for header in headers])
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


class Upload(BytesIO):
    def __init__(self, data):
        super().__init__(data)
        self.name = "SPM_FRESH ENDO.xlsx"


class AssignmentSchemeTests(unittest.TestCase):
    def test_fresh_layout_maps_identifiers_contacts_dates_and_placement(self):
        data = source_data()
        info = inspect_workbook(data)[0]
        self.assertEqual(info.header_row, 1)
        self.assertFalse(info.has_area)
        self.assertTrue({"primary_key", "cardnumber", "level", "email", "phone"} <= set(info.available_fields))
        record = read_source(data, "SPM List", 1, allow_missing=True).records[0]
        self.assertEqual(record["primary_key"], "0001")
        self.assertEqual(record["cardnumber"], "00001234567890123451")
        self.assertEqual(record["FullName"], "Customer 1")
        self.assertEqual(record["endodate"], datetime(2026, 11, 18))
        self.assertEqual(record["level"], "Level 1")
        self.assertEqual(record["product"], "CARDS")
        self.assertEqual(record["email"], "person1@example.com")
        self.assertEqual(record["phone"], "09171234567")
        self.assertIsNone(record["ChristmasAmnesty"])
        self.assertIsNone(record["indexoneuid"])

    def test_scheme_levels_are_unambiguous_and_unknowns_go_to_review(self):
        cases = (
            ("Recovery (L1 Unsecured)", "Level 1", "blank_cards_level_1"),
            (" recovery (l6 unsecured) ", "Level 6", "blank_cards_level_6"),
            ("Recovery (Level 1)", "Level 1", "blank_cards_level_1"),
            ("Recovery (L10 Unsecured)", "Level 10", "blank_review"),
            ("Recovery (L1 / L6)", "Recovery (L1 / L6)", "blank_review"),
            ("Recovery (PL1)", "Recovery (PL1)", "blank_review"),
            ("Recovery (L1.5)", "Recovery (L1.5)", "blank_review"),
            ("Unspecified", "Unspecified", "blank_review"),
            (None, None, "blank_review"),
        )
        for value, expected, group in cases:
            with self.subTest(value=value):
                self.assertEqual(assignment_scheme_level(value), expected)
                groups = separate_negative([{"product": "MC", "level": value, "email": "one@example.com"}])
                self.assertEqual([key for key, records in groups.items() if records], [group])

    def test_existing_reco_template_classes_assignment_and_source_membership_are_retained(self):
        data = source_data(rows=[
            {"Assignment Scheme": "Recovery (L1 Unsecured)"},
            {"Assignment Scheme": "Recovery (L6 Unsecured)"},
            {"Product": "BEL"}, {"Product": "PL"},
        ])
        mapped = read_source(data, "SPM List", 1, allow_missing=True)
        bundle = build_negative_exports(
            create_default_csv_template(), mapped.records, "", 1, ".csv",
            source_sheet="SPM List", agents=GEO_RECO_AGENTS,
        )
        self.assertEqual({file.group_key for file in bundle.files},
                         {"blank_cards_level_1", "blank_cards_level_6", "blank_bel", "blank_other_products"})
        validate_source_exports(bundle.files, source_account_counts(mapped.records), source_sheet="SPM List")
        for file in bundle.files:
            rows = list(csv.DictReader(StringIO(file.data.decode("cp1252"))))
            self.assertEqual(tuple(rows[0]), DEFAULT_TEMPLATE_HEADERS)
            expected_class = "SBC Cards L1" if "L1" in file.filename else "SBC Cards L6" if "L6" in file.filename else "PL Recovery"
            self.assertEqual(rows[0]["doclass"], expected_class)
            self.assertIn(rows[0]["CHCode"], {code for _, code in GEO_RECO_AGENTS})
            self.assertEqual(rows[0]["ChristmasAmnesty"], "")
            workbook = load_workbook(BytesIO(build_export_excel(file)))
            try:
                headers = [cell.value for cell in workbook.active[1]]
                card = workbook.active.cell(2, headers.index("Terms") + 1)
                self.assertEqual(card.data_type, "s")
                self.assertTrue(card.value.startswith("0000123456789"))
            finally:
                workbook.close()

    def test_manual_level_selection_is_respected_and_duplicates_stay_ambiguous(self):
        data = source_data(extra_headers=("Level",), rows=[{"Level": 6}])
        with self.assertRaisesRegex(AutomationError, "ambiguous"):
            read_source(data, "SPM List", 1, allow_missing=True)
        record = read_source(data, "SPM List", 1, allow_missing=True, source_columns={"level": len(HEADERS)}).records[0]
        self.assertEqual(record["level"], 6)
        record = read_source(data, "SPM List", 1, allow_missing=True, source_columns={"level": 7}).records[0]
        self.assertEqual(record["level"], "Level 1")


class AssignmentSchemePageTests(unittest.TestCase):
    def test_upload_groups_without_amnesty_and_replacing_upload_rebuilds_campaigns_and_logs(self):
        uploads = {"source_upload": Upload(source_data())}
        with patch("streamlit.file_uploader", side_effect=lambda label, **kwargs: uploads.get(kwargs["key"])), \
                patch("agent_cycle_page.render_agent_cycle", return_value=(datetime(2026, 11, 19).date(), GEO_RECO_AGENTS)):
            app = AppTest.from_file(str(ROOT / "app.py"))
            # A session from before this mapping only offered all-accounts mode.
            app.session_state["sheet_output_modes"] = {
                (sha256(uploads["source_upload"].getvalue()).hexdigest(), "SPM List", 1): "All accounts in one group",
            }
            app.run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(len(app.error), 0)
            result = app.session_state["result"]
            self.assertTrue(result["grouped"])
            self.assertFalse(result["geo"])
            self.assertEqual(app.radio[0].value, "Separate RECO placement groups")
            self.assertTrue(all("L1" in file.filename for file in result["files"]))
            self.assertEqual([file.count for file in result["files"]], [1, 1])
            logs = app.session_state["geo_autostat_result"]["bundle"]
            self.assertEqual([file.count for file in logs.files], [1, 1])
            self.assertEqual(logs.files[0].preview[0]["Account No."], "00001234567890123451")
            uploads["source_upload"] = Upload(source_data(scheme="Recovery (L6 Unsecured)"))
            app.run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(len(app.error), 0)
            self.assertNotEqual(app.session_state["result"]["signature"], result["signature"])
            self.assertTrue(all("L6" in file.filename for file in app.session_state["result"]["files"]))
            app.radio[0].set_value("All accounts in one group").run(timeout=20)
            self.assertFalse(app.session_state["result"]["grouped"])
            app.run(timeout=20)
            self.assertFalse(app.session_state["result"]["grouped"])

    def test_assignment_scheme_with_area_uses_geo_groups(self):
        upload = Upload(source_data(extra_headers=("Area",), rows=[{"Area": "Region 3"}]))
        with patch("streamlit.file_uploader", side_effect=lambda label, **kwargs: upload if kwargs["key"] == "source_upload" else None):
            app = AppTest.from_file(str(ROOT / "app.py")).run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(len(app.error), 0)
            result = app.session_state["result"]
            self.assertTrue(result["grouped"])
            self.assertTrue(result["geo"])
            self.assertEqual({file.group_key for file in result["files"]}, {"r3_blank_cards"})
            self.assertTrue(all(record["doclass"] == "SBC GEO R3" for file in result["files"] for record in file.records))


if __name__ == "__main__":
    unittest.main()
