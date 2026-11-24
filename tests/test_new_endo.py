from __future__ import annotations

import csv
from datetime import datetime
import unittest
from io import BytesIO, StringIO
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

from msoffcrypto.format.ooxml import OOXMLFile
from openpyxl import Workbook, load_workbook
from streamlit.testing.v1 import AppTest
from streamlit.proto.TextInput_pb2 import TextInput

from automation import (
    AutomationError, DEFAULT_TEMPLATE_HEADERS, create_default_csv_template,
    inspect_workbook, read_source, source_column_positions,
)
from export_config import GEO_AGENTS, GEO_RECO_AGENTS
from exports import (
    build_csv_zip, build_export_csv, build_export_excel, build_geo_exports,
    build_negative_exports, build_new_endo_exports, build_sheet_exports,
    source_account_counts, validate_source_exports,
)
from workbook_password import needs_workbook_password, unlock_workbook


ROOT = Path(__file__).resolve().parents[1]
TEST_PASSWORD = "synthetic-workbook-password"


def source_data(*, geo=False, blank_tag=False, tag_header=None):
    workbook = Workbook()
    workbook.active.title = "Summary"
    workbook.active.append(["Product", "Total"])
    workbook.active.append(["MC", 2])
    sheet = workbook.create_sheet("Sheet1" if geo else "SPM List")
    if geo:
        sheet.append([
            "CURRENT ACA", "PRODUCT", "CARD NUMBER / SHORT PN #", "ACCOUNT NUMBER / LONG PN #",
            tag_header or "TAGGING", "CORRECT ASSIGNMENT SCHEME", "CUSTOMER NAME", "ASSIGNMENT DATE", "OB",
            "Region", "PHONE 1", "EMAIL ADDRESS", "Level",
        ])
        sheet.append(["SPM", "MC", "00001234567890123451", "0001", "SOURCE-MINDANAO", "Recovery (L6 Unsecured)",
                      "First customer", None, 100, "Mindanao", "09171234567", "first@example.com", "Level 6"])
        sheet.append(["SPM", "MC", "00001234567890123452", "0002", None if blank_tag else "SOURCE-R3", "Recovery (L6 Unsecured)",
                      "Second customer", None, 900, "Region 3", "09181234567", "second@example.com", "Level 6"])
    else:
        sheet.append(["Product", "Card Number/ Short PN", "Account Code", tag_header or "CMS TAGGING", "Assignment Scheme",
                      "Customer Name", "Assignment Date", "OB", "PHONE 1", "EMAIL ADD"])
        sheet.append(["MC", "00001234567890123451", "0001", "SOURCE-L1-A", "Recovery (L1 Unsecured)",
                      "First customer", None, 100, "09171234567", "first@example.com"])
        sheet.append(["MC", "00001234567890123452", "0002", None if blank_tag else "SOURCE-L1-B", "Recovery (L1 Unsecured)",
                      "Second customer", None, 900, "09181234567", "second@example.com"])
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def encrypted(data):
    output = BytesIO()
    OOXMLFile(BytesIO(data)).encrypt(TEST_PASSWORD, output)
    return output.getvalue()


def csv_rows(file):
    return list(csv.DictReader(StringIO(file.data.decode("cp1252"))))


def blank_region_endo():
    workbook = load_workbook(BytesIO(source_data(geo=True, tag_header="TAG TO")))
    sheet = workbook["Sheet1"]
    for row in (2, 3):
        sheet.cell(row, 10).value = None
        sheet.cell(row, 8).value = datetime(2026, 11, 13)
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


class Upload(BytesIO):
    def __init__(self, data, name="NEW ENDO.xlsx"):
        super().__init__(data)
        self.name = name


class NewEndoMappingTests(unittest.TestCase):
    def test_geo_region_long_pn_and_tagging_are_detected_without_using_summary_rows(self):
        data = source_data(geo=True)
        info = inspect_workbook(data)[1]
        self.assertTrue(info.has_area)
        self.assertTrue({"AREA", "cms_tagging", "primary_key", "cardnumber", "phone", "email"} <= set(info.available_fields))
        records = read_source(data, info.name, info.header_row, allow_missing=True).records
        self.assertEqual([(record["primary_key"], record["cardnumber"]) for record in records],
                         [("0001", "00001234567890123451"), ("0002", "00001234567890123452")])
        self.assertEqual([record["AREA"] for record in records], ["Mindanao", "Region 3"])
        self.assertEqual([record["cms_tagging"] for record in records], ["SOURCE-MINDANAO", "SOURCE-R3"])

    def test_source_codes_and_blank_monthly_survive_grouping_channels_and_excel(self):
        for geo, tag_header in ((False, None), (True, None), (False, "TAG TO"), (True, "TAG TO")):
            with self.subTest(geo=geo, tag_header=tag_header):
                data = source_data(geo=geo, blank_tag=True, tag_header=tag_header)
                sheet = "Sheet1" if geo else "SPM List"
                records = read_source(data, sheet, 1, allow_missing=True).records
                expected = {record["primary_key"]: record["cms_tagging"] or "" for record in records}
                builder = build_geo_exports if geo else build_negative_exports
                bundles = [
                    builder(create_default_csv_template(), records, "", 1, ".csv", source_sheet=sheet, agents=GEO_RECO_AGENTS),
                    build_sheet_exports(create_default_csv_template(), records, "", 1, template_extension=".csv",
                                        base_name="NEW ENDO - L1", label="NEW ENDO", agents=GEO_RECO_AGENTS, reco=not geo),
                ]
                for bundle in bundles:
                    for file in bundle.files:
                        rows = csv_rows(file)
                        self.assertEqual(tuple(file.headers), DEFAULT_TEMPLATE_HEADERS)
                        self.assertTrue(all(row["CHCode"] == expected[row["primary_key"]] for row in rows))
                        self.assertTrue(all(row["Monthly"] == "" and row["NewEndo"] == "" for row in rows))
                        workbook = load_workbook(BytesIO(build_export_excel(file)))
                        try:
                            values = list(workbook.active.values)
                            positions = {header: index for index, header in enumerate(values[0])}
                            for row in values[1:]:
                                self.assertEqual(row[positions["CHCode"]] or "", expected[row[positions["primary_key"]]])
                                self.assertIsNone(row[positions["Monthly"]])
                                self.assertTrue(row[positions["Terms"]].startswith("0000123456789"))
                        finally:
                            workbook.close()
                if geo:
                    self.assertEqual({file.group_key for file in bundles[0].files}, {"mindanao_blank_cards", "r3_blank_cards"})

    def test_custom_template_gets_chcode_and_tagged_rows_do_not_advance_automatic_assignment(self):
        records = [
            {"primary_key": "source", "cardnumber": "CARD-SOURCE", "product": "CARDS", "level": 1,
             "email": "source@example.com", "OB": 900, "cms_tagging": "FROM-SOURCE"},
            {"primary_key": "auto-1", "product": "CARDS", "level": 1, "email": "one@example.com", "OB": 500},
            {"primary_key": "auto-2", "product": "CARDS", "level": 1, "email": "two@example.com", "OB": 100},
        ]
        bundle = build_negative_exports(create_default_csv_template(), records, "", 1, ".csv", agents=GEO_RECO_AGENTS)
        rows = csv_rows(bundle.files[0])
        self.assertEqual([row["CHCode"] for row in rows], ["FROM-SOURCE", GEO_RECO_AGENTS[0][1], GEO_RECO_AGENTS[1][1]])
        headers = [header for header in DEFAULT_TEMPLATE_HEADERS if header not in {"Monthly", "CHCode"}]
        template = (",".join(headers) + "\n").encode()
        file = build_sheet_exports(template, records[:1], "", 1, template_extension=".csv", base_name="NEW ENDO", label="NEW ENDO").files[0]
        self.assertEqual(file.headers, [*headers, "CHCode"])
        self.assertEqual(csv_rows(file)[0]["CHCode"], "FROM-SOURCE")
        self.assertEqual(csv_rows(file)[0]["Terms"], "CARD-SOURCE")

    def test_tag_to_aliases_take_priority_over_other_tagging_columns_including_blank_cells(self):
        for header in ("TAG TO", "tag to", " TAG_TO ", "Tag-To"):
            with self.subTest(header=header):
                workbook = load_workbook(BytesIO(source_data(geo=True, blank_tag=True, tag_header=header)))
                sheet = workbook["Sheet1"]
                column = sheet.max_column + 1
                for offset, label in enumerate(("CMS TAGGING", "TAGGING")):
                    sheet.cell(1, column + offset, label)
                    sheet.cell(2, column + offset, "OTHER-AGENT")
                    sheet.cell(3, column + offset, "DO-NOT-FILL-BLANK")
                output = BytesIO()
                workbook.save(output)
                workbook.close()
                data = output.getvalue()
                info = inspect_workbook(data)[1]
                self.assertEqual(source_column_positions(info.headers)["cms_tagging"], 4)
                records = read_source(data, info.name, info.header_row, allow_missing=True).records
                self.assertEqual([row["cms_tagging"] for row in records], ["SOURCE-MINDANAO", None])

    def test_duplicate_tag_to_columns_require_an_explicit_mapping(self):
        workbook = load_workbook(BytesIO(source_data(tag_header="TAG TO")))
        sheet = workbook["SPM List"]
        sheet.cell(1, sheet.max_column + 1, "TAG_TO")
        output = BytesIO()
        workbook.save(output)
        workbook.close()
        data = output.getvalue()
        with self.assertRaisesRegex(AutomationError, "Duplicate or ambiguous header"):
            read_source(data, "SPM List", 1, allow_missing=True)
        records = read_source(data, "SPM List", 1, allow_missing=True, source_columns={"cms_tagging": 3}).records
        self.assertEqual([row["cms_tagging"] for row in records], ["SOURCE-L1-A", "SOURCE-L1-B"])


class CombinedNewEndoTests(unittest.TestCase):
    def test_saved_agent_names_supply_chcode_in_new_endo_and_geo_downloads(self):
        tags = [
            ("CALIMQUIM, JULIUS SANCHEZ", "JSCALIMQUIM"),
            ("  barua, jehan a  ", "JBARUA"),
            ("Erika Mae C. Dela Cruz", "EMCDELACRUZ"),
            ("BULATAO, JOSHUA BORJA - JBBULATAO", "JBBULATAO"),
            ("jb arua", "JBARUA"),
            ("SOURCE-AGENT", "SOURCE-AGENT"),
            (None, ""),
        ]
        expected = {f"{index:04}": code for index, (_, code) in enumerate(tags, 1)}

        def tagged_source(area=None):
            workbook = Workbook()
            workbook.active.title = "SPM List"
            workbook.active.append(["Product", "Card Number/ Short PN", "Account Code", "TAG TO", "Assignment Scheme",
                                    "Customer Name", "PHONE 1", "EMAIL ADD", "Region"])
            for index, (tag, _) in enumerate(tags, 1):
                workbook.active.append(["MC", f"{index:08}", f"{index:04}", tag, "Recovery (L6 Unsecured)",
                                        "Sample customer", "09171234567", f"sample{index}@example.com", area])
            output = BytesIO()
            workbook.save(output)
            workbook.close()
            return Upload(output.getvalue())

        uploads = {"source_upload": tagged_source()}
        with patch("streamlit.file_uploader", side_effect=lambda label, **kwargs: uploads.get(kwargs["key"])):
            app = AppTest.from_file(str(ROOT / "app.py")).run(timeout=20)
            for area in (None, "Mindanao"):
                if area is not None:
                    uploads["source_upload"] = tagged_source(area)
                    app.run(timeout=20)
                self.assertFalse(app.exception)
                self.assertFalse(app.error)
                result = app.session_state["result"]
                self.assertEqual(result["new_endo"], area is None)
                self.assertEqual(result["geo"], area is not None)
                for file in result["files"]:
                    self.assertEqual({row["primary_key"]: row["CHCode"] for row in csv_rows(file)}, expected)
                    workbook = load_workbook(BytesIO(build_export_excel(file)))
                    try:
                        headers, *rows = workbook.active.values
                        self.assertEqual({row[headers.index("primary_key")]: row[headers.index("CHCode")] or "" for row in rows},
                                         expected)
                    finally:
                        workbook.close()

    def test_reference_layout_keeps_both_contacts_codes_dates_and_account_protection(self):
        records = read_source(blank_region_endo(), "Sheet1", 1, allow_missing=True).records
        bundle = build_new_endo_exports(create_default_csv_template(), records, "", 1, ".csv")
        self.assertEqual([file.filename for file in bundle.files], ["NEW ENDO L6.csv"])
        file = bundle.files[0]
        self.assertEqual(file.count, 2)
        self.assertEqual(file.headers, list(DEFAULT_TEMPLATE_HEADERS))
        rows = list(csv.DictReader(StringIO(build_export_csv(file).decode("cp1252"))))
        self.assertEqual([row["primary_key"] for row in rows], ["\t0002", "\t0001"])
        self.assertEqual([row["CHCode"] for row in rows], ["SOURCE-R3", "SOURCE-MINDANAO"])
        self.assertEqual([row["phone"] for row in rows], ["9181234567", "9171234567"])
        self.assertTrue(all(row["email"] and row["phone"] and row["Terms"].startswith("\t0000") for row in rows))
        self.assertTrue(all(row["doclass"] == "SBC Cards L6" and row["EndoDate"] == "13/11/2026 0:00" for row in rows))
        self.assertTrue(all(row[key] == "" for row in rows for key in ("Monthly", "AREA", "ChristmasAmnesty", "NewEndo")))
        validate_source_exports(bundle.files, source_account_counts(records), source_sheet="Sheet1")
        with ZipFile(BytesIO(build_csv_zip(bundle.files))) as archive:
            self.assertEqual(archive.namelist(), ["NEW ENDO L6.csv"])
            self.assertEqual(archive.read(file.filename), build_export_csv(file))
        book = load_workbook(BytesIO(build_export_excel(file)))
        try:
            headers = {cell.value: cell.column for cell in book.active[1]}
            self.assertEqual(book.active.cell(2, headers["EndoDate"]).value, datetime(2026, 11, 13))
            self.assertEqual(book.active.cell(2, headers["CHCode"]).value, "SOURCE-R3")
            self.assertEqual(book.active.cell(2, headers["phone"]).value, "9181234567")
            self.assertEqual(book.active.cell(2, headers["Terms"]).value, "00001234567890123452")
        finally:
            book.close()

    def test_placements_partial_contacts_and_review_rows_remain_separate(self):
        records = [
            {"primary_key": "L1", "product": "MC", "level": 1, "email": "one@example.com", "cms_tagging": "AGENT-1"},
            {"primary_key": "L6", "product": "CARDS", "level": 6, "phone": "09171234567", "cms_tagging": None},
            {"primary_key": "PL", "product": "BEL", "level": 6, "email": "pl@example.com", "cms_tagging": "AGENT-PL"},
            {"primary_key": "REVIEW", "product": "CARDS", "level": "Unknown", "email": "review@example.com", "cms_tagging": "AGENT-R"},
            {"primary_key": "NO-CONTACT", "product": "CARDS", "level": 6, "cms_tagging": "NO-CONTACT"},
        ]
        original = [dict(record) for record in records]
        bundle = build_new_endo_exports(create_default_csv_template(), records, "", 1, ".csv")
        self.assertEqual({file.filename for file in bundle.files},
                         {"NEW ENDO L1.csv", "NEW ENDO L6.csv", "NEW ENDO PL.csv", "NEW ENDO CARDS REVIEW.csv"})
        self.assertEqual(bundle.no_contact_rows_skipped, 1)
        self.assertEqual(sum(file.count for file in bundle.files), 4)
        l6 = next(file for file in bundle.files if file.filename == "NEW ENDO L6.csv")
        self.assertEqual(csv_rows(l6)[0]["CHCode"], "")
        self.assertEqual(csv_rows(l6)[0]["email"], "")
        validate_source_exports(bundle.files, source_account_counts(records[:-1]), source_sheet="Accounts")
        self.assertEqual(records, original)

    def test_blank_region_upload_uses_reference_output_and_refreshes_for_populated_geo_areas(self):
        uploads = {"source_upload": Upload(encrypted(blank_region_endo()))}
        with patch("streamlit.file_uploader", side_effect=lambda label, **kwargs: uploads.get(kwargs["key"])), \
                patch("agent_availability_page.render_agent_availability", return_value=None), \
                patch("agent_cycle_page.render_agent_cycle", return_value=None):
            app = AppTest.from_file(str(ROOT / "app.py")).run(timeout=20)
            next(widget for widget in app.text_input if widget.label == "Workbook password").set_value(TEST_PASSWORD).run(timeout=20)
            self.assertFalse(app.exception)
            self.assertFalse(app.error)
            result = app.session_state["result"]
            self.assertTrue(result["new_endo"])
            self.assertFalse(result["geo"])
            self.assertEqual([file.filename for file in result["files"]], ["NEW ENDO L6.csv"])
            self.assertEqual(app.radio[0].value, "Separate NEW ENDO placement groups")
            self.assertIn("new_endo_source_cache", app.session_state)
            uploads["source_upload"] = Upload(source_data(geo=True, tag_header="TAG TO"))
            app.run(timeout=20)
            self.assertFalse(app.exception)
            self.assertFalse(app.error)
            result = app.session_state["result"]
            self.assertFalse(result["new_endo"])
            self.assertTrue(result["geo"])
            self.assertEqual(len(result["files"]), 4)
            app.button(key="clear_all_files").click().run(timeout=20)
            self.assertNotIn("new_endo_source_cache", app.session_state)
            self.assertNotIn("result", app.session_state)


class EncryptedWorkbookTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.plain = source_data()
        cls.protected = encrypted(cls.plain)

    def test_unlocks_in_memory_and_preserves_ordinary_workbooks(self):
        self.assertFalse(needs_workbook_password(self.plain))
        self.assertTrue(needs_workbook_password(self.protected))
        self.assertIs(unlock_workbook(self.plain, ""), self.plain)
        self.assertEqual(unlock_workbook(self.protected, TEST_PASSWORD), self.plain)
        for password, message in (("", "Enter the workbook password"), ("wrong-password", "password is incorrect")):
            with self.subTest(password=password), self.assertRaisesRegex(AutomationError, message):
                unlock_workbook(self.protected, password)

    def test_password_entry_errors_clear_stale_results_and_clear_all_removes_password(self):
        upload = Upload(self.protected)
        uploads = {"source_upload": upload}
        with patch("streamlit.file_uploader", side_effect=lambda label, **kwargs: uploads.get(kwargs["key"])), \
                patch("agent_cycle_page.render_agent_cycle", return_value=None):
            app = AppTest.from_file(str(ROOT / "app.py")).run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertNotIn("result", app.session_state)
            password = next(widget for widget in app.text_input if widget.label == "Workbook password")
            self.assertEqual(password.proto.type, TextInput.PASSWORD)
            password.set_value(TEST_PASSWORD).run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(len(app.error), 0)
            result = app.session_state["result"]
            self.assertEqual(result["count"], 2)
            self.assertEqual(result["source_sheet"], "SPM List")
            self.assertTrue(all(row["Terms"].startswith("SOURCE-L1") for file in result["files"] for row in file.records))
            self.assertFalse(any(heading.value == "RECO agent cycle" for heading in app.subheader))
            password.set_value("wrong-password").run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertNotIn("result", app.session_state)
            self.assertNotIn("geo_autostat_result", app.session_state)
            self.assertIn("password is incorrect", app.error[0].value)
            password.set_value(TEST_PASSWORD).run(timeout=20)
            app.button(key="clear_all_files").click().run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertNotIn("result", app.session_state)
            self.assertFalse(any(key.startswith("source_password_") for key in app.session_state.filtered_state))


class GeoNewEndoPageTests(unittest.TestCase):
    def test_encrypted_tag_to_upload_keeps_source_codes_without_agent_assignment(self):
        uploads = {"source_upload": Upload(encrypted(source_data(geo=True, tag_header="TAG TO")))}
        with patch("streamlit.file_uploader", side_effect=lambda label, **kwargs: uploads.get(kwargs["key"])), \
                patch("agent_availability_page.render_agent_availability", return_value=None) as availability, \
                patch("agent_cycle_page.render_agent_cycle", return_value=None) as rotation:
            app = AppTest.from_file(str(ROOT / "app.py")).run(timeout=20)
            password = next(widget for widget in app.text_input if widget.label == "Workbook password")
            password.set_value(TEST_PASSWORD).run(timeout=20)
            self.assertFalse(app.exception)
            self.assertFalse(app.error)
            rotation.assert_not_called()
            availability.assert_not_called()
            result = app.session_state["result"]
            self.assertEqual(result["source_sheet"], "Sheet1")
            self.assertEqual(result["count"], 2)
            expected = {"0001": "SOURCE-MINDANAO", "0002": "SOURCE-R3"}
            for file in result["files"]:
                for row in csv_rows(file):
                    self.assertEqual(row["CHCode"], expected[row["primary_key"]])
            app.radio[0].set_value("All accounts in one group").run(timeout=20)
            self.assertFalse(app.error)
            self.assertTrue(all(row["CHCode"] == expected[row["primary_key"]]
                                for file in app.session_state["result"]["files"] for row in csv_rows(file)))

    def test_utf8_selection_preserves_names_that_standard_csv_cannot_encode(self):
        workbook = load_workbook(BytesIO(source_data(geo=True)))
        workbook["Sheet1"]["G2"] = "Customer \u674e"
        output = BytesIO()
        workbook.save(output)
        workbook.close()
        upload = Upload(output.getvalue())
        with patch("streamlit.file_uploader", side_effect=lambda label, **kwargs: upload if kwargs["key"] == "source_upload" else None):
            app = AppTest.from_file(str(ROOT / "app.py")).run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertIn("Select CSV UTF-8", app.error[0].value)
            next(widget for widget in app.selectbox if widget.label == "CSV file type").select("CSV UTF-8 (Comma delimited)").run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(len(app.error), 0)
            file = next(file for file in app.session_state["result"]["files"]
                        if file.group_key == "mindanao_blank_cards" and file.channel == "email")
            rows = list(csv.DictReader(StringIO(file.data.decode("utf-8-sig"))))
            self.assertEqual(rows[0]["FullName"], "Customer \u674e")
            self.assertEqual(rows[0]["CHCode"], "SOURCE-MINDANAO")

    def test_geo_upload_splits_mindanao_r3_uses_source_tagging_and_prepares_logs(self):
        upload = Upload(source_data(geo=True))
        with patch("streamlit.file_uploader", side_effect=lambda label, **kwargs: upload if kwargs["key"] == "source_upload" else None), \
                patch("agent_availability_page.render_agent_availability", return_value=None), \
                patch("agent_cycle_page.render_agent_cycle", return_value=None):
            app = AppTest.from_file(str(ROOT / "app.py")).run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(len(app.error), 0)
            result = app.session_state["result"]
            self.assertTrue(result["geo"] and result["grouped"])
            self.assertEqual(result["source_sheet"], "Sheet1")
            self.assertEqual(app.radio[0].value, "Separate Mindanao and R3 groups")
            self.assertEqual({file.group_key for file in result["files"]}, {"mindanao_blank_cards", "r3_blank_cards"})
            self.assertEqual([file.count for file in app.session_state["geo_autostat_result"]["bundle"].files], [2, 2])
            app.radio[0].set_value("All accounts in one group").run(timeout=20)
            self.assertEqual(len(app.error), 0)
            self.assertTrue(all(row["Terms"].startswith("SOURCE-") for file in app.session_state["result"]["files"] for row in file.records))


if __name__ == "__main__":
    unittest.main()
