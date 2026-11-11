from __future__ import annotations

import csv
from io import BytesIO, StringIO
from pathlib import Path
import unittest
from unittest.mock import patch
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl import Workbook, load_workbook
from streamlit.testing.v1 import AppTest

from automation import AutomationError, DEFAULT_TEMPLATE_HEADERS
from exports import build_export_csv, build_export_excel
from sbc_insurance import (
    SBC_INSURANCE_INCLUDED_STATUSES,
    build_filtered_datagrid, inspect_insurance_masterlist, prepare_insurance,
    read_insurance_masterlist, read_insurance_source,
)


ROOT = Path(__file__).resolve().parents[1]
HEADERS = ["Name", "Old I.C.", "Account No.", "Financing/Card No.", "Amount Referred",
           "Collector", "Status Code", "Last Called No."]
MASTER_HEADERS = ["PN_Number", "EMAIL", "MOBILE", "ADDRESS1", "ADDRESS2", "ADDRESS3"]


def workbook_bytes(rows, headers=HEADERS, *, bad_dimensions=False):
    workbook = Workbook()
    workbook.active.title = "Insurance"
    workbook.active.append(headers)
    for row in rows:
        workbook.active.append(row)
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    if not bad_dimensions:
        return output.getvalue()
    fixed = BytesIO()
    with ZipFile(BytesIO(output.getvalue())) as original, ZipFile(fixed, "w", ZIP_DEFLATED) as target:
        for name in original.namelist():
            data = original.read(name)
            if name == "xl/worksheets/sheet1.xml":
                import re
                data = re.sub(rb'<dimension ref="[^"]+"', b'<dimension ref="A1:C1"', data)
            target.writestr(name, data)
    return fixed.getvalue()


def source(rows, headers=HEADERS, **kwargs):
    return read_insurance_source(workbook_bytes(rows, headers, **kwargs), "Insurance", 1)


def row(account, financing, status, *, name="Person", amount=123.5):
    return [name, "0009", account, financing, amount, "TESTCOLLECTOR", status, "09171234567"]


class InsuranceTests(unittest.TestCase):
    def test_automatic_status_filter_runs_before_financing_deduplication(self):
        data = source([
            row("FIRST", "00001", "PTP"),
            row("SECOND", "00001", "  rna  "),
            row("THIRD", " 00001 ", "RNA"),
            row("FOURTH", "00002", "BUSY"),
            row("FIFTH", "00003", None),
        ])
        result = prepare_insurance(data)
        self.assertEqual([r.record["primary_key"] for r in result.rows], ["SECOND", "FOURTH", "FIFTH"])
        self.assertEqual(result.stats["Removed by status"], 1)
        self.assertEqual(result.stats["Duplicate financing/card accounts removed"], 1)

    def test_ineligible_and_unseen_statuses_do_not_fall_back_to_exporting_everything(self):
        statuses = ["OUTGOING CALL - PTP PAID", "OUTGOING CALL - PTP FULLYPAID", "BP", "A NEW UNKNOWN STATUS"]
        result = prepare_insurance(source([row(str(i), f"F{i}", status) for i, status in enumerate(statuses)]))
        self.assertEqual(result.rows, [])
        self.assertEqual(result.stats["Removed by status"], len(statuses))
        self.assertEqual(len(result.file.records), 0)
        self.assertEqual(len(SBC_INSURANCE_INCLUDED_STATUSES), 23)

    def test_bad_dimensions_totals_and_blank_financing_match_the_reference(self):
        data = source([
            row("A", "00001", "RNA"),
            row("B", None, "RNA"),
            row("C", "00002", "RNA"),
            row("D", None, "RNA"),
            [None, None, None, None, 1000, None, None, None],
        ], bad_dimensions=True)
        result = prepare_insurance(data)
        self.assertEqual(len(data.rows), 4)
        self.assertEqual([r.financing for r in result.rows], ["00001", None, "00002"])
        self.assertEqual(result.stats["Selected rows with blank Financing/Card No."], 2)
        self.assertEqual(result.stats["Duplicate financing/card accounts removed"], 1)
        self.assertEqual(data.non_account_rows, 1)

    def test_exports_share_campaign_layout_and_preserve_exact_identifiers(self):
        account = "000123456789012345678901"
        financing = "000999999999999999999999"
        result = prepare_insurance(source([row(account, financing, "RNA", name="Name \u0145", amount=987.65)]))
        csv_rows = list(csv.DictReader(StringIO(build_export_csv(result.file).decode("utf-8-sig"))))
        self.assertEqual(tuple(csv_rows[0]), DEFAULT_TEMPLATE_HEADERS)
        record = csv_rows[0]
        self.assertEqual(record["primary_key"], "\t" + account)
        self.assertEqual(record["Terms"], "\t" + financing)
        self.assertEqual(record["indexoneuid"], "0009")
        self.assertEqual(record["CHCode"], "TESTCOLLECTOR")
        self.assertEqual(record["doclass"], "SBC INSURANCE")
        self.assertEqual((record["email"], record["phone"], record["Monthly"]), ("", "", ""))
        self.assertEqual(record["FullName"], "Name \u0145")
        workbook = load_workbook(BytesIO(build_export_excel(result.file)))
        sheet = workbook.active
        self.assertEqual(tuple(c.value for c in sheet[1]), DEFAULT_TEMPLATE_HEADERS)
        self.assertEqual((sheet["B2"].value, sheet["B2"].data_type, sheet["B2"].number_format), (account, "s", "@"))
        self.assertEqual((sheet["D2"].value, sheet["D2"].number_format), (financing, "@"))
        self.assertEqual((sheet["G2"].value, sheet["G2"].data_type), (987.65, "n"))
        self.assertEqual(sheet["J2"].value, "SBC INSURANCE")
        workbook.close()
        workbook = load_workbook(BytesIO(build_filtered_datagrid(result)))
        self.assertEqual([c.value for c in workbook.active[1]], HEADERS)
        self.assertEqual(workbook.active["D2"].value, financing)
        self.assertEqual(workbook.active["G2"].value, "RNA")
        workbook.close()

    def test_missing_ambiguous_or_invalid_status_cannot_bypass_filtering(self):
        with self.assertRaisesRegex(AutomationError, "Missing Status"):
            source([row("A", "F", "RNA")], [*HEADERS[:6], "Other", HEADERS[7]])
        with self.assertRaisesRegex(AutomationError, "Ambiguous Status"):
            source([row("A", "F", "RNA") + ["RNA"]], HEADERS + ["STATUS"])
        with self.assertRaisesRegex(AutomationError, "formula has no saved value"):
            source([row("A", "F", '=IF(1=1,"RNA","BUSY")')])
        with self.assertRaisesRegex(AutomationError, "scientific notation"):
            source([row("A", "1.23E+20", "RNA")])


class Upload(BytesIO):
    name = "insurance.xlsx"


class InsuranceMasterlistTests(unittest.TestCase):
    def test_exact_reference_uses_the_kept_account_contacts_and_preserves_blank_financing(self):
        master_headers = ["IS Ref No#", *MASTER_HEADERS, "IS Ref No#"]
        master = read_insurance_masterlist(workbook_bytes([
            ["REFA", "PN1", "old@example.com", "09171234567", "Old street", None, None, "REFA"],
            ["REFB", "PN1", "kept@example.com", "09181234567", "Kept street", None, None, "REFB"],
            ["REFC", "PN2", "blank@example.com", None, "Blank financing street", None, None, "REFC"],
        ], master_headers), "Insurance", 1)
        result = prepare_insurance(source([
            row("REFA", "PN1", "PTP"), row("REFB", "PN1", "RNA"), row("REFC", None, "RNA"),
        ]), masterlist=master)
        self.assertEqual([r["email"] for r in result.file.records], ["kept@example.com", "blank@example.com"])
        self.assertEqual(result.contacts[0].address[0], "Kept street")
        self.assertIsNone(result.file.records[1]["cardnumber"])
        self.assertEqual(result.stats["Accounts matched by IS Ref No#"], 2)
        self.assertEqual(result.stats["Accounts matched by PN_Number fallback"], 0)

    def test_conflicting_reference_identifiers_block_contacts_without_trying_another_account(self):
        master = read_insurance_masterlist(workbook_bytes([
            ["A", "PN2", "wrong@example.com", None, None, None, None],
            ["B", "PN1", "fallback@example.com", None, None, None, None],
            ["C", "PN3", "one@example.com", None, None, None, None],
            ["C", "PN4", "two@example.com", None, None, None, None],
        ], ["IS Ref No#", *MASTER_HEADERS]), "Insurance", 1)
        result = prepare_insurance(source([row("A", "PN1", "RNA"), row("C", "PN3", "RNA")]), masterlist=master)
        self.assertEqual(result.stats["Accounts with conflicting masterlist identifiers"], 2)
        self.assertTrue(all(record["email"] is None for record in result.file.records))

    def test_contacts_join_after_filtering_without_adding_or_mutating_source_accounts(self):
        original = source([
            row("EXCLUDED", "0001", "PTP"), row("KEPT", "0001", "RNA"), row("DUPLICATE", "0001", "RNA"),
            row("SECOND", "0002", "RNA"), row("UNMATCHED", "0003", "RNA"), row("BLANK", None, "RNA"),
        ])
        master = read_insurance_masterlist(workbook_bytes([
            ["0001", "bad email", "123", None, None, None],
            ["0001", "first@example.com", "+639171234567", "Street 1", "City 1", "Region 1"],
            ["0001", "FIRST@example.com", "09171234567", "Street 1", "City 1", "Region 1"],
            ["0001", "other@example.com", "09181234567", "Street 2", "City 2", "Region 2"],
            ["0002", "second@example.com", None, "Second street", None, None],
            ["UNRELATED", "unrelated@example.com", "09191234567", "Other street", None, None],
        ], MASTER_HEADERS), "Insurance", 1)
        result = prepare_insurance(original, masterlist=master)
        self.assertEqual([record["primary_key"] for record in result.file.records], ["KEPT", "SECOND", "UNMATCHED", "BLANK"])
        first = result.file.records[0]
        self.assertEqual((first["email"], first["phone"]), ("first@example.com", "9171234567"))
        self.assertEqual(result.contacts[0].address, ("Street 1", "City 1", "Region 1"))
        self.assertEqual(result.contacts[0].source_rows, 4)
        self.assertEqual(result.stats["Accounts matched to masterlist"], 2)
        self.assertEqual(result.stats["Accounts matched by PN_Number fallback"], 2)
        self.assertEqual(result.stats["Accounts without a masterlist match"], 2)
        self.assertEqual(result.stats["Matched accounts with alternative valid emails"], 1)
        self.assertEqual(result.stats["Matched accounts with alternative valid mobiles"], 1)
        self.assertEqual(result.stats["Matched accounts with alternative postal addresses"], 1)
        self.assertEqual(result.stats["Invalid masterlist email entries"], 1)
        self.assertTrue(all(not row.record["email"] and not row.record["phone"] for row in original.rows))
        self.assertTrue(all(not r["email"] and not r["phone"] for r in prepare_insurance(original).file.records))

    def test_addresses_are_available_without_changing_the_campaign_columns(self):
        financing = "000123456789012345678901"
        master = read_insurance_masterlist(workbook_bytes([
            [financing, "person@example.com", "09171234567", "Street", "City", "Province"],
        ], MASTER_HEADERS, bad_dimensions=True), "Insurance", 1)
        result = prepare_insurance(source([row("ACCOUNT", financing, "RNA")]), masterlist=master)
        records = list(csv.DictReader(StringIO(build_export_csv(result.file).decode("utf-8-sig"))))
        self.assertEqual(tuple(records[0]), DEFAULT_TEMPLATE_HEADERS)
        self.assertEqual(records[0]["email"], "person@example.com")
        self.assertEqual(records[0]["phone"], "\t9171234567")
        workbook = load_workbook(BytesIO(build_export_excel(result.file)))
        self.assertEqual(workbook.active["F2"].value, "9171234567")
        self.assertEqual(workbook.active["F2"].number_format, "@")
        self.assertEqual(workbook.active.max_column, 16)
        workbook.close()
        workbook = load_workbook(BytesIO(build_filtered_datagrid(result)))
        self.assertEqual([c.value for c in workbook.active[1]], HEADERS)
        workbook.close()
        workbook = load_workbook(BytesIO(build_filtered_datagrid(result, include_masterlist=True)))
        self.assertEqual([c.value for c in workbook.active[1]], [*HEADERS, "Masterlist EMAIL", "Masterlist MOBILE",
                                                               "Masterlist ADDRESS1", "Masterlist ADDRESS2", "Masterlist ADDRESS3"])
        self.assertEqual([c.value for c in workbook.active[2]][-5:], ["person@example.com", "9171234567", "Street", "City", "Province"])
        self.assertEqual(workbook.active["J2"].number_format, "@")
        workbook.close()

    def test_invalid_identifiers_and_contact_errors_do_not_create_guessed_matches(self):
        master = read_insurance_masterlist(workbook_bytes([
            ["000123", "=1+1", "#VALUE!", None, None, None],
            ["000123", "valid@example.com", "09171234567", "Street", None, None],
            ["1.23E+20", "wrong@example.com", None, None, None, None],
            [1234567890123456, "wrong@example.com", None, None, None, None],
            ["=1+1", "wrong@example.com", None, None, None, None],
        ], MASTER_HEADERS), "Insurance", 1)
        self.assertEqual(set(master.accounts), {"000123"})
        self.assertEqual(master.accounts["000123"].email, "valid@example.com")
        self.assertEqual(master.stats["Masterlist rows with unusable PN_Number"], 3)
        self.assertEqual(master.stats["Masterlist contact cells with Excel errors or uncached formulas"], 2)
        result = prepare_insurance(source([row("000123", "123", "RNA")]), masterlist=master)
        self.assertEqual(result.stats["Accounts matched to masterlist"], 0)
        self.assertIsNone(result.file.records[0]["email"])

    def test_sheet_detection_ignores_summary_and_duplicate_unused_reference_headers(self):
        workbook = Workbook()
        workbook.active.title = "SUMMARY"
        workbook.active.append(["Summary only"])
        sheet = workbook.create_sheet("LIST")
        sheet.append(["Title"])
        sheet.append(["IS Ref No#", "PN_Number", "EMAIL", "IS Ref No#", "ADDRESS1"])
        sheet.append([1, "PN001", "one@example.com", 1, "Street"])
        output = BytesIO()
        workbook.save(output)
        workbook.close()
        info = inspect_insurance_masterlist(output.getvalue())
        self.assertEqual([(s.name, s.header_row) for s in info], [("LIST", 2)])
        master = read_insurance_masterlist(output.getvalue(), "LIST", 2)
        self.assertEqual(master.accounts["pn001"].email, "one@example.com")
        with self.assertRaisesRegex(AutomationError, "PN_Number and EMAIL"):
            inspect_insurance_masterlist(workbook_bytes([row("A", "F", "RNA")]))


class InsurancePageTests(unittest.TestCase):
    def test_masterlist_replacement_removal_and_clear_rebuild_contact_downloads(self):
        upload = Upload(workbook_bytes([row("A", "0001", "RNA")]))
        master = Upload(workbook_bytes([["0001", "first@example.com", "09171234567", "Street", None, None]], MASTER_HEADERS))
        keys = []

        def uploader(label, **kwargs):
            keys.append(kwargs["key"])
            return {"sbc_insurance_upload": upload, "sbc_insurance_masterlist_upload": master}.get(kwargs["key"])

        with patch("streamlit.file_uploader", side_effect=uploader):
            app = AppTest.from_file(str(ROOT / "pages" / "SBC_INSURANCE.py")).run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(len(app.multiselect), 0)
            self.assertEqual(app.session_state["sbc_insurance_result"]["result"].file.records[0]["email"], "first@example.com")
            master = Upload(workbook_bytes([["0001", "new@example.com", None, "New street", None, None]], MASTER_HEADERS))
            app.run(timeout=20)
            self.assertEqual(app.session_state["sbc_insurance_result"]["result"].file.records[0]["email"], "new@example.com")
            master = None
            app.run(timeout=20)
            self.assertIsNone(app.session_state["sbc_insurance_result"]["result"].file.records[0]["email"])
            master = Upload(workbook_bytes([["0001", "new@example.com", None, "New street", None, None]], MASTER_HEADERS))
            app.run(timeout=20)
            self.assertEqual(app.session_state["sbc_insurance_result"]["result"].file.records[0]["email"], "new@example.com")
            app.button(key="clear_all_files").click().run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertNotIn("sbc_insurance_result", app.session_state)
            self.assertNotIn("sbc_insurance_source_info", app.session_state)
            self.assertIn("sbc_insurance_masterlist_upload_clear_1", keys)
            self.assertEqual(len(app.get("download_button")), 0)

    def test_upload_automatically_creates_output_and_ignores_old_status_selections(self):
        upload = Upload(workbook_bytes([row("A", "0001", "RNA"), row("B", "0001", "BUSY")]))

        def uploader(label, **kwargs):
            return upload if kwargs["key"] == "sbc_insurance_upload" else None

        with patch("streamlit.file_uploader", side_effect=uploader):
            app = AppTest.from_file(str(ROOT / "pages" / "SBC_INSURANCE.py"))
            app.session_state["sbc_insurance_statuses"] = ["BUSY"]
            app.session_state["_retained_upload_settings"] = {"sbc_insurance_statuses": ["BUSY"]}
            app.run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(len(app.multiselect), 0)
            self.assertEqual(len(app.get("download_button")), 2)
            self.assertEqual(app.get("download_button")[0].label, "Download SBC INSURANCE Output.xlsx")
            self.assertEqual(app.session_state["sbc_insurance_result"]["result"].file.records[0]["primary_key"], "A")
            self.assertNotIn("sbc_insurance_statuses", app.session_state)
            upload = Upload(workbook_bytes([row("PAID", "0001", "OUTGOING CALL - PTP FULLYPAID"), row("NEW", "0002", "RNA")]))
            app.run(timeout=20)
            self.assertEqual(app.session_state["sbc_insurance_result"]["result"].file.records[0]["primary_key"], "NEW")
            app.selectbox(key="sbc_insurance_output_format").select("Excel (.xlsx)").run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            app.button(key="clear_all_files").click().run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertNotIn("sbc_insurance_result", app.session_state)
            self.assertNotIn("sbc_insurance_source_info", app.session_state)
            self.assertEqual(len(app.get("download_button")), 0)
            self.assertEqual(app.selectbox(key="sbc_insurance_output_format").value, "Excel (.xlsx)")
            self.assertNotIn("sbc_insurance_statuses", app.session_state["_retained_upload_settings"])


if __name__ == "__main__":
    unittest.main()
