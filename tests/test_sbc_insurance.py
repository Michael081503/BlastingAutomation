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
from sbc_insurance import build_filtered_datagrid, prepare_insurance, read_insurance_source


ROOT = Path(__file__).resolve().parents[1]
HEADERS = ["Name", "Old I.C.", "Account No.", "Financing/Card No.", "Amount Referred",
           "Collector", "Status Code", "Last Called No."]


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
    def test_status_filter_runs_before_financing_deduplication(self):
        data = source([
            row("FIRST", "00001", "PTP"),
            row("SECOND", "00001", "  rna  "),
            row("THIRD", " 00001 ", "RNA"),
            row("FOURTH", "00002", "BUSY"),
            row("FIFTH", "00003", None),
        ])
        result = prepare_insurance(data, ["RNA", "busy"])
        self.assertEqual([r.record["primary_key"] for r in result.rows], ["SECOND", "FOURTH"])
        self.assertEqual(result.stats["Removed by status"], 2)
        self.assertEqual(result.stats["Duplicate financing/card accounts removed"], 1)
        self.assertEqual(prepare_insurance(data, [""]).rows[0].record["primary_key"], "FIFTH")
        with self.assertRaisesRegex(AutomationError, "Select at least one"):
            prepare_insurance(data, [])

    def test_bad_dimensions_totals_and_blank_financing_match_the_reference(self):
        data = source([
            row("A", "00001", "RNA"),
            row("B", None, "RNA"),
            row("C", "00002", "RNA"),
            row("D", None, "RNA"),
            [None, None, None, None, 1000, None, None, None],
        ], bad_dimensions=True)
        result = prepare_insurance(data, ["RNA"])
        self.assertEqual(len(data.rows), 4)
        self.assertEqual([r.financing for r in result.rows], ["00001", None, "00002"])
        self.assertEqual(result.stats["Selected rows with blank Financing/Card No."], 2)
        self.assertEqual(result.stats["Duplicate financing/card accounts removed"], 1)
        self.assertEqual(data.non_account_rows, 1)

    def test_exports_share_campaign_layout_and_preserve_exact_identifiers(self):
        account = "000123456789012345678901"
        financing = "000999999999999999999999"
        result = prepare_insurance(source([row(account, financing, "RNA", name="Name \u0145", amount=987.65)]), ["RNA"])
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


class InsurancePageTests(unittest.TestCase):
    def test_status_changes_replacement_and_clear_cannot_leave_stale_downloads(self):
        upload = Upload(workbook_bytes([row("A", "0001", "RNA"), row("B", "0001", "BUSY")]))

        def uploader(label, **kwargs):
            return upload if kwargs["key"] == "sbc_insurance_upload" else None

        with patch("streamlit.file_uploader", side_effect=uploader):
            app = AppTest.from_file(str(ROOT / "pages" / "SBC_INSURANCE.py")).run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertNotIn("sbc_insurance_result", app.session_state)
            app.multiselect(key="sbc_insurance_statuses").select("BUSY").run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(app.session_state["sbc_insurance_result"]["result"].file.records[0]["primary_key"], "B")
            app.multiselect(key="sbc_insurance_statuses").set_value(["RNA"]).run(timeout=20)
            self.assertEqual(app.session_state["sbc_insurance_result"]["result"].file.records[0]["primary_key"], "A")
            upload = Upload(workbook_bytes([row("NEW", "0002", "RNA")]))
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
            self.assertEqual(app.session_state["_retained_upload_settings"]["sbc_insurance_statuses"], ["RNA"])


if __name__ == "__main__":
    unittest.main()
