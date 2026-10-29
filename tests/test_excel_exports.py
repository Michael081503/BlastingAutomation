from __future__ import annotations

import csv
import unittest
from io import BytesIO, StringIO
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from streamlit.testing.v1 import AppTest

from automation import create_default_csv_template
from autostat import SourceSelection, build_autostat_exports, inspect_contact_source
from exports import build_excel_zip, build_export_excel, build_sheet_exports


class ExcelExportTests(unittest.TestCase):
    def test_excel_download_preserves_full_rows_and_literal_identifiers_through_autostat(self):
        accounts = ["5156041234567897", "001234567890123456", "123456789012345678901234", "=2+2",
                    *[f"000000{index:012d}" for index in range(110)]]
        records = [{
            "primary_key": f"KEY-{index}", "cardnumber": account, "indexoneuid": f"UID-{index}",
            "email": f"account{index}@example.com", "phone": "09171234567", "OB": index,
        } for index, account in enumerate(accounts)]
        for encoding in ("cp1252", "utf-8-sig"):
            with self.subTest(encoding=encoding):
                bundle = build_sheet_exports(
                    create_default_csv_template(), records, "", 1, base_name="Accounts", label="Accounts",
                    template_extension=".csv", encoding=encoding,
                )
                original_csvs = [file.data for file in bundle.files]
                with ZipFile(BytesIO(build_excel_zip(bundle.files))) as archive:
                    self.assertEqual(archive.namelist(), ["Accounts_Email.xlsx", "Accounts_Viber.xlsx"])
                    for file in bundle.files:
                        excel_data = archive.read(Path(file.filename).with_suffix(".xlsx").name)
                        workbook = load_workbook(BytesIO(excel_data))
                        sheet = workbook.active
                        raw_rows = list(csv.reader(StringIO(file.data.decode(encoding))))
                        self.assertEqual(list(sheet.values), [tuple(value or None for value in row) for row in raw_rows])
                        self.assertEqual(sheet.max_row, len(accounts) + 1)
                        for row in sheet.iter_rows(min_row=2):
                            for column in ("primary_key", "cardnumber", "indexoneuid"):
                                cell = row[file.headers.index(column)]
                                self.assertEqual(cell.data_type, "s")
                                self.assertEqual(cell.number_format, "@")
                        workbook.close()
                        selection = SourceSelection("Accounts.xlsx", excel_data, inspect_contact_source(excel_data, ".xlsx")[0])
                        logs = build_autostat_exports([selection], timestamp=None)
                        log = next(log for log in logs.files if log.count)
                        actual = list(csv.DictReader(StringIO(log.data.decode("cp1252"))))
                        self.assertEqual([row["Account No."] for row in actual], list(reversed(accounts)))
                        self.assertEqual(sum(row["Account precision issues"] for row in logs.summary), 0)
                self.assertEqual([file.data for file in bundle.files], original_csvs)

    def test_main_page_offers_excel_zip_and_individual_workbooks_without_building_them_on_rerun(self):
        # A real source workbook exercises the page's normal preparation flow.
        workbook = Workbook()
        workbook.active.append(["cardnumber", "email", "OB"])
        workbook.active.append(["5156041234567897", "first@example.com", 1000])
        output = BytesIO()
        workbook.save(output)
        workbook.close()
        upload = BytesIO(output.getvalue())
        upload.name = "accounts.xlsx"
        root = Path(__file__).resolve().parents[1]
        with patch("streamlit.file_uploader", side_effect=lambda label, **kw: upload if kw.get("key") == "source_upload" else None), \
                patch("exports.build_excel_zip", wraps=build_excel_zip) as zip_builder, \
                patch("exports.build_export_excel", wraps=build_export_excel) as excel_builder, \
                patch("streamlit.download_button") as download:
            app = AppTest.from_file(str(root / "app.py")).run(timeout=20)
            next(button for button in app.button if button.label == "Preview and prepare files").click().run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(len(app.error), 0)
            zip_builder.assert_not_called()
            excel_builder.assert_not_called()
            zip_call = next(call for call in download.call_args_list if call.kwargs.get("key") == "excel_zip")
            self.assertEqual(zip_call.kwargs["type"], "primary")
            self.assertTrue(callable(zip_call.args[1]))
            self.assertTrue(zip_call.args[2].endswith("_Excel.zip"))
            with ZipFile(BytesIO(zip_call.args[1]())) as archive:
                self.assertEqual(len(archive.namelist()), 2)
                self.assertTrue(all(name.endswith(".xlsx") for name in archive.namelist()))
            single = next(call for call in download.call_args_list if call.kwargs.get("key", "").startswith("excel_download_"))
            workbook = load_workbook(BytesIO(single.args[1]()))
            headers = [cell.value for cell in workbook.active[1]]
            cell = workbook.active.cell(2, headers.index("cardnumber") + 1)
            self.assertEqual(cell.value, "5156041234567897")
            self.assertEqual(cell.data_type, "s")
            self.assertEqual(cell.number_format, "@")
            workbook.close()


if __name__ == "__main__":
    unittest.main()
