from __future__ import annotations

import csv
import unittest
from datetime import datetime
from io import BytesIO, StringIO
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table

from automation import (
    DEFAULT_TEMPLATE_HEADERS, create_default_csv_template, fill_csv_template,
    fill_excel_template, inspect_workbook, read_source,
)
from autostat import SourceSelection, build_autostat_exports, inspect_contact_source
from b2_autostat import B2ActivitySource, b2_metadata_columns, build_b2_autostat_exports
from b2_matching import build_b2_exports, build_b4_exports
from exports import (
    build_csv_zip, build_excel_zip, build_export_csv, build_export_excel,
    build_geo_exports, build_negative_exports, source_account_counts, validate_source_exports,
)


LEGACY_HEADERS = [
    "email", "primary_key", "indexoneuid", "cardnumber", "FullName", "phone", "OB",
    "Monthly", "Terms", "doclass", "Level", "Product", "EndoDate", "AREA",
    "ChristmasAmnesty", "NewEndo",
]
EXPECTED_HEADERS = [
    "email", "primary_key", "indexoneuid", "Terms", "FullName", "phone", "OB",
    "Monthly", "CHCode", "doclass", "Level", "Promo", "EndoDate", "AREA",
    "ChristmasAmnesty", "NewEndo",
]


def csv_data(headers, rows=()):
    output = StringIO(newline="")
    writer = csv.writer(output)
    writer.writerow(headers)
    writer.writerows(rows)
    return output.getvalue().encode("cp1252")


def excel_data(headers, rows=(), *, table=False):
    workbook = Workbook()
    workbook.active.title = "Accounts"
    workbook.active.append(headers)
    for row in rows:
        workbook.active.append(row)
    if table:
        workbook.active.add_table(Table(displayName="Campaign", ref="A1:P2"))
        workbook.create_sheet("Notes").append(["Terms", "Product", "cardnumber"])
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def selection(name, data):
    return SourceSelection(name, data, inspect_contact_source(data, name[name.rfind("."):])[0])


class CampaignHeaderTests(unittest.TestCase):
    def test_legacy_and_new_templates_keep_every_value_and_column_position(self):
        record = {
            "email": "one@example.com", "primary_key": "000001", "indexoneuid": "0002",
            "cardnumber": "00001234567890123456", "FullName": "One, Example", "phone": "09171234567",
            "OB": 1234.5, "Monthly": "Agent Name", "Terms": "AGENTCODE", "level": 1,
            "product": "CARDS", "endodate": datetime(2026, 11, 10), "AREA": "Mindanao",
            "ChristmasAmnesty": "QUALIFIED",
        }
        self.assertEqual(list(DEFAULT_TEMPLATE_HEADERS), EXPECTED_HEADERS)
        outputs = []
        for headers in (LEGACY_HEADERS, EXPECTED_HEADERS, [
            "cardbnumber" if header == "cardnumber" else header for header in LEGACY_HEADERS
        ]):
            with self.subTest(headers=headers):
                data = fill_csv_template(csv_data(headers), [record])
                rows = list(csv.reader(StringIO(data.decode("cp1252"))))
                self.assertEqual(rows[0], EXPECTED_HEADERS)
                self.assertEqual(rows[1][3], record["cardnumber"])
                self.assertEqual(rows[1][8], "AGENTCODE")
                self.assertEqual(rows[1][11], "CARDS")
                outputs.append(rows[1])
        self.assertTrue(all(row == outputs[0] for row in outputs))
        reordered = list(reversed(LEGACY_HEADERS)) + ["Extra"]
        rows = list(csv.reader(StringIO(fill_csv_template(csv_data(reordered), [record]).decode("cp1252"))))
        self.assertEqual(rows[0], list(reversed(EXPECTED_HEADERS)) + ["Extra"])
        self.assertEqual(rows[1], list(reversed(outputs[0])) + [""])
        workbook = load_workbook(BytesIO(fill_excel_template(
            excel_data(LEGACY_HEADERS, [["stale"] * 16], table=True), [record], "Accounts", 1,
        )))
        self.assertEqual([cell.value for cell in workbook.active[1]], EXPECTED_HEADERS)
        self.assertEqual([column.name for column in workbook.active.tables["Campaign"].tableColumns], EXPECTED_HEADERS)
        self.assertEqual((workbook.active["D2"].value, workbook.active["D2"].number_format),
                         (record["cardnumber"], "@"))
        self.assertEqual(list(workbook["Notes"].values), [("Terms", "Product", "cardnumber")])
        workbook.close()

    def test_geo_reco_csv_excel_zip_and_empty_channels_use_renamed_headers(self):
        record = {"primary_key": "0001", "cardnumber": "00001234567890123456", "indexoneuid": "0002",
                  "email": "one@example.com", "OB": 500, "product": "CARDS", "level": 1, "AREA": "Mindanao"}
        for builder in (build_geo_exports, build_negative_exports):
            bundle = builder(csv_data(LEGACY_HEADERS), [record], "", 1, ".csv", agents=(("Agent Name", "AGENTCODE"),))
            validate_source_exports(bundle.files, source_account_counts([record]), source_sheet="NEGATIVE")
            with ZipFile(BytesIO(build_csv_zip(bundle.files))) as csv_zip, \
                    ZipFile(BytesIO(build_excel_zip(bundle.files))) as excel_zip:
                for file in bundle.files:
                    rows = list(csv.DictReader(StringIO(file.data.decode("cp1252"))))
                    self.assertEqual(file.headers, EXPECTED_HEADERS)
                    protected = build_export_csv(file)
                    self.assertEqual(csv_zip.read(file.filename), protected)
                    workbook = load_workbook(BytesIO(build_export_excel(file)))
                    self.assertEqual([cell.value for cell in workbook.active[1]], EXPECTED_HEADERS)
                    self.assertEqual(excel_zip.read(file.filename.replace(".csv", ".xlsx"))[:2], b"PK")
                    if file.channel == "email":
                        self.assertEqual(rows[0]["Terms"], record["cardnumber"])
                        self.assertEqual(rows[0]["CHCode"], file.records[0]["Terms"])
                        self.assertEqual(rows[0]["Promo"], "CARDS")
                        protected_row = next(csv.DictReader(StringIO(protected.decode("cp1252"))))
                        self.assertEqual(protected_row["Terms"], "\t" + record["cardnumber"])
                        self.assertEqual(protected_row["CHCode"], rows[0]["CHCode"])
                        self.assertEqual((workbook.active["D2"].value, workbook.active["D2"].number_format),
                                         (record["cardnumber"], "@"))
                    else:
                        self.assertEqual(rows, [])
                        self.assertEqual(workbook.active.max_row, 1)
                    workbook.close()

    def test_renamed_exports_and_legacy_exports_produce_identical_autostat_logs(self):
        values = ["one@example.com", "000001", "0002", "00001234567890123456", "Example", "09171234567",
                  "500", "Agent Name", "AGENTCODE", "", "1", "CARDS", "", "Mindanao", "", ""]
        for extension, encoder in ((".csv", csv_data), (".xlsx", excel_data)):
            legacy = selection("campaign" + extension, encoder(LEGACY_HEADERS, [values]))
            renamed = selection("campaign" + extension, encoder(EXPECTED_HEADERS, [values]))
            self.assertEqual(renamed.sheet.headers[renamed.sheet.account_column], "Terms")
            old_logs = build_autostat_exports([legacy], timestamp=None)
            new_logs = build_autostat_exports([renamed], timestamp=None)
            self.assertEqual([file.data for file in new_logs.files], [file.data for file in old_logs.files])
            legacy_activity = B2ActivitySource(legacy, *b2_metadata_columns(legacy.sheet.headers))
            renamed_activity = B2ActivitySource(renamed, *b2_metadata_columns(renamed.sheet.headers))
            for bucket in (2, 4):
                old_logs = build_b2_autostat_exports([legacy_activity], timestamp=None, bucket=bucket)
                new_logs = build_b2_autostat_exports([renamed_activity], timestamp=None, bucket=bucket)
                self.assertEqual([file.data for file in new_logs.files], [file.data for file in old_logs.files])
            if extension == ".xlsx":
                info = inspect_workbook(renamed.data)[0]
                mapped = read_source(renamed.data, info.name, info.header_row, allow_missing=True)
                self.assertEqual(mapped.records[0]["cardnumber"], values[3])
                self.assertEqual(mapped.records[0]["product"], "CARDS")

    def test_b2_b4_headers_keep_accounts_collectors_and_products_separate(self):
        for bucket, builder in ((2, build_b2_exports), (4, build_b4_exports)):
            database = selection("database.xlsx", excel_data(
                ["Account Key", "Contract Number", "Bucket Code", "Product Category Code", "Email (01)", "Mobile 1"],
                [["000001", "00001234567890123456", bucket, "MC", "one@example.com", "09171234567"]],
            ))
            grid = selection("list.xlsx", excel_data(
                ["Account No.", "Financing/Card No."], [["000001", "00001234567890123456"]],
            ))
            bundle = builder(database, grid).bundle
            for file in bundle.files:
                self.assertEqual(list(file.headers), EXPECTED_HEADERS)
                if file.count:
                    row = next(csv.DictReader(StringIO(file.data.decode("cp1252"))))
                    self.assertEqual(row["Terms"], "00001234567890123456")
                    self.assertEqual(row["CHCode"], "AMRANCELI")
                    self.assertEqual(row["Promo"], "CARDS")
                    activity_source = selection(file.filename, build_export_csv(file))
                    logs = build_b2_autostat_exports([
                        B2ActivitySource(activity_source, *b2_metadata_columns(activity_source.sheet.headers)),
                    ], timestamp=None, bucket=bucket)
                    email = next(csv.DictReader(StringIO(logs.files[0].data.decode("cp1252"))))
                    self.assertEqual(email["Account No."], "000001")
                    self.assertEqual(email["Collector"], "AMRANCELI")
