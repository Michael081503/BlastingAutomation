from __future__ import annotations

import csv
import unittest
from io import BytesIO, StringIO

from openpyxl import Workbook, load_workbook

from automation import create_default_csv_template, read_source
from autostat import SourceSelection, build_autostat_channel_excel, inspect_contact_source
from b2_autostat import B2ActivitySource, build_b2_autostat_exports, build_b4_autostat_exports
from b2_matching import build_b2_exports, build_b4_exports
from exports import build_csv_zip, build_export_csv, build_export_excel, build_geo_exports, build_negative_exports
from phone_validation import first_valid_phone, validated_phone
from zipfile import ZipFile


def workbook_data(headers, rows):
    book = Workbook()
    book.active.title = "NEGATIVE"
    book.active.append(headers)
    for row in rows:
        book.active.append(row)
    output = BytesIO()
    book.save(output)
    book.close()
    return output.getvalue()


def selection(data, name):
    return SourceSelection(name, data, inspect_contact_source(data, ".xlsx")[0])


def csv_rows(file):
    return list(csv.DictReader(StringIO(file.data.decode("utf-8-sig" if file.data.startswith(b"\xef\xbb\xbf") else "cp1252"))))


class PhoneNormalizationTests(unittest.TestCase):
    def test_rightmost_ten_digits_cover_prefixes_formatting_and_numeric_cells(self):
        for source in ("09171234567", "+639171234567", "00639171234567", "+63 (917) 123-4567",
                       "9171234567", 9171234567, 9171234567.0, "9171234567.0", "9.171234567E9"):
            with self.subTest(source=source):
                self.assertEqual(validated_phone(source), "9171234567")
        for source in (None, "", "123", "0281234567", "phone9171234567", "09171234567 ext 12", True, float("nan")):
            with self.subTest(source=source):
                self.assertIsNone(validated_phone(source))
        self.assertEqual(first_valid_phone("invalid / +639171234567; 09181234567"), "9171234567")

    def test_reco_geo_campaign_previews_csv_excel_and_zip_use_ten_digit_phones(self):
        data = workbook_data(
            ["primary_key", "cardnumber", "email", "phone", "product", "level", "ChristmasAmnesty", "Area"],
            [["00001234567890123456", "00009876543210987654", "one@example.com", "+639171234567", "MC", 1, None, "Mindanao"],
             ["00001234567890123457", "00009876543210987655", "two@example.com", "09181234567", "MC", 6, None, "Region 3"]],
        )
        mapped = read_source(data, "NEGATIVE", 1, allow_missing=True)
        self.assertEqual([record["phone"] for record in mapped.records], ["9171234567", "9181234567"])
        for builder in (build_negative_exports, build_geo_exports):
            with self.subTest(builder=builder.__name__):
                bundle = builder(create_default_csv_template(), mapped.records, "", 1, ".csv")
                for file in bundle.files:
                    if file.channel != "viber":
                        continue
                    for row in csv_rows(file):
                        self.assertRegex(row["phone"], r"^9\d{9}$")
                        self.assertTrue(row["primary_key"].startswith("0000123456789012345"))
                        self.assertTrue(row["Terms"].startswith("0000987654321098765"))
                    self.assertTrue(all(len(record["phone"]) == 10 for record in file.preview))
                    csv_output = list(csv.DictReader(StringIO(build_export_csv(file).decode("cp1252"))))
                    self.assertTrue(all(len(row["phone"].strip()) == 10 for row in csv_output))
                    book = load_workbook(BytesIO(build_export_excel(file)))
                    try:
                        rows = list(book.active.values)
                        column = rows[0].index("phone")
                        self.assertTrue(all(len(row[column]) == 10 and row[column].startswith("9") for row in rows[1:]))
                    finally:
                        book.close()
                with ZipFile(BytesIO(build_csv_zip(bundle.files))) as archive:
                    for name in archive.namelist():
                        if "Viber" in name:
                            rows = csv.DictReader(StringIO(archive.read(name).decode("cp1252")))
                            self.assertTrue(all(len(row["phone"].strip()) == 10 for row in rows))

    def test_direct_reco_geo_builders_normalize_and_do_not_route_invalid_phones(self):
        records = [
            {"primary_key": "KEEP", "email": None, "phone": "+639171234567", "product": "CARDS", "level": 1, "AREA": "Mindanao"},
            {"primary_key": "EMAIL", "email": "one@example.com", "phone": "123", "product": "CARDS", "level": 1, "AREA": "Mindanao"},
            {"primary_key": "SKIP", "email": None, "phone": "123", "product": "CARDS", "level": 1, "AREA": "Mindanao"},
        ]
        for builder in (build_negative_exports, build_geo_exports):
            bundle = builder(create_default_csv_template(), records, "", 1, ".csv")
            phone = next(file for file in bundle.files if file.channel == "viber")
            self.assertEqual([(row["primary_key"], row["phone"]) for row in csv_rows(phone)], [("KEEP", "9171234567")])
        self.assertEqual(records[0]["phone"], "+639171234567")

    def test_b2_b4_campaigns_and_sms_viber_activity_columns_and_remarks_use_ten_digits(self):
        for bucket in (2, 4):
            with self.subTest(bucket=bucket):
                database = selection(workbook_data(
                    ["Account Key", "Contract Number", "Bucket Code", "Product Category Code", "Email (01)", "Mobile 1"],
                    [["00001234567890123456", "00009876543210987654", bucket, "MC", "one@example.com", "+639171234567"]],
                ), "SBC.xlsx")
                accounts = selection(workbook_data(["Account No.", "Financing/Card No."],
                                                   [["00001234567890123456", "00009876543210987654"]]), "Accounts.xlsx")
                result = build_b2_exports(database, accounts) if bucket == 2 else build_b4_exports(database, accounts)
                self.assertEqual(csv_rows(result.bundle.files[0])[0]["phone"], "9171234567")
                file = result.bundle.files[0]
                activity_source = SourceSelection(file.filename, file.data, inspect_contact_source(file.data, ".csv")[0])
                activity = B2ActivitySource(activity_source)
                builder = build_b2_autostat_exports if bucket == 2 else build_b4_autostat_exports
                logs = builder([activity], timestamp=None)
                for log in logs.files[1:]:
                    row = csv_rows(log)[0]
                    self.assertEqual(row["Mobile Number"], "9171234567")
                    self.assertTrue(row["REMARKS"].startswith("9171234567 |"))
                    book = load_workbook(BytesIO(build_autostat_channel_excel(log)))
                    try:
                        rows = list(book.active.values)
                        self.assertEqual(rows[1][rows[0].index("Mobile Number")], "9171234567")
                    finally:
                        book.close()


if __name__ == "__main__":
    unittest.main()
