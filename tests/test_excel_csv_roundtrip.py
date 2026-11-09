"""Opt-in integration check using an installed, hidden Windows Excel instance.

Set RUN_EXCEL_CSV_TESTS=1 to run 20 open/save/close cycles per CSV.
"""

from __future__ import annotations

import csv
import json
import os
import subprocess
import sys
import unittest
from io import BytesIO, StringIO
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook

from automation import create_default_csv_template
from autostat import SourceSelection, inspect_contact_source
from b2_autostat import B2ActivitySource, build_b2_autostat_exports
from b2_matching import build_b2_exports, build_b4_exports
from exports import build_export_csv, build_geo_exports, build_negative_exports


@unittest.skipUnless(sys.platform == "win32" and os.environ.get("RUN_EXCEL_CSV_TESTS") == "1",
                     "Set RUN_EXCEL_CSV_TESTS=1 on Windows with Excel installed.")
class ExcelCsvRoundtripTests(unittest.TestCase):
    def test_all_campaign_accounts_survive_twenty_excel_save_and_reopen_cycles(self):
        identifiers = [
            ("00001234567890123456", "5156041234567897", "000019"),
            ("123456789012345678901234", "000001234567890123", "123E5"),
            ("0000000000000000", "0000000000000001", "JAN1"),
        ]
        records = [dict(primary_key=key, cardnumber=card, indexoneuid=reference,
                        email=f"person{index}@example.com", phone="09171234567", OB=1000 - index,
                        product="MC", level=1, ChristmasAmnesty="QUALIFIED", AREA="Mindanao")
                   for index, (key, card, reference) in enumerate(identifiers)]
        # indexoneuid is intentionally plain CSV; Excel may infer its cell type.
        fields = ("primary_key", "Terms", "phone")
        with TemporaryDirectory(prefix="campaign-excel-csv-") as directory:
            folder = Path(directory)
            manifest = []
            for workflow, builder in (("GEO", build_geo_exports), ("RECO", build_negative_exports)):
                for encoding in ("cp1252", "utf-8-sig"):
                    bundle = builder(create_default_csv_template(), records, "", 1, ".csv", encoding=encoding)
                    for file in bundle.files:
                        path = folder / f"{workflow}-{encoding}-{file.channel}.csv"
                        path.write_bytes(build_export_csv(file))
                        expected = [["\t" + str(row[target]) if row.get(target) else ""
                                     for target in ("primary_key", "cardnumber", "phone")]
                                    for row in file.records]
                        manifest.append({"path": str(path), "columns": [file.headers.index(field) + 1 for field in fields],
                                         "fields": fields, "expected": expected})
            for bucket, builder in ((2, build_b2_exports), (4, build_b4_exports)):
                sources = []
                for database in (True, False):
                    workbook = Workbook()
                    sheet = workbook.active
                    if database:
                        sheet.append(["Account Key", "Contract Number", "Bucket Code", "Product Category Code",
                                      "Email (01)", "Mobile 1"])
                        for index, (key, card, _) in enumerate(identifiers):
                            sheet.append([key, card, bucket, "BEL" if index == 1 else "MC",
                                          f"person{index}@example.com", "09171234567"])
                    else:
                        sheet.append(["Account No.", "Financing/Card No.", "Old I.C."])
                        for values in identifiers:
                            sheet.append(values)
                    output = BytesIO()
                    workbook.save(output)
                    workbook.close()
                    data = output.getvalue()
                    sources.append(SourceSelection("database.xlsx" if database else "list.xlsx", data,
                                                   inspect_contact_source(data, ".xlsx")[0]))
                for encoding in ("cp1252", "utf-8-sig"):
                    bundle = builder(*sources, encoding=encoding).bundle
                    activity_sources = []
                    for file in bundle.files:
                        path = folder / f"B{bucket}-{encoding}-{file.filename}"
                        protected = build_export_csv(file)
                        path.write_bytes(protected)
                        expected = [["\t" + str(row[field]) if row.get(field) else "" for field in fields]
                                    for row in file.preview]
                        manifest.append({"path": str(path), "columns": [file.headers.index(field) + 1 for field in fields],
                                         "fields": fields, "expected": expected})
                        activity_sources.append(B2ActivitySource(SourceSelection(
                            file.filename, protected, inspect_contact_source(protected, ".csv")[0],
                        )))
                    logs = build_b2_autostat_exports(activity_sources, timestamp=None, bucket=bucket, encoding=encoding)
                    for file in logs.files:
                        path = folder / f"B{bucket}-{encoding}-{file.filename}"
                        path.write_bytes(build_export_csv(file))
                        log_fields = ("Account No.",) if file.filename.endswith("Email.csv") else ("Account number", "Mobile Number")
                        expected = [["\t" + row[field] for field in log_fields] for row in file.preview]
                        manifest.append({"path": str(path), "columns": [file.headers.index(field) + 1 for field in log_fields],
                                         "fields": log_fields, "expected": expected})
            manifest_path = folder / "manifest.json"
            manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
            result = subprocess.run(
                ["powershell.exe", "-NoProfile", "-File", str(Path(__file__).with_name("excel_csv_roundtrip.ps1")),
                 "-ManifestPath", str(manifest_path), "-Cycles", "20"],
                stderr=subprocess.PIPE, text=True, timeout=240,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertEqual(len(manifest), 28)
            for item in manifest:
                data = Path(item["path"]).read_bytes()
                encoding = "utf-8-sig" if data.startswith(b"\xef\xbb\xbf") else "cp1252"
                rows = list(csv.DictReader(StringIO(data.decode(encoding), newline="")))
                self.assertEqual(len(rows), len(item["expected"]))
                for row, expected in zip(rows, item["expected"], strict=True):
                    self.assertEqual([row[field] for field in item["fields"]], expected)
