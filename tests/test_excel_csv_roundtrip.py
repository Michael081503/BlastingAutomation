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
from io import StringIO
from pathlib import Path
from tempfile import TemporaryDirectory

from automation import create_default_csv_template
from exports import build_export_csv, build_geo_exports, build_negative_exports


@unittest.skipUnless(sys.platform == "win32" and os.environ.get("RUN_EXCEL_CSV_TESTS") == "1",
                     "Set RUN_EXCEL_CSV_TESTS=1 on Windows with Excel installed.")
class ExcelCsvRoundtripTests(unittest.TestCase):
    def test_geo_and_reco_accounts_survive_twenty_excel_save_and_reopen_cycles(self):
        identifiers = [
            ("00001234567890123456", "5156041234567897", "000019"),
            ("123456789012345678901234", "000001234567890123", "123E5"),
            ("0000000000000000", "0000000000000001", "JAN1"),
        ]
        records = [dict(primary_key=key, cardnumber=card, indexoneuid=reference,
                        email=f"person{index}@example.com", phone="09171234567", OB=1000 - index,
                        product="MC", level=1, ChristmasAmnesty="QUALIFIED", AREA="Mindanao")
                   for index, (key, card, reference) in enumerate(identifiers)]
        fields = ("primary_key", "cardnumber", "indexoneuid", "phone")
        with TemporaryDirectory(prefix="geo-reco-excel-csv-") as directory:
            folder = Path(directory)
            manifest = []
            for workflow, builder in (("GEO", build_geo_exports), ("RECO", build_negative_exports)):
                for encoding in ("cp1252", "utf-8-sig"):
                    bundle = builder(create_default_csv_template(), records, "", 1, ".csv", encoding=encoding)
                    for file in bundle.files:
                        path = folder / f"{workflow}-{encoding}-{file.channel}.csv"
                        path.write_bytes(build_export_csv(file))
                        expected = [["\t" + str(row[field]) if row.get(field) else "" for field in fields]
                                    for row in file.records]
                        manifest.append({"path": str(path), "columns": [file.headers.index(field) + 1 for field in fields],
                                         "expected": expected})
            manifest_path = folder / "manifest.json"
            manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
            result = subprocess.run(
                ["powershell.exe", "-NoProfile", "-File", str(Path(__file__).with_name("excel_csv_roundtrip.ps1")),
                 "-ManifestPath", str(manifest_path), "-Cycles", "20"],
                stderr=subprocess.PIPE, text=True, timeout=240,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertEqual(len(manifest), 8)
            for item in manifest:
                data = Path(item["path"]).read_bytes()
                encoding = "utf-8-sig" if data.startswith(b"\xef\xbb\xbf") else "cp1252"
                rows = list(csv.DictReader(StringIO(data.decode(encoding), newline="")))
                self.assertEqual(len(rows), len(identifiers))
                for row, expected in zip(rows, item["expected"], strict=True):
                    self.assertEqual([row[field] for field in fields], expected)
