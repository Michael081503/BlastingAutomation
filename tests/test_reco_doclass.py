from __future__ import annotations

import csv
import unittest
from io import BytesIO, StringIO
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from streamlit.testing.v1 import AppTest

from automation import DEFAULT_TEMPLATE_HEADERS, create_default_csv_template
from export_config import GEO_OUTPUT_FILENAMES, GEO_RECO_AGENTS, OUTPUT_FILENAMES
from exports import (
    build_csv_zip, build_excel_zip, build_export_excel, build_geo_exports,
    build_negative_exports, build_sheet_exports, campaign_output_files,
)
from folder_output import save_files_to_folder


ROOT = Path(__file__).resolve().parents[1]
CLASSES = {
    "qualified_cards_level_1": "SBC Cards L1", "blank_cards_level_1": "SBC Cards L1",
    "qualified_cards_level_6": "SBC Cards L6", "blank_cards_level_6": "SBC Cards L6",
    "qualified_bel": "PL Recovery", "blank_bel": "PL Recovery", "blank_other_products": "PL Recovery",
    "qualified_review": "", "blank_review": "",
}


def records():
    return [
        {"primary_key": f"000{index}", "cardnumber": f"00000{index}", "email": f"account{index}@example.com",
         "phone": "09171234567", "OB": index * 100, "product": product, "level": level,
         "ChristmasAmnesty": status, "AREA": "Mindanao", "doclass": "Old class"}
        for index, (product, level, status) in enumerate([
            ("CARDS", 1, "QUALIFIED"), ("CARDS", 6, "QUALIFIED"), ("BEL", 4, "QUALIFIED"),
            ("CARDS", 1, None), ("CARDS", 6, None), ("BEL", 4, None), ("PL", 2, None),
            ("PL", 2, "QUALIFIED"), ("CARDS", 2, None),
        ], 1)
    ]


def csv_rows(data):
    return list(csv.DictReader(StringIO(data.decode("cp1252"))))


def excel_rows(data):
    workbook = load_workbook(BytesIO(data))
    rows = list(workbook.active.values)
    workbook.close()
    return [dict(zip(rows[0], row)) for row in rows[1:]]


class RecoDoclassTests(unittest.TestCase):
    def test_grouped_reco_class_matches_in_previews_csv_excel_zips_and_folder_saves(self):
        source = records()
        unchanged = [dict(record) for record in source]
        headers = ["Extra", "Do Class", *reversed([header for header in DEFAULT_TEMPLATE_HEADERS if header != "doclass"])]
        workbook = Workbook()
        workbook.active.append(headers)
        output = BytesIO()
        workbook.save(output)
        workbook.close()
        for template, extension, sheet_name, class_header in (
            (create_default_csv_template(), ".csv", "", "doclass"),
            (output.getvalue(), ".xlsx", "Sheet", "Do Class"),
        ):
            with self.subTest(template=extension):
                bundle = build_negative_exports(template, source, sheet_name, 1, extension, agents=GEO_RECO_AGENTS)
                self.assertEqual({file.group_key for file in bundle.files}, set(CLASSES))
                with ZipFile(BytesIO(build_csv_zip(bundle.files, protect_identifiers=True))) as csv_zip, \
                        ZipFile(BytesIO(build_excel_zip(bundle.files))) as excel_zip:
                    for file in bundle.files:
                        expected = CLASSES[file.group_key]
                        self.assertEqual(file.doclass, expected)
                        self.assertTrue(all(row["doclass"] == expected for row in file.preview))
                        self.assertTrue(all(row["doclass"] == expected for row in file.records))
                        for data in (file.data, csv_zip.read(file.filename)):
                            self.assertTrue(all(row[class_header] == expected for row in csv_rows(data)))
                        for data in (build_export_excel(file), excel_zip.read(Path(file.filename).with_suffix(".xlsx").name)):
                            self.assertTrue(all((row[class_header] or "") == expected for row in excel_rows(data)))
                for excel in (False, True):
                    with TemporaryDirectory() as directory:
                        files = campaign_output_files(bundle.files, excel=excel, protect_csv_identifiers=True)
                        folder = save_files_to_folder(files, directory, batch_name="RECO")
                        for file, (filename, _) in zip(bundle.files, files):
                            values = excel_rows((folder / filename).read_bytes()) if excel else csv_rows((folder / filename).read_bytes())
                            self.assertTrue(all((row[class_header] or "") == CLASSES[file.group_key] for row in values))
        self.assertEqual(source, unchanged)

    def test_all_accounts_use_whole_filename_labels_in_the_requested_priority(self):
        for filename, expected in (
            ("Campaign - l1", "SBC Cards L1"), ("Campaign_L6", "SBC Cards L6"),
            ("Campaign - pl", "PL Recovery"), ("Campaign - PL - L6", "SBC Cards L6"),
            ("Campaign - PL - L6 - L1", "SBC Cards L1"), ("Sample L10 L60", ""),
        ):
            with self.subTest(filename=filename):
                bundle = build_sheet_exports(create_default_csv_template(), records(), "", 1,
                    base_name=filename, label="RECO", template_extension=".csv", reco=True, agents=GEO_RECO_AGENTS)
                for file in bundle.files:
                    self.assertTrue(all(row["doclass"] == expected for row in csv_rows(file.data)))
                    self.assertTrue(all(row["doclass"] == expected for row in file.preview))
                    self.assertTrue(all((row["doclass"] or "") == expected for row in excel_rows(build_export_excel(file))))

    def test_renamed_reco_outputs_use_the_filename_instead_of_source_product_or_level(self):
        with patch.dict(OUTPUT_FILENAMES, {"blank_cards_level_1": "Renamed campaign - PL"}):
            bundle = build_negative_exports(create_default_csv_template(), records(), "", 1, ".csv", source_sheet="CONTACTED")
        for file in bundle.files:
            if file.group_key == "blank_cards_level_1":
                self.assertTrue(file.filename.startswith("CONTACTED - Renamed campaign - PL"))
                self.assertEqual(csv_rows(file.data)[0]["doclass"], "PL Recovery")
                self.assertEqual(excel_rows(build_export_excel(file))[0]["doclass"], "PL Recovery")

    def test_missing_doclass_header_is_appended_without_changing_existing_columns(self):
        headers = [header for header in DEFAULT_TEMPLATE_HEADERS if header != "doclass"]
        bundle = build_negative_exports((",".join(headers) + "\n").encode(), records(), "", 1, ".csv")
        for file in bundle.files:
            self.assertEqual(file.headers, [*headers, "doclass"])
            self.assertTrue(all(row["doclass"] == CLASSES[file.group_key] for row in csv_rows(file.data)))

    def test_geo_keeps_its_area_class_even_when_its_filename_contains_l1(self):
        with patch.dict(GEO_OUTPUT_FILENAMES, {"mindanao_qualified_cards": "GEO campaign - L1"}):
            bundle = build_geo_exports(create_default_csv_template(), records(), "", 1, ".csv")
        for file in bundle.files:
            self.assertIsNone(file.doclass)
            self.assertTrue(all(row["doclass"] == "SBC GEO Mindanao" for row in csv_rows(file.data)))
            self.assertTrue(all(row["doclass"] == "SBC GEO Mindanao" for row in excel_rows(build_export_excel(file))))

    def test_app_previews_reco_class_and_invalidates_downloads_after_filename_changes(self):
        workbook = Workbook()
        workbook.active.title = "NEGATIVE"
        headers = ["primary_key", "cardnumber", "email", "phone", "OB", "Product", "Level", "ChristmasAmnesty"]
        workbook.active.append(headers)
        for record in records():
            workbook.active.append([record.get({"Product": "product", "Level": "level"}.get(header, header)) for header in headers])
        output = BytesIO()
        workbook.save(output)
        workbook.close()
        upload = BytesIO(output.getvalue())
        upload.name = "RECO - L1.xlsx"
        with TemporaryDirectory() as directory, patch("agent_cycle.SETTINGS_PATH", Path(directory) / "cycle.json"), \
                patch("streamlit.file_uploader", side_effect=lambda label, **kwargs: upload if kwargs.get("key") == "source_upload" else None):
            app = AppTest.from_file(str(ROOT / "app.py")).run(timeout=20)
            app.run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(len(app.error), 0)
            for file in app.session_state["result"]["files"]:
                self.assertTrue(all(row["doclass"] == CLASSES[file.group_key] for row in csv_rows(file.data)))
            self.assertEqual(app.main.dataframe[-1].value.iloc[0]["doclass"], "SBC Cards L1")
            with patch.dict(OUTPUT_FILENAMES, {"qualified_cards_level_1": "Renamed campaign - PL"}):
                app.run(timeout=20)
                self.assertEqual(app.session_state["result"]["files"][0].doclass, "PL Recovery")
                self.assertEqual(app.main.dataframe[-1].value.iloc[0]["doclass"], "PL Recovery")
            app.radio[0].set_value("All accounts in one group").run(timeout=20)
            app.run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            for file in app.session_state["result"]["files"]:
                self.assertTrue(all(row["doclass"] == "SBC Cards L1" for row in csv_rows(file.data)))
