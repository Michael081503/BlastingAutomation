from __future__ import annotations

import csv
import unittest
from io import BytesIO, StringIO
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
from zipfile import ZipFile

from openpyxl import Workbook
from streamlit.testing.v1 import AppTest

from automation import AutomationError, read_source


ROOT = Path(__file__).resolve().parents[1]


class Upload(BytesIO):
    def __init__(self, data, name="RECO.xlsx"):
        super().__init__(data)
        self.name = name


def source(account="00001234567890123456"):
    workbook = Workbook()
    workbook.active.title = "NEGATIVE"
    workbook.active.append(["primary_key", "email", "phone", "OB", "Product", "Level", "ChristmasAmnesty"])
    for balance in (100, 900):
        workbook.active.append([account, "one@example.com", "09171234567", balance, "CARDS", 1, "QUALIFIED"])
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return Upload(output.getvalue())


class AutomaticPreparationTests(unittest.TestCase):
    def setUp(self):
        directory = TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.uploads = {"source_upload": source()}
        self.addCleanup(patch.stopall)
        patch("agent_cycle.SETTINGS_PATH", Path(directory.name) / "cycle.json").start()
        patch("streamlit.file_uploader", side_effect=lambda label, **kwargs: self.uploads.get(kwargs["key"])).start()

    def test_upload_alone_prepares_one_click_zip_and_reuses_result(self):
        with patch("automation.read_source", wraps=read_source) as reader, patch("streamlit.download_button") as download:
            app = AppTest.from_file(str(ROOT / "app.py")).run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(len(app.error), 0)
            self.assertEqual(app.session_state["result"]["count"], 2)
            self.assertFalse(any(button.label == "Preview and prepare files" for button in app.button))
            for label in ("Optional settings", "Preview and other downloads"):
                self.assertFalse(next(item for item in app.expander if item.label == label).proto.expanded)
            self.assertFalse(any(item.label in ("AUTOSTAT logs (optional)", "AutoStat log settings")
                                 for item in app.expander))
            self.assertTrue({"geo_autostat_email_excel", "geo_autostat_viber_excel"} <=
                            {call.kwargs.get("key") for call in download.call_args_list})
            zip_call = next(call for call in download.call_args_list if call.kwargs.get("key") == "csv_zip")
            self.assertEqual(zip_call.args[0], "Download RECO CSV ZIP (.csv files)")
            with ZipFile(BytesIO(zip_call.args[1]())) as archive:
                self.assertEqual(len(archive.namelist()), 2)
                for name in archive.namelist():
                    rows = list(csv.DictReader(StringIO(archive.read(name).decode("cp1252"))))
                    self.assertEqual([row["OB"] for row in rows], ["900", "100"])
                    self.assertTrue(all(row["primary_key"] == "\t00001234567890123456" for row in rows))
                    self.assertTrue(all(row["doclass"] == "SBC Cards L1" for row in rows))
            app.run(timeout=20)
            self.assertEqual(reader.call_count, 1)

    def test_replacing_upload_and_invalid_template_never_offer_stale_zip(self):
        app = AppTest.from_file(str(ROOT / "app.py")).run(timeout=20)
        old_signature = app.session_state["result"]["signature"]
        self.uploads["source_upload"] = source("00009999999999999999")
        app.run(timeout=20)
        self.assertNotEqual(app.session_state["result"]["signature"], old_signature)
        self.assertEqual(app.session_state["result"]["files"][0].records[0]["primary_key"], "00009999999999999999")
        self.uploads["template_upload"] = Upload(b"not an Excel file", "template.xlsx")
        app.run(timeout=20)
        self.assertEqual(len(app.exception), 0)
        self.assertTrue(app.error)
        self.assertNotIn("result", app.session_state)
        self.assertFalse(any(item.key == "csv_zip" for item in app.get("download_button")))
        del self.uploads["template_upload"]
        app.run(timeout=20)
        self.assertEqual(app.session_state["result"]["count"], 2)
        self.uploads.clear()
        app.run(timeout=20)
        self.assertNotIn("result", app.session_state)
        self.assertFalse(any(item.key == "csv_zip" for item in app.get("download_button")))

    def test_processing_error_is_cached_until_inputs_change(self):
        with patch("automation.read_source", side_effect=AutomationError("Ambiguous account column")) as reader:
            app = AppTest.from_file(str(ROOT / "app.py")).run(timeout=20)
            app.run(timeout=20)
            self.assertEqual(reader.call_count, 1)
            self.assertEqual(app.error[0].value, "Ambiguous account column")
            self.assertFalse(any(item.key == "csv_zip" for item in app.get("download_button")))
        self.uploads["source_upload"] = source("00007777777777777777")
        app.run(timeout=20)
        self.assertEqual(len(app.exception), 0)
        self.assertEqual(len(app.error), 0)
        self.assertEqual(app.session_state["result"]["files"][0].records[0]["primary_key"], "00007777777777777777")

    def test_same_workbook_with_new_filename_rebuilds_all_accounts_outputs(self):
        app = AppTest.from_file(str(ROOT / "app.py")).run(timeout=20)
        app.radio[0].set_value("All accounts in one group").run(timeout=20)
        old_signature = app.session_state["result"]["signature"]
        self.uploads["source_upload"] = Upload(self.uploads["source_upload"].getvalue(), "RECO - PL.xlsx")
        app.run(timeout=20)
        result = app.session_state["result"]
        self.assertNotEqual(result["signature"], old_signature)
        self.assertTrue(all(file.filename.startswith("RECO - PL_") for file in result["files"]))
        self.assertTrue(all(row["doclass"] == "PL Recovery" for file in result["files"] for row in file.records))


if __name__ == "__main__":
    unittest.main()
