from __future__ import annotations

import unittest
from io import BytesIO
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook
from streamlit.testing.v1 import AppTest

from automation import create_default_csv_template
from upload_state import FILE_STATE_KEYS, UPLOAD_KEYS, clear_all_files, upload_widget_key


ROOT = Path(__file__).resolve().parents[1]


class Upload(BytesIO):
    def __init__(self, name, data):
        super().__init__(data)
        self.name = name


def workbook_bytes(headers, row, name="NEGATIVE"):
    workbook = Workbook()
    workbook.active.title = name
    workbook.active.append(headers)
    workbook.active.append(row)
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


class ClearUploadStateTests(unittest.TestCase):
    def test_clears_all_upload_generations_results_caches_and_source_mapping_but_keeps_settings(self):
        settings = {"csv_file_type": "CSV UTF-8 (Comma delimited)", "autostat_remark_by": "USER",
                    "geo_b2_autostat_remark_by": "B2USER", "output_folder": "my-output",
                    "autostat_upload_type": "B2", "unrelated_setting": "keep"}
        file_keys = [*UPLOAD_KEYS, *FILE_STATE_KEYS, *(name + "_clear_2" for name in UPLOAD_KEYS),
                     "source_sheet_digest", "manual_mapping_digest", "source_mapping_digest_email",
                     "autostat_12_file_sheets", "autostat_b2_12_file_sheets", "geo_b2_autostat_12_file_sheets",
                     "b2_sheet_True_digest", "b2_header_False_digest", "geo_b2_autostat_preview_123"]
        state = {**settings, **dict.fromkeys(file_keys, object()), "_upload_generation": 2}
        with patch("upload_state.st.session_state", state):
            clear_all_files()
            self.assertTrue(all(key not in state for key in file_keys))
            self.assertTrue(all(state[key] == value for key, value in settings.items()))
            self.assertEqual(state["_upload_generation"], 3)
            self.assertEqual([upload_widget_key(name) for name in UPLOAD_KEYS],
                             [name + "_clear_3" for name in UPLOAD_KEYS])
            clear_all_files()
            self.assertEqual(state["_upload_generation"], 4)


class ClearUploadPageTests(unittest.TestCase):
    def test_geo_reco_clears_source_template_and_hidden_page_results_without_resurrecting_uploads(self):
        source = Upload("GEO-RECO.xlsx", workbook_bytes(
            ["primary_key", "cardnumber", "email", "phone"], ["00001", "CARD1", "one@example.com", "09171234567"],
        ))
        template = Upload("template.csv", create_default_csv_template())
        uploads = {"source_upload": source, "template_upload": template}
        seen = []

        def uploader(label, **kwargs):
            seen.append(kwargs["key"])
            return uploads.get(kwargs["key"])

        with patch("streamlit.file_uploader", side_effect=uploader):
            app = AppTest.from_file(str(ROOT / "app.py")).run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            next(button for button in app.button if button.label == "Preview and prepare files").click().run(timeout=20)
            self.assertEqual(len(app.error), 0)
            for key in ("result", "geo_autostat_result"):
                self.assertIn(key, app.session_state)
            for key in ("autostat_result", "autostat_b2_result", "b2_result", "geo_b2_autostat_result"):
                app.session_state[key] = {"previous": True}
            app.button(key="clear_all_files").click().run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(len(app.error), 0)
            for key in ("result", "geo_autostat_result", "geo_b2_autostat_result", "autostat_result", "autostat_b2_result", "b2_result"):
                self.assertNotIn(key, app.session_state)
            self.assertTrue({name + "_clear_1" for name in uploads} <= set(seen))
            self.assertEqual(len(app.get("download_button")), 1)  # Built-in blank template only.
            app.run(timeout=20)
            self.assertNotIn("geo_b2_autostat_result", app.session_state)
            uploads["source_upload_clear_1"] = source
            app.run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertIn("geo_autostat_result", app.session_state)
            app.button(key="clear_all_files").click().run(timeout=20)
            self.assertNotIn("geo_autostat_result", app.session_state)
            self.assertIn("source_upload_clear_2", seen)

    def test_autostat_clear_removes_every_file_in_standard_and_b2_modes(self):
        upload = Upload("contacts.csv", b"primary_key,email,phone\n00001,one@example.com,09171234567\n")
        for mode, prefix in (("Standard", "autostat_"), ("B2", "autostat_b2_")):
            with self.subTest(mode=mode), patch("streamlit.file_uploader", side_effect=lambda label, **kw: [upload, upload] if kw["key"] == prefix + "uploads" else []):
                app = AppTest.from_file(str(ROOT / "pages" / "AUTOSTAT.py")).run(timeout=20)
                if mode == "B2":
                    app.radio(key="autostat_upload_type").set_value(mode).run(timeout=20)
                app.text_input(key=prefix + "remark_by").set_value("CUSTOM").run(timeout=20)
                app.button(key=prefix + "prepare").click().run(timeout=20)
                self.assertIn(prefix + "result", app.session_state)
                app.button(key="clear_all_files").click().run(timeout=20)
                self.assertEqual(len(app.exception), 0)
                self.assertNotIn(prefix + "result", app.session_state)
                self.assertEqual(len(app.get("download_button")), 0)
                self.assertEqual(app.text_input(key=prefix + "remark_by").value, "CUSTOM")
                self.assertEqual(app.radio(key="autostat_upload_type").value, mode)
                app.run(timeout=20)
                self.assertNotIn(prefix + "result", app.session_state)

    def test_negative_b2_clears_database_account_list_activity_uploads_and_all_results(self):
        database = Upload("SBC.xlsx", workbook_bytes(
            ["Account Key", "Contract Number", "Bucket Code", "Product Category Code", "EMAIL (01)", "Mobile 1"],
            ["00001", "CARD1", 2, "MC", "one@example.com", "09171234567"], "SBC",
        ))
        grid = Upload("DataGrid.xlsx", workbook_bytes(
            ["Account No.", "Financing/Card No.", "Name"], ["00001", "CARD1", "Person"], "DataGrid",
        ))
        activity = Upload("NEGATIVE B2.csv", b"primary_key,email,phone\n00002,two@example.com,09181234567\n")
        uploads = {"b2_database_upload": database, "b2_grid_upload": grid, "geo_b2_autostat_uploads": [activity]}
        seen = []

        def uploader(label, **kwargs):
            seen.append(kwargs["key"])
            return uploads.get(kwargs["key"])

        with patch("streamlit.file_uploader", side_effect=uploader):
            app = AppTest.from_file(str(ROOT / "pages" / "NEGATIVE_B2.py")).run(timeout=20)
            self.assertIn("b2_result", app.session_state)
            self.assertIn("geo_b2_autostat_result", app.session_state)
            app.text_input(key="geo_b2_autostat_remark_by").set_value("CUSTOM").run(timeout=20)
            app.button(key="clear_all_files").click().run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertNotIn("b2_result", app.session_state)
            self.assertNotIn("geo_b2_autostat_result", app.session_state)
            self.assertEqual(app.session_state["b2_source_info"], {})
            self.assertTrue({name + "_clear_1" for name in uploads} <= set(seen))
            self.assertEqual(len(app.get("download_button")), 0)
            app.run(timeout=20)
            self.assertNotIn("b2_result", app.session_state)
            uploads["geo_b2_autostat_uploads_clear_1"] = [activity]
            app.run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(app.text_input(key="geo_b2_autostat_remark_by").value, "CUSTOM")
            self.assertIn("geo_b2_autostat_result", app.session_state)
            app.button(key="clear_all_files").click().run(timeout=20)
            self.assertNotIn("geo_b2_autostat_result", app.session_state)
            self.assertIn("geo_b2_autostat_uploads_clear_2", seen)


if __name__ == "__main__":
    unittest.main()
