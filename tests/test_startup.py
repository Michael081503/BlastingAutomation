from __future__ import annotations

import subprocess
import sys
import textwrap
import unittest
from pathlib import Path


PROJECT_DIRECTORY = Path(__file__).resolve().parents[1]


class StartupTests(unittest.TestCase):
    def run_isolated(self, script: str) -> None:
        result = subprocess.run(
            [sys.executable, "-c", textwrap.dedent(script)],
            cwd=PROJECT_DIRECTORY, capture_output=True, text=True, timeout=30,
        )
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_normal_reruns_keep_module_classes_and_settings(self):
        self.run_isolated("""
            from app_startup import refresh_application_modules
            assert refresh_application_modules() == set()
            import automation, export_config, exports
            error_class = automation.AutomationError
            file_class = exports.ExportFile
            filenames = export_config.OUTPUT_FILENAMES
            assert refresh_application_modules() == set()
            import app
            assert automation.AutomationError is error_class
            assert app.AutomationError is error_class
            assert exports.ExportFile is file_class
            assert export_config.OUTPUT_FILENAMES is filenames
        """)

    def test_cached_old_configuration_is_refreshed_before_app_import(self):
        self.run_isolated("""
            import automation, export_config, exports, folder_output
            error_class = automation.AutomationError
            for name in ('GEO_AREA_SUFFIXES', 'GEO_OUTPUT_FILENAMES', 'GEO_ZIP_FILENAME'):
                delattr(export_config, name)
            try:
                from export_config import GEO_AREA_SUFFIXES
            except ImportError:
                pass
            else:
                raise AssertionError('Expected the old configuration import to fail')
            import app
            assert app.GEO_AREA_SUFFIXES is export_config.GEO_AREA_SUFFIXES
            assert exports.GEO_OUTPUT_FILENAMES is export_config.GEO_OUTPUT_FILENAMES
            assert app.AutomationError is error_class
            assert folder_output.AutomationError is error_class
            from app_startup import refresh_application_modules
            assert refresh_application_modules() == set()
        """)

    def test_cached_pre_geo_modules_refresh_together_and_export(self):
        self.run_isolated("""
            import csv
            from io import StringIO
            import automation, export_config, exports, folder_output
            for name in ('GEO_AREA_SUFFIXES', 'GEO_OUTPUT_FILENAMES', 'GEO_ZIP_FILENAME'):
                delattr(export_config, name)
            for name in ('MAPPING_VERSION', 'OPTIONAL_FIELD_MAPPING'):
                delattr(automation, name)
            for name in ('GEO_GROUP_LABELS', 'GEO_REVIEW_GROUP_KEYS', 'build_geo_exports'):
                delattr(exports, name)
            import app
            assert exports.AutomationError is automation.AutomationError
            assert folder_output.AutomationError is automation.AutomationError
            assert app.AutomationError is automation.AutomationError
            records = [{
                'email': 'ana@example.com', 'phone': '09171234567',
                'primary_key': '0001', 'AREA': 'Region 3', 'ChristmasAmnesty': 'QUALIFIED',
                'product': 'CARDS', 'level': 'Level 6',
            }]
            bundle = app.build_geo_exports(
                automation.create_default_csv_template(), records, '', 1, '.csv',
            )
            assert len(bundle.files) == 2
            assert all(file.group_key == 'r3_qualified_cards_level_6' for file in bundle.files)
            for file in bundle.files:
                rows = list(csv.DictReader(StringIO(file.data.decode('cp1252'))))
                assert rows[0]['AREA'] == 'Region 3'
                assert rows[0]['primary_key'] == '0001'
        """)


if __name__ == "__main__":
    unittest.main()
