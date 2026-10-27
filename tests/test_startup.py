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
            for name in ('GEO_OUTPUT_FILENAMES', 'GEO_ZIP_FILENAME'):
                delattr(export_config, name)
            try:
                from export_config import GEO_OUTPUT_FILENAMES
            except ImportError:
                pass
            else:
                raise AssertionError('Expected the old configuration import to fail')
            import app
            assert app.GEO_OUTPUT_FILENAMES is export_config.GEO_OUTPUT_FILENAMES
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
            for name in ('GEO_OUTPUT_FILENAMES', 'GEO_ZIP_FILENAME'):
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
            assert all(file.group_key == 'r3_qualified_cards' for file in bundle.files)
            for file in bundle.files:
                rows = list(csv.DictReader(StringIO(file.data.decode('cp1252'))))
                assert rows[0]['AREA'] == 'Region 3'
                assert rows[0]['primary_key'] == '0001'
        """)

    def test_cached_level_grouping_and_old_bootstrap_refresh_to_product_groups(self):
        self.run_isolated("""
            import app_startup, automation, export_config, exports
            error_class = automation.AutomationError
            export_config.GEO_OUTPUT_FILENAMES = {
                'qualified_cards_level_1': 'Old Level 1',
                'qualified_cards_level_6': 'Old Level 6',
            }
            del exports.GEO_PRODUCT_GROUP_LABELS
            app_startup.refresh_application_modules = lambda: set()
            import app
            assert 'mindanao_qualified_cards' in app.GEO_OUTPUT_FILENAMES
            assert 'r3_qualified_cards' in app.GEO_OUTPUT_FILENAMES
            assert 'qualified_cards_level_1' not in app.GEO_OUTPUT_FILENAMES
            assert app.AutomationError is error_class
            records = [
                {'email': 'ana@example.com', 'phone': '09171234567', 'AREA': 'Mindanao',
                 'ChristmasAmnesty': 'QUALIFIED', 'product': 'CARDS', 'level': level}
                for level in (1, 2, 3, 4, 5, 6)
            ]
            bundle = app.build_geo_exports(automation.create_default_csv_template(), records, '', 1, '.csv')
            assert len(bundle.files) == 2
            assert all(file.group_key == 'mindanao_qualified_cards' and file.count == 6 for file in bundle.files)
        """)

    def test_cached_modules_refresh_for_filename_config_and_email_counts(self):
        self.run_isolated("""
            import automation, exports, folder_output
            del automation.MappedData.valid_email_rows
            del automation.MappedData.missing_email_rows
            del exports.GEO_FILENAME_CONFIG_VERSION
            import app
            assert hasattr(automation.MappedData, 'valid_email_rows')
            assert hasattr(automation.MappedData, 'missing_email_rows')
            assert app.GEO_FILENAME_CONFIG_VERSION == 2
            assert exports.AutomationError is automation.AutomationError
            assert folder_output.AutomationError is automation.AutomationError
        """)


if __name__ == "__main__":
    unittest.main()
