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

    def test_cached_autostat_refreshes_to_cardnumber_mapping(self):
        self.run_isolated("""
            import autostat, autostat_page
            del autostat.PREFERRED_ACCOUNT_HEADER
            autostat.AUTOSTAT_VERSION = 1
            autostat.ACCOUNT_ALIASES = ('primary_key', 'Account No.')
            from app_startup import refresh_application_modules
            refreshed = refresh_application_modules()
            assert 'autostat' in refreshed
            assert 'autostat_page' in refreshed
            data = b'primary_key,cardnumber,email\\nOTHER,000123,first@example.com\\n'
            sheet = autostat.inspect_contact_source(data, '.csv')[0]
            contacts = autostat.extract_contacts(autostat.SourceSelection('contacts.csv', data, sheet))
            assert contacts.emails == [('000123', 'first@example.com')]
            assert autostat_page.AUTOSTAT_VERSION == autostat.AUTOSTAT_VERSION
            assert refresh_application_modules() == set()
        """)

    def test_cached_autostat_refreshes_for_bel_split_and_filenames(self):
        self.run_isolated("""
            from datetime import datetime
            import autostat, autostat_page, export_config
            del export_config.AUTOSTAT_BEL_FILENAMES
            del autostat.build_bel_exports
            from app_startup import refresh_application_modules
            refreshed = refresh_application_modules()
            assert {'export_config', 'autostat', 'autostat_page'} <= refreshed
            data = b'primary_key,email\\n@BEL-1,first@example.com\\n'
            sheet = autostat.inspect_contact_source(data, '.csv')[0]
            bundle = autostat.build_bel_exports(
                [autostat.SourceSelection('contacts.csv', data, sheet)], timestamp=datetime(2026, 10, 28),
            )
            assert [file.filename for file in bundle.files] == ['NEGATIVE B4 - BEL.csv', 'NEGATIVE B4 - CARDS.csv']
            assert autostat_page.build_autostat_exports is autostat.build_autostat_exports
            assert refresh_application_modules() == set()
        """)

    def test_cached_autostat_refreshes_for_account_deduplication(self):
        self.run_isolated("""
            from datetime import datetime
            import autostat, autostat_page
            del autostat.ACCOUNT_DEDUP_VERSION
            del autostat.unique_account_contacts
            from app_startup import refresh_application_modules
            refreshed = refresh_application_modules()
            assert {'autostat', 'autostat_page'} <= refreshed
            data = b'cardnumber,email,phone\\n0001,first@example.com,09171234567\\n0001,next@example.com,09181234567\\n'
            sheet = autostat.inspect_contact_source(data, '.csv')[0]
            bundle = autostat.build_autostat_exports(
                [autostat.SourceSelection('contacts.csv', data, sheet)], timestamp=datetime(2026, 10, 28),
            )
            assert [file.count for file in bundle.files] == [1, 1]
            assert autostat_page.AUTOSTAT_VERSION == autostat.AUTOSTAT_VERSION
            assert refresh_application_modules() == set()
        """)

    def test_cached_autostat_refreshes_to_two_compiled_channel_files(self):
        self.run_isolated("""
            from datetime import datetime
            import autostat, autostat_page
            del autostat.COMPILED_OUTPUT_VERSION
            autostat.AUTOSTAT_VERSION = 5
            autostat_page.build_autostat_exports = lambda *args, **kwargs: None
            from app_startup import refresh_application_modules
            refreshed = refresh_application_modules()
            assert {'autostat', 'autostat_page'} <= refreshed
            contents = [
                b'cardnumber,email,phone\\n0001,first@example.com,\\n',
                b'cardnumber,email,phone\\n0001,later@example.com,09171234567\\n0002,second@example.com,09181234567\\n',
            ]
            sources = [autostat.SourceSelection(
                f'source-{index}.csv', data, autostat.inspect_contact_source(data, '.csv')[0],
            ) for index, data in enumerate(contents)]
            bundle = autostat_page.build_autostat_exports(sources, timestamp=datetime(2026, 10, 28))
            assert [file.filename for file in bundle.files] == ['AUTOSTAT_Email.csv', 'AUTOSTAT_Viber.csv']
            assert [file.count for file in bundle.files] == [2, 2]
            assert len(bundle.summary) == 2
            assert autostat.AUTOSTAT_VERSION == 8
            assert autostat_page.AUTOSTAT_PAGE_VERSION == 4
            assert refresh_application_modules() == set()
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

    def test_cached_modules_refresh_for_geo_doclass_and_export_preview(self):
        self.run_isolated("""
            import csv
            from io import StringIO
            import automation, exports, folder_output
            del automation.geo_doclass
            del automation.DERIVED_TEMPLATE_HEADERS
            import app
            assert app.MAPPING_VERSION == 6
            assert 'doclass' in app.DERIVED_TEMPLATE_HEADERS
            assert exports.geo_doclass is automation.geo_doclass
            assert folder_output.AutomationError is automation.AutomationError
            bundle = exports.build_sheet_exports(
                automation.create_default_csv_template(),
                [{'email': 'first@example.com', 'AREA': 'Mindanao'}],
                '', 1, template_extension='.csv', base_name='GEO', label='GEO',
            )
            rows = list(csv.DictReader(StringIO(bundle.files[0].data.decode('cp1252'))))
            assert rows[0]['doclass'] == 'SBC GEO Mindanao'
            assert bundle.files[0].preview[0]['doclass'] == 'SBC GEO Mindanao'
            from app_startup import refresh_application_modules
            assert refresh_application_modules() == set()
        """)

    def test_cached_exports_refresh_for_grouping_other_source_sheets(self):
        self.run_isolated("""
            import automation, exports, folder_output
            error_class = automation.AutomationError
            del exports.GROUPED_SHEET_EXPORT_VERSION
            import app
            assert app.GROUPED_SHEET_EXPORT_VERSION == 1
            assert exports.AutomationError is error_class
            assert folder_output.AutomationError is error_class
            bundle = app.build_negative_exports(
                automation.create_default_csv_template(),
                [{'email': 'first@example.com', 'product': 'CARDS', 'level': 1,
                  'ChristmasAmnesty': 'QUALIFIED'}],
                '', 1, '.csv', source_sheet='POSITIVE',
            )
            assert all(file.filename.startswith('POSITIVE - ') for file in bundle.files)
            assert bundle.counts['qualified_cards_level_1'] == 1
            from app_startup import refresh_application_modules
            assert refresh_application_modules() == set()
        """)

    def test_cached_modules_refresh_for_partial_source_headers(self):
        self.run_isolated("""
            from io import BytesIO
            from openpyxl import Workbook
            import automation, exports
            del automation.SheetInfo.available_fields
            del automation.inspect_source_sheet
            import app
            assert app.MAPPING_VERSION == 6
            assert exports.AutomationError is automation.AutomationError
            workbook = Workbook()
            workbook.active.append(['primary_key', 'email', 'PHONE'])
            workbook.active.append(['0001', 'first@example.com', '09171234567'])
            output = BytesIO()
            workbook.save(output)
            workbook.close()
            data = output.getvalue()
            info = automation.inspect_workbook(data)[0]
            assert set(info.available_fields) == {'primary_key', 'email', 'phone'}
            mapped = app.read_source(data, info.name, 1, allow_missing=True)
            assert mapped.records[0]['primary_key'] == '0001'
            assert mapped.records[0]['phone'] == '09171234567'
            assert mapped.records[0]['product'] is None
        """)


    def test_cached_modules_refresh_for_b2_matching_and_workbook_dimensions(self):
        self.run_isolated("""
            import automation, autostat, b2_matching, b2_page, export_config
            del automation.WORKBOOK_READER_VERSION
            del export_config.AUTOSTAT_B2_FILENAMES
            del b2_matching.build_b2_exports
            import app
            assert app.MAPPING_VERSION == 6
            assert automation.WORKBOOK_READER_VERSION == 1
            assert b2_page.build_b2_exports is b2_matching.build_b2_exports
            assert export_config.AUTOSTAT_B2_FILENAMES['CARDS'] == 'NEGATIVE B2 - CARDS'
            assert autostat.AutomationError is automation.AutomationError
            from app_startup import refresh_application_modules
            assert refresh_application_modules() == set()
        """)

    def test_cached_b2_page_refreshes_for_blasting_template(self):
        self.run_isolated("""
            import inspect
            import b2_matching, b2_page
            del b2_matching.DATAGRID_ROUTING_VERSION
            del b2_matching.B2_TEMPLATE_VERSION
            b2_page.render_b2_page = lambda: None
            import app
            assert b2_matching.is_datagrid_sheet(['Account No.', 'Financing/Card No.', 'Status Code'])
            assert b2_matching.B2_MATCHING_VERSION == 2
            assert len(b2_page.B2_TEMPLATE_HEADERS) == 16
            assert 'Time' not in b2_page.B2_TEMPLATE_HEADERS
            assert 'timestamp' not in inspect.signature(b2_matching.build_b2_exports).parameters
            from app_startup import refresh_application_modules
            assert refresh_application_modules() == set()
        """)

    def test_cached_autostat_page_refreshes_after_b2_moves_to_its_own_page(self):
        self.run_isolated("""
            import autostat_page
            del autostat_page.AUTOSTAT_PAGE_VERSION
            autostat_page.render_autostat = lambda: None
            from app_startup import refresh_application_modules
            refreshed = refresh_application_modules()
            assert 'autostat_page' in refreshed
            assert autostat_page.AUTOSTAT_PAGE_VERSION == 4
            assert autostat_page.render_autostat.__name__ == 'render_autostat'
            assert refresh_application_modules() == set()
        """)


if __name__ == "__main__":
    unittest.main()
