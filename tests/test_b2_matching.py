from __future__ import annotations

import csv
import re
import unittest
from datetime import datetime
from io import BytesIO, StringIO
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl import Workbook, load_workbook
from streamlit.testing.v1 import AppTest

from automation import AutomationError, inspect_workbook, read_source
from autostat import EMAIL_HEADERS, SourceSelection, inspect_contact_source
from b2_matching import build_b2_exports, matching_sheet_score
from folder_output import MANILA_TIMEZONE


ROOT = Path(__file__).resolve().parents[1]
NOW = datetime(2026, 10, 28, 15, 4, 5, tzinfo=MANILA_TIMEZONE)
RAW_HEADERS = ['Account Key', 'Contract Number', 'Product Category Code', 'Bucket Code',
               'Phone (01)', 'Phone (02)', 'Email (01)', 'Email (02)']
GRID_HEADERS = ['No', 'Name', 'New I.C.', 'Account No.', 'Email', 'Financing/Card No.']


def workbook_data(headers, rows, *, sheet_name, wrong_dimensions=True):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = sheet_name
    sheet.append(headers)
    for row in rows:
        sheet.append(row)
    if sheet_name == 'SBC':
        for row in sheet.iter_rows(min_row=2):
            if row[0].value == 10:
                row[0].number_format = '00000'
                row[1].number_format = '000000'
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    if not wrong_dimensions:
        return output.getvalue()
    rewritten = BytesIO()
    with ZipFile(BytesIO(output.getvalue())) as source, ZipFile(rewritten, 'w', ZIP_DEFLATED) as target:
        for entry in source.infolist():
            data = source.read(entry.filename)
            if entry.filename.startswith('xl/worksheets/sheet') and entry.filename.endswith('.xml'):
                data = re.sub(rb'(<dimension\b[^>]*\bref=")[^"]+(")', rb'\g<1>A1:C1\2', data, count=1)
            target.writestr(entry, data)
    return rewritten.getvalue()


def fixtures():
    raw = workbook_data(RAW_HEADERS, [
        ['00001', '0000000000000001', 'MC', 2, '09171234567', '0281234567', 'bad..email@example.com', 'first@example.com'],
        ['@BEL-1', '@BEL-1', 'BEL', 2, '0281234567', '+639181234567', 'bel@example.com', None],
        ['00002', '0002', 'MC', 4, '09191234567', None, 'b4@example.com', None],
        ['00003', '0003', 'MC', 2, '09191234567', None, 'database-only@example.com', None],
        ['00004', '0004', 'MC', 2, '09191234567', None, 'four@example.com', None],
        ['00005', '0005', 'MC', 2, '09191234567', None, 'five@example.com', None],
        ['00006', '0006', 'MC', 2, '09191234567', None, 'six@example.com', None],
        ['00006', '0007', 'MC', 2, '09191234567', None, 'seven@example.com', None],
        ['00008', '0008', 'PL', 2, '09191234567', None, 'loan@example.com', None],
        ['00009', '0009', 'MC', 2, '#VALUE!', '09191234567', '=1+1', 'recovered@example.com'],
        [10, 123, 'MC', 2, '09201234567', None, 'masked@example.com', None],
        ['00011', '000011', 'MC', 2, '123', None, 'bad..email@example.com', None],
    ], sheet_name='SBC')
    pairs = [('00001', '0000000000000001'), ('@BEL-1', '@BEL-1'), (' 00001 ', '0000000000000001'),
             ('00099', '00099'), (None, None), ('00002', '0002'), ('00004', '0005'), ('00006', None),
             ('00008', '0008'), ('00009', '0009'), ('00010', '000123'), ('00011', '000011')]
    grid = workbook_data(GRID_HEADERS, [[index, 'Account', None, account, None, contract]
                                       for index, (account, contract) in enumerate(pairs, 1)], sheet_name='Sheet')
    return raw, grid


def selection(name, data, database):
    sheet = max(inspect_contact_source(data, '.xlsx'), key=lambda s: matching_sheet_score(s, database=database))
    return SourceSelection(name, data, sheet)


def csv_rows(file):
    return list(csv.DictReader(StringIO(file.data.decode('cp1252'))))


class B2Tests(unittest.TestCase):
    def test_matching_uses_only_grid_accounts_validates_keys_buckets_and_deduplicates(self):
        raw, grid = fixtures()
        result = build_b2_exports(selection('database.xlsx', raw, True), selection('grid.xlsx', grid, False), timestamp=NOW)
        self.assertEqual(result.stats['Database rows'], 12)
        self.assertEqual(result.stats['DataGrid rows'], 12)
        self.assertEqual(result.stats['Matched rows'], 8)
        self.assertEqual(result.stats['Unique B2 accounts'], 5)
        self.assertEqual(result.stats['Unmatched rows'], 1)
        self.assertEqual(result.stats['Missing account numbers'], 1)
        self.assertEqual(result.stats['Conflicting or ambiguous keys'], 2)
        self.assertEqual(result.stats['Other or missing bucket'], 1)
        self.assertEqual(result.stats['Unsupported product'], 1)
        self.assertEqual(result.stats['Duplicate account rows removed'], 1)
        cards, bel = result.bundle.files
        self.assertEqual([file.filename for file in result.bundle.files], ['NEGATIVE B2 - CARDS.csv', 'NEGATIVE B2 - BEL.csv'])
        self.assertEqual([file.count for file in result.bundle.files], [6, 2])
        self.assertEqual({row['Account No.'] for row in csv_rows(cards)}, {'0000000000000001', '0009', '000123'})
        self.assertEqual({row['Account No.'] for row in csv_rows(bel)}, {'@BEL-1'})
        self.assertEqual(result.bundle.summary[0]['No valid contact'], 1)
        for file in result.bundle.files:
            rows = csv_rows(file)
            self.assertEqual(tuple(rows[0]), EMAIL_HEADERS)
            identities = [(row['Account No.'], row['Status']) for row in rows]
            self.assertEqual(len(identities), len(set(identities)))
            for row in rows:
                self.assertEqual(row['Time'], '2026-10-28 03:04:05 PM')
                self.assertEqual(row['Remark By'], 'MBURGOS')
                self.assertIn(row['EMAIL/ NUMBER'], row['Remark'])
                self.assertTrue('@' in row['EMAIL/ NUMBER'] or re.fullmatch(r'9[0-9]{9}', row['EMAIL/ NUMBER']))
        contacts = {row['EMAIL/ NUMBER'] for row in csv_rows(cards)}
        self.assertIn('recovered@example.com', contacts)
        self.assertIn('9191234567', contacts)
        self.assertNotIn('database-only@example.com', contacts)
        with ZipFile(BytesIO(result.bundle.zip_data)) as archive:
            self.assertEqual(archive.namelist(), [file.filename for file in result.bundle.files])
            for file in result.bundle.files:
                self.assertEqual(archive.read(file.filename), file.data)

    def test_incorrect_excel_dimensions_do_not_hide_headers_or_rows_in_either_reader(self):
        raw, grid = fixtures()
        workbook = load_workbook(BytesIO(grid), read_only=True)
        self.assertEqual((workbook.active.max_row, workbook.active.max_column), (1, 3))
        workbook.close()
        info = inspect_workbook(grid)[0]
        self.assertEqual(len(info.headers), len(GRID_HEADERS))
        self.assertIn('email', info.available_fields)
        self.assertIn('cardnumber', info.available_fields)
        mapped = read_source(grid, info.name, info.header_row, allow_missing=True)
        self.assertEqual(mapped.no_contact_rows_skipped, 12)
        source = selection('database.xlsx', raw, True)
        self.assertEqual(len(source.sheet.headers), len(RAW_HEADERS))
        self.assertTrue(source.sheet.ready)

    def test_missing_database_key_headers_produce_an_actionable_error(self):
        raw = workbook_data(['Account Key', 'Contract Number', 'Email (01)'], [['0001', '001', 'first@example.com']], sheet_name='SBC')
        _, grid = fixtures()
        with self.assertRaisesRegex(AutomationError, 'missing Bucket Code'):
            build_b2_exports(selection('database.xlsx', raw, True), selection('grid.xlsx', grid, False), timestamp=NOW)


class Upload(BytesIO):
    def __init__(self, name, data):
        super().__init__(data)
        self.name = name


class B2AppTests(unittest.TestCase):
    def test_matching_workflow_downloads_two_files_saves_and_invalidates(self):
        raw, grid = fixtures()
        database, accounts = Upload('database.xlsx', raw), Upload('DataGrid.xlsx', grid)
        def uploader(label, **kwargs):
            return database if kwargs.get('key') == 'b2_database_upload' else accounts if kwargs.get('key') == 'b2_grid_upload' else None
        with patch('streamlit.file_uploader', side_effect=uploader), TemporaryDirectory(prefix='b2-matching-test-') as directory:
            app = AppTest.from_file(str(ROOT / 'pages' / 'AUTOSTAT.py')).run(timeout=20)
            app.radio(key='autostat_workflow').set_value('B2: Match DataGrid with SBC database').run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            app.button(key='b2_prepare').click().run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(len(app.error), 0)
            result = app.session_state['b2_result']['matched']
            self.assertEqual(len(result.bundle.files), 2)
            self.assertEqual(len(app.get('download_button')), 3)
            app.text_input(key='b2_output_folder').set_value(directory).run(timeout=20)
            app.button(key='b2_save').click().run(timeout=20)
            self.assertEqual(len(app.error), 0)
            saved = Path(app.session_state['b2_result']['saved_folder'])
            self.assertEqual({path.name for path in saved.iterdir()}, {file.filename for file in result.bundle.files})
            for file in result.bundle.files:
                self.assertEqual((saved / file.filename).read_bytes(), file.data)
            app.text_input(key='b2_remark_by').set_value('OTHER').run(timeout=20)
            self.assertNotIn('b2_result', app.session_state)


if __name__ == '__main__':
    unittest.main()
