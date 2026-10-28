from __future__ import annotations

import csv
import unittest
from io import BytesIO, StringIO
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook
from streamlit.testing.v1 import AppTest

from automation import AutomationError, read_source, source_column_positions


ROOT = Path(__file__).resolve().parents[1]


def source_data(headers, rows):
    workbook = Workbook()
    workbook.active.title = 'Accounts'
    workbook.active.append(headers)
    for row in rows:
        workbook.active.append(row)
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


class Upload(BytesIO):
    def __init__(self, data):
        super().__init__(data)
        self.name = 'custom.xlsx'


class SourceMappingTests(unittest.TestCase):
    def test_common_contact_headers_and_explicit_custom_columns_are_read(self):
        for email, phone in [('Email Address', 'Phone No.'), ('Email ID', 'Mobile No.'), ('Debtor Email Address', 'Cellphone Number')]:
            with self.subTest(email=email, phone=phone):
                headers = ['Account Number', email, phone]
                data = source_data(headers, [['00001', 'first@example.com', '09171234567']])
                self.assertEqual(source_column_positions(headers), {'email': 1, 'primary_key': 0, 'phone': 2})
                self.assertEqual(read_source(data, 'Accounts', 1, allow_missing=True).records[0]['email'], 'first@example.com')
        data = source_data(['Ledger Ref', 'Inbox', 'Handset'], [['00001', 'first@example.com', '09171234567']])
        row = read_source(data, 'Accounts', 1, allow_missing=True, source_columns={'primary_key': 0, 'email': 1, 'phone': 2}).records[0]
        self.assertEqual(row['primary_key'], '00001')
        self.assertEqual(row['phone'], '09171234567')

    def test_manual_mapping_resolves_ambiguity_and_rejects_invalid_selections(self):
        data = source_data(['primary_key', 'email', 'Email Address'], [['00001', 'first@example.com', 'chosen@example.com']])
        with self.assertRaisesRegex(AutomationError, 'ambiguous'):
            read_source(data, 'Accounts', 1, allow_missing=True)
        result = read_source(data, 'Accounts', 1, allow_missing=True, source_columns={'email': 2})
        self.assertEqual(result.records[0]['email'], 'chosen@example.com')
        for mapping, error in [({'email': 99}, 'existing source column'), ({'wrong': 1}, 'unknown template field'), ({'email': 0}, 'separate source columns')]:
            with self.subTest(mapping=mapping), self.assertRaisesRegex(AutomationError, error):
                read_source(data, 'Accounts', 1, allow_missing=True, source_columns=mapping)

    def test_unknown_headers_can_be_mapped_in_the_page(self):
        upload = Upload(source_data(['Ledger Ref', 'Inbox', 'Handset', 'Person'], [['00001', 'first@example.com', '09171234567', 'First person']]))
        with patch('streamlit.file_uploader', side_effect=lambda label, **kw: upload if kw.get('key') == 'source_upload' else None):
            app = AppTest.from_file(str(ROOT / 'app.py')).run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertTrue(app.button[0].disabled)
            for target, column in [('primary_key', 0), ('email', 1), ('phone', 2), ('FullName', 3)]:
                next(widget for widget in app.selectbox if widget.label == f'Source column for {target}').select(column).run(timeout=20)
            self.assertFalse(app.button[0].disabled)
            app.button[0].click().run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(len(app.error), 0)
            result = app.session_state['result']
            self.assertEqual([file.count for file in result['files']], [1, 1])
            values = list(csv.DictReader(StringIO(result['files'][0].data.decode('cp1252'))))[0]
            self.assertEqual(values['primary_key'], '00001')
            self.assertEqual(values['FullName'], 'First person')
            next(widget for widget in app.selectbox if widget.label == 'Source column for email').select(None).run(timeout=20)
            self.assertNotIn('result', app.session_state)


if __name__ == '__main__':
    unittest.main()
