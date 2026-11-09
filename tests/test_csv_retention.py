from __future__ import annotations

import csv
import unittest
from dataclasses import replace
from datetime import datetime
from io import BytesIO, StringIO
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from streamlit.testing.v1 import AppTest

from automation import DEFAULT_TEMPLATE_HEADERS, create_default_csv_template
from autostat import SourceSelection, build_autostat_excel, build_autostat_exports, inspect_contact_source
from b2_autostat import B2ActivitySource, b2_metadata_columns, build_b2_autostat_exports
from export_config import COLLECTION_AGENTS
from exports import build_csv_zip, build_export_csv, build_geo_exports, build_negative_exports, campaign_output_files


class ProtectedCsvTests(unittest.TestCase):
    def test_b2_b4_downloads_protect_identifiers_and_feed_exact_primary_keys_to_autostat(self):
        identifiers = [('00001234567890123456', '5156041234567897', '000019'),
                       ('123456789012345678901234', '000001234567890123', 'JAN1')]
        uploads = {}
        for bucket in (2, 4):
            for database in (True, False):
                workbook = Workbook()
                sheet = workbook.active
                if database:
                    sheet.append(['Account Key', 'Contract Number', 'Bucket Code', 'Product Category Code',
                                  'Email (01)', 'Mobile 1', 'Name', 'Total Outstanding'])
                    for index, (key, card, _) in enumerate(identifiers):
                        sheet.append([key, card, bucket, ('MC', 'BEL')[index],
                                      f'person{index}@example.com', '09171234567', 'Person, Example', 1234.5])
                else:
                    sheet.append(['Account No.', 'Financing/Card No.', 'Old I.C.'])
                    for values in identifiers:
                        sheet.append(values)
                upload = BytesIO()
                workbook.save(upload)
                workbook.close()
                upload.name = f'B{bucket} {"database" if database else "list"}.xlsx'
                uploads[f'b{bucket}_{"database" if database else "grid"}_upload'] = upload

        def resolve(data):
            return data() if callable(data) else data

        with patch('streamlit.file_uploader', side_effect=lambda label, **kw: uploads.get(kw['key'])), \
                patch('b2_page.render_agent_availability', return_value=COLLECTION_AGENTS), \
                patch('streamlit.download_button') as download:
            app = AppTest.from_file(str(Path(__file__).resolve().parents[1] / 'pages' / 'NEGATIVE_B2.py'))
            for encoding, option in (('cp1252', 'CSV (Comma delimited)'), ('utf-8-sig', 'CSV UTF-8 (Comma delimited)')):
                app.session_state['b2_csv_type'] = option
                download.reset_mock()
                app.run(timeout=20)
                self.assertEqual(len(app.exception), 0)
                self.assertEqual(len(app.error), 0)
                calls = {call.kwargs.get('key'): call for call in download.call_args_list}
                for bucket in (2, 4):
                    with self.subTest(bucket=bucket, encoding=encoding):
                        bundle = app.session_state[f'b{bucket}_result']['matched'].bundle
                        sources = []
                        with ZipFile(BytesIO(resolve(calls[f'b{bucket}_zip'].args[1]))) as archive:
                            self.assertEqual(archive.namelist(), [file.filename for file in bundle.files])
                            for index, file in enumerate(bundle.files):
                                protected = resolve(calls[f'b{bucket}_download_{index}'].args[1])
                                self.assertEqual(archive.read(file.filename), protected)
                                self.assertEqual(protected.startswith(b'\xef\xbb\xbf'), encoding == 'utf-8-sig')
                                row = next(csv.DictReader(StringIO(protected.decode(encoding))))
                                key, card, reference = identifiers[index]
                                for header, expected in (('primary_key', key), ('cardnumber', card), ('indexoneuid', reference)):
                                    self.assertEqual(row[header], expected if header == 'indexoneuid' else '\t' + expected)
                                    self.assertEqual(file.preview[0][header], expected)
                                self.assertEqual(row['phone'], '\t9171234567')
                                self.assertEqual(row['OB'], '1234.5')
                                self.assertEqual(row['FullName'], 'Person, Example')
                                sheet = inspect_contact_source(protected, '.csv')[0]
                                source = SourceSelection(file.filename, protected, sheet)
                                sources.append(B2ActivitySource(source, *b2_metadata_columns(sheet.headers)))
                        logs = build_b2_autostat_exports(sources, timestamp=None, bucket=bucket, encoding=encoding)
                        with ZipFile(BytesIO(build_csv_zip(logs.files))) as archive:
                            for file in logs.files:
                                header = 'Account No.' if file.filename.endswith('Email.csv') else 'Account number'
                                rows = list(csv.DictReader(StringIO(file.data.decode(encoding))))
                                self.assertEqual([row[header] for row in rows], [key for key, _, _ in identifiers])
                                protected = build_export_csv(file)
                                self.assertEqual(archive.read(file.filename), protected)
                                rows = list(csv.DictReader(StringIO(protected.decode(encoding))))
                                self.assertEqual([row[header] for row in rows], ['\t' + key for key, _, _ in identifiers])
                                if 'Mobile Number' in file.headers:
                                    self.assertTrue(all(row['Mobile Number'] == '\t09171234567' for row in rows))
                                self.assertEqual(build_export_csv(replace(file, data=protected)), protected)
                        saved = load_workbook(BytesIO(build_autostat_excel(logs)))
                        for sheet in saved:
                            column = 'B' if sheet.title.endswith('Email') else 'A'
                            self.assertEqual([(cell.value, cell.data_type, cell.number_format) for cell in sheet[column][1:]],
                                             [(key, 's', '@') for key, _, _ in identifiers])
                        saved.close()

    def test_geo_reco_csv_protects_only_identifiers_and_roundtrips_into_autostat(self):
        accounts = [('00001234567890123456', '5156041234567897'),
                    ('123456789012345678901234', '000001234567890123')]
        records = [dict(primary_key=key, cardnumber=card, indexoneuid=f'0000{index}',
                        email=f'person{index}@example.com', phone='09171234567',
                        FullName='Person, "Example"\nSecond line', OB=1234.5, level=1, product='MC',
                        AREA='Mindanao', endodate=datetime(2026, 10, 30), ChristmasAmnesty='QUALIFIED')
                   for index, (key, card) in enumerate(accounts)]
        for builder in (build_geo_exports, build_negative_exports):
            for encoding in ('cp1252', 'utf-8-sig'):
                with self.subTest(builder=builder.__name__, encoding=encoding):
                    headers = list(reversed(DEFAULT_TEMPLATE_HEADERS)) + ['Extra field']
                    template = StringIO(newline='')
                    csv.writer(template).writerow(headers)
                    bundle = builder(template.getvalue().encode(encoding), records, '', 1, '.csv', encoding=encoding)
                    originals = [file.data for file in bundle.files]
                    with ZipFile(BytesIO(build_csv_zip(bundle.files, protect_identifiers=True))) as archive:
                        for file in bundle.files:
                            protected = archive.read(file.filename)
                            self.assertEqual(protected, build_export_csv(file, protect_identifiers=True))
                            self.assertEqual(protected.startswith(b'\xef\xbb\xbf'), encoding == 'utf-8-sig')
                            rows = list(csv.DictReader(StringIO(protected.decode(encoding), newline='')))
                            for row, record in zip(rows, file.records, strict=True):
                                for header in ('primary_key', 'cardnumber'):
                                    self.assertEqual(row[header], '\t' + record[header])
                                self.assertEqual(row['indexoneuid'], record['indexoneuid'])
                                self.assertEqual(row['phone'], '\t09171234567' if file.channel == 'viber' else '')
                                self.assertEqual(row['OB'], '1234.5')
                                self.assertEqual(row['Level'], '1')
                                self.assertEqual(row['FullName'], records[0]['FullName'])
                                self.assertEqual(row['EndoDate'], '2026-10-30 00:00:00')
                                self.assertEqual(row['Extra field'], '')
                            # Reusing an already protected CSV must not stack prefixes.
                            self.assertEqual(build_export_csv(replace(file, data=protected), protect_identifiers=True), protected)
                            source = SourceSelection(file.filename, protected, inspect_contact_source(protected, '.csv')[0])
                            logs = build_autostat_exports([source], timestamp=None)
                            exported = next(log for log in logs.files if log.count)
                            rows = list(csv.DictReader(StringIO(exported.data.decode('cp1252'))))
                            self.assertEqual([row['Account No.'] for row in rows], [card for _, card in accounts])
                            self.assertTrue(all(row['Account No.'] == row['Account No.'].strip() for row in rows))
                    self.assertEqual([file.data for file in bundle.files], originals)
                    with ZipFile(BytesIO(build_csv_zip(bundle.files, protect_identifiers=False))) as archive:
                        for file in bundle.files:
                            self.assertEqual(archive.read(file.filename), file.data)

    def test_all_csv_delivery_helpers_protect_account_numbers_by_default(self):
        records = [dict(primary_key='00001234567890123456', cardnumber='5156041234567897',
                        indexoneuid='000019', email='one@example.com', phone='09171234567',
                        product='MC', level=1, AREA='Mindanao')]
        for builder in (build_geo_exports, build_negative_exports):
            with self.subTest(builder=builder.__name__):
                files = builder(create_default_csv_template(), records, '', 1, '.csv').files
                saved = dict(campaign_output_files(files, excel=False))
                with ZipFile(BytesIO(build_csv_zip(files))) as archive:
                    for file in files:
                        data = build_export_csv(file)
                        self.assertEqual(archive.read(file.filename), data)
                        self.assertEqual(saved[file.filename], data)
                        row = list(csv.DictReader(StringIO(data.decode('cp1252'))))[0]
                        self.assertEqual(row['primary_key'], '\t00001234567890123456')
                        self.assertEqual(row['cardnumber'], '\t5156041234567897')
                        self.assertEqual(row['indexoneuid'], '000019')

    def test_empty_channels_keep_only_the_unmodified_csv_headers(self):
        bundle = build_negative_exports(create_default_csv_template(),
            [dict(primary_key='00001', email='one@example.com', product='MC', level=1)], '', 1, '.csv')
        empty = next(file for file in bundle.files if file.count == 0)
        self.assertEqual(build_export_csv(empty, protect_identifiers=True), empty.data)


if __name__ == '__main__':
    unittest.main()
