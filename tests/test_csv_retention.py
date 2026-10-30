from __future__ import annotations

import csv
import unittest
from dataclasses import replace
from datetime import datetime
from io import BytesIO, StringIO
from zipfile import ZipFile

from automation import DEFAULT_TEMPLATE_HEADERS, create_default_csv_template
from autostat import SourceSelection, build_autostat_exports, inspect_contact_source
from exports import build_csv_zip, build_export_csv, build_geo_exports, build_negative_exports


class ProtectedCsvTests(unittest.TestCase):
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
                                for header in ('primary_key', 'cardnumber', 'indexoneuid'):
                                    self.assertEqual(row[header], '\t' + record[header])
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
                    with ZipFile(BytesIO(build_csv_zip(bundle.files))) as archive:
                        for file in bundle.files:
                            self.assertEqual(archive.read(file.filename), file.data)

    def test_empty_channels_keep_only_the_unmodified_csv_headers(self):
        bundle = build_negative_exports(create_default_csv_template(),
            [dict(primary_key='00001', email='one@example.com', product='MC', level=1)], '', 1, '.csv')
        empty = next(file for file in bundle.files if file.count == 0)
        self.assertEqual(build_export_csv(empty, protect_identifiers=True), empty.data)


if __name__ == '__main__':
    unittest.main()
