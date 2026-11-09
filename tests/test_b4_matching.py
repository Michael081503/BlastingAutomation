from __future__ import annotations

import csv
import unittest
from io import BytesIO, StringIO
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from streamlit.testing.v1 import AppTest

from automation import AutomationError
from autostat import SourceSelection, build_autostat_excel, inspect_contact_source
from b2_autostat import B2ActivitySource, b2_metadata_columns, build_b2_autostat_exports
from b2_matching import B2_TEMPLATE_HEADERS, build_b4_exports
from export_config import COLLECTION_AGENTS


ROOT = Path(__file__).resolve().parents[1]
HEADERS = ["Account Key", "Contract Number", "Bucket Code", "Product Category Code",
           "Mobile 2", "Email (02)", "Phone (01)", "Email (01)", "Name"]
GRID_HEADERS = ["Account No.", "Financing/Card No.", "Old I.C."]


def workbook_data(headers, records):
    workbook = Workbook()
    workbook.active.append(headers)
    for row in records:
        workbook.active.append(row)
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def selection(name, data):
    return SourceSelection(name, data, inspect_contact_source(data, ".xlsx")[0])


def rows(file):
    return list(csv.DictReader(StringIO(file.data.decode("cp1252"))))


def fixtures(email="first@example.com"):
    raw = workbook_data(HEADERS, [
        ["00001", "000000000000000001", 4, "MC", "09991234567", "old@example.com", "+639171234567", email, "First"],
        ["BEL-2", "BEL-002", "B4", "BEL", "09981234567", "old-bel@example.com", "09181234567", "bad..email@example.com", "BEL"],
        ["00003", "000003", 4, "MC", "09971234567", "old3@example.com", "123", "third@example.com", "Third"],
        ["00004", "000004", 4, "MC", "09961234567", "fallback@example.com", "123", "invalid", "No valid 01 contacts"],
        ["00005", "000005", 6, "MC", None, None, "09191234567", "b6@example.com", "Other bucket"],
        ["00006", "000006", 2, "MC", None, None, "09191234567", "b2@example.com", "B2"],
        ["00007", "000007", 4, "MC", None, None, "09191234567", "not-in-list@example.com", "Not listed"],
    ])
    grid = workbook_data(GRID_HEADERS, [
        ["00001", "000000000000000001", "IC1"], ["BEL-2", "BEL-002", "IC2"],
        ["00003", "000003", "IC3"], ["00004", "000004", "IC4"], ["00005", "000005", "IC5"],
        ["00006", "000006", "IC6"], ["00001", "000000000000000001", "IC1"],
        ["00001", "000003", "CONFLICT"], ["MISSING", "MISSING", "MISSING"],
    ])
    return raw, grid


class Upload(BytesIO):
    def __init__(self, name, data):
        super().__init__(data)
        self.name = name


class B4MatchingTests(unittest.TestCase):
    def test_available_agent_subset_is_forwarded_and_rejects_no_agents(self):
        raw, grid = fixtures()
        database, accounts = selection('database.xlsx', raw), selection('grid.xlsx', grid)
        available = COLLECTION_AGENTS[1]
        result = build_b4_exports(database, accounts, agents=(available,))
        for file, summary in zip(result.bundle.files, result.bundle.summary):
            self.assertGreater(file.count, 0)
            self.assertEqual({(row['Monthly'], row['CHCode']) for row in rows(file)}, {available})
            self.assertEqual(summary[f'{available[1]} accounts'], file.count)
            self.assertEqual(summary[f'{COLLECTION_AGENTS[0][1]} accounts'], 0)
        with self.assertRaises(AutomationError):
            build_b4_exports(database, accounts, agents=())

    def test_primary_mobile_and_email_headers_override_legacy_contacts(self):
        for primary_email in ('Email (1)', ' EMAIL (01) '):
            with self.subTest(primary_email=primary_email):
                headers = ['Account Key', 'Contract Number', 'Bucket Code', 'Product Category Code',
                           'Phone (01)', 'Email (02)', ' MOBILE 1 ', primary_email]
                raw = workbook_data(headers, [
                    ['0001', '000001', 4, 'MC', '09171234567', 'old@example.com',
                     '+63 918 123 4567', 'primary@example.com'],
                    ['BEL-2', 'BEL-002', 4, 'BEL', '09191234567', 'other@example.com',
                     '09201234567', 'invalid'],
                    ['0003', '000003', 4, 'MC', '09211234567', 'fallback@example.com',
                     '123', 'invalid'],
                    ['0004', '000004', 4, 'MC', '09221234567', 'fallback@example.com',
                     None, None],
                    ['0005', '000005', 4, 'MC', '09231234567', 'old@example.com',
                     '123', 'email-only@example.com'],
                ])
                grid = workbook_data(GRID_HEADERS, [
                    ['0001', '000001', 'IC1'], ['BEL-2', 'BEL-002', 'IC2'],
                    ['0003', '000003', 'IC3'], ['0004', '000004', 'IC4'], ['0005', '000005', 'IC5'],
                ])
                result = build_b4_exports(selection('NEGATIVE B4 DATABASE.xlsx', raw), selection('list.xlsx', grid))
                output = {row['primary_key']: row for file in result.bundle.files for row in rows(file)}
                self.assertEqual(set(output), {'0001', 'BEL-2', '0005'})
                self.assertEqual((output['0001']['email'], output['0001']['phone']),
                                 ('primary@example.com', '9181234567'))
                self.assertEqual((output['BEL-2']['email'], output['BEL-2']['phone']), ('', '9201234567'))
                self.assertEqual((output['0005']['email'], output['0005']['phone']), ('email-only@example.com', ''))
                self.assertEqual(result.bundle.summary[0]['No valid contact'], 2)

    def test_email_1_precedes_email_01_without_row_fallback(self):
        headers = ['Account Key', 'Contract Number', 'Bucket Code', 'Product Category Code',
                   'Email (01)', 'Email (1)', 'Mobile 1']
        raw = workbook_data(headers, [
            ['0001', '000001', 4, 'MC', 'legacy@example.com', 'primary@example.com', '09171234567'],
            ['0002', '000002', 4, 'MC', 'fallback@example.com', None, '09181234567'],
        ])
        grid = workbook_data(GRID_HEADERS, [['0001', '000001', 'IC1'], ['0002', '000002', 'IC2']])
        result = build_b4_exports(selection('database.xlsx', raw), selection('list.xlsx', grid))
        output = {row['primary_key']: row for file in result.bundle.files for row in rows(file)}
        self.assertEqual(output['0001']['email'], 'primary@example.com')
        self.assertEqual(output['0002']['email'], '')

    def test_duplicate_primary_contact_headers_are_rejected(self):
        for duplicated in ('Email (1)', 'Mobile 1'):
            with self.subTest(duplicated=duplicated):
                headers = ['Account Key', 'Contract Number', 'Bucket Code', 'Product Category Code',
                           'Email (1)', 'Mobile 1', duplicated.lower().replace(' ', '')]
                raw = workbook_data(headers, [
                    ['0001', '000001', 4, 'MC', 'first@example.com', '09171234567', 'other@example.com'],
                ])
                grid = workbook_data(GRID_HEADERS, [['0001', '000001', 'IC1']])
                with self.assertRaisesRegex(AutomationError, 'ambiguous'):
                    build_b4_exports(selection('database.xlsx', raw), selection('list.xlsx', grid))

    def test_exact_01_contacts_b4_bucket_and_template_with_matching_audit(self):
        raw, grid = fixtures()
        result = build_b4_exports(selection("endorsement.xlsx", raw), selection("NEGATIVE B4 DB.xlsx", grid))
        self.assertEqual([file.filename for file in result.bundle.files], ["NEGATIVE B4 - CARDS.csv", "NEGATIVE B4 - BEL.csv"])
        self.assertEqual([file.count for file in result.bundle.files], [2, 1])
        self.assertEqual(result.stats["Unique B4 accounts"], 4)
        self.assertEqual(result.stats["Other or missing bucket"], 2)
        self.assertEqual(result.stats["Duplicate account rows removed"], 1)
        self.assertEqual(result.stats["Conflicting or ambiguous keys"], 1)
        self.assertEqual(result.stats["Unmatched rows"], 1)
        output = {row["primary_key"]: row for file in result.bundle.files for row in rows(file)}
        self.assertEqual((output["00001"]["email"], output["00001"]["phone"]), ("first@example.com", "9171234567"))
        self.assertEqual(output["00001"]["Terms"], "000000000000000001")
        self.assertEqual(output["00001"]["doclass"], "SBC CARDS B4")
        self.assertEqual(output["BEL-2"]["email"], "")
        self.assertEqual(output["BEL-2"]["phone"], "9181234567")
        self.assertEqual(output["BEL-2"]["doclass"], "SBC BEL B4")
        self.assertEqual(output["00003"]["phone"], "")
        self.assertEqual(set(output), {"00001", "00003", "BEL-2"})
        self.assertEqual(result.bundle.summary[0]["No valid contact"], 1)
        for file in result.bundle.files:
            self.assertEqual(file.headers, B2_TEMPLATE_HEADERS)
            self.assertEqual([(row["Monthly"], row["CHCode"]) for row in rows(file)],
                             [("Almera Ranceli", "AMRANCELI"), ("Kyarill Palima", "KPALIMA")][:file.count])
        with ZipFile(BytesIO(result.bundle.zip_data)) as archive:
            self.assertEqual(archive.namelist(), [file.filename for file in result.bundle.files])
            for file in result.bundle.files:
                self.assertEqual(archive.read(file.filename), file.data)

    def test_b4_matched_files_use_primary_key_in_all_autostat_account_columns(self):
        raw, grid = fixtures()
        matched = build_b4_exports(selection("endorsement.xlsx", raw), selection("NEGATIVE B4 DB.xlsx", grid))
        sources = []
        for file in matched.bundle.files:
            sheet = inspect_contact_source(file.data, ".csv")[0]
            self.assertEqual(sheet.headers[sheet.account_column], "Terms")  # Generic detection.
            sources.append(B2ActivitySource(SourceSelection(file.filename, file.data, sheet),
                                           *b2_metadata_columns(sheet.headers)))
        bundle = build_b2_autostat_exports(sources, timestamp=None)
        for file, expected in zip(bundle.files, (
            ["00001", "00003"], ["00001", "BEL-2"], ["00001", "BEL-2"],
        )):
            header = "Account No." if file.filename.endswith("Email.csv") else "Account number"
            self.assertEqual([row[header] for row in rows(file)], expected)
        workbook = load_workbook(BytesIO(build_autostat_excel(bundle)))
        for sheet in workbook:
            column = "B" if sheet.title.endswith("Email") else "A"
            expected = ["00001", "00003"] if sheet.title.endswith("Email") else ["00001", "BEL-2"]
            self.assertEqual([cell.value for cell in sheet[column][1:]], expected)
            self.assertTrue(all(cell.data_type == "s" and cell.number_format == "@" for cell in sheet[column][1:]))
        workbook.close()

    def test_b4_requires_unambiguous_email_01_and_phone_01(self):
        raw, grid = fixtures()
        for headers, message in (([h.replace("Phone (01)", "Phone (02)") for h in HEADERS], "Phone \\(01\\)"),
                                 ([h.replace("Email (01)", "Other email") for h in HEADERS], "Email \\(01\\)")):
            data = workbook_data(headers, [["1", "C1", 4, "MC", "09171234567", "old@example.com", "09181234567", "new@example.com", "Name"]])
            with self.subTest(message=message), self.assertRaisesRegex(AutomationError, message):
                build_b4_exports(selection("endorsement.xlsx", data), selection("grid.xlsx", grid))


class B4PageTests(unittest.TestCase):
    def test_daily_availability_reassigns_outputs_and_removes_unavailable_downloads(self):
        raw, grid = fixtures()
        uploads = {'b4_database_upload': Upload('database.xlsx', raw),
                   'b4_grid_upload': Upload('list.xlsx', grid)}
        with patch('streamlit.file_uploader', side_effect=lambda label, **kw: uploads.get(kw['key'])) as uploader, \
                patch('b2_page.render_agent_availability', return_value=COLLECTION_AGENTS) as availability:
            app = AppTest.from_file(str(ROOT / 'pages' / 'NEGATIVE_B2.py')).run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            original = app.session_state['b4_result']
            availability.return_value = (COLLECTION_AGENTS[0],)
            app.run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            current = app.session_state['b4_result']
            self.assertNotEqual(current['signature'], original['signature'])
            self.assertEqual({row['CHCode'] for file in current['matched'].bundle.files for row in rows(file)},
                             {COLLECTION_AGENTS[0][1]})
            self.assertEqual(len(app.get('download_button')), 3)
            availability.return_value = None
            uploader.reset_mock()
            app.run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertNotIn('b4_result', app.session_state)
            self.assertEqual(len(app.get('download_button')), 0)
            self.assertEqual({call.kwargs['key'] for call in uploader.call_args_list},
                             {'b2_database_upload', 'b2_grid_upload', 'b4_database_upload', 'b4_grid_upload'})

    def test_database_with_mobile_1_and_email_1_prepares_downloads(self):
        raw = workbook_data(
            ['Account Key', 'Contract Number', 'Bucket Code', 'Product Category Code', 'Mobile 1', 'Email (1)'],
            [['0001', '000001', 4, 'MC', '+63 918 123 4567', 'first@example.com'],
             ['BEL-2', 'BEL-002', 4, 'BEL', '09191234567', 'bel@example.com']],
        )
        grid = workbook_data(GRID_HEADERS, [['0001', '000001', 'IC1'], ['BEL-2', 'BEL-002', 'IC2']])
        uploads = {'b4_database_upload': Upload('NEGATIVE B4 DATABASE.xlsx', raw),
                   'b4_grid_upload': Upload('list.xlsx', grid)}
        with patch('streamlit.file_uploader', side_effect=lambda label, **kw: uploads.get(kw['key'])):
            app = AppTest.from_file(str(ROOT / 'pages' / 'NEGATIVE_B2.py')).run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(len(app.error), 0)
            result = app.session_state['b4_result']['matched']
            self.assertEqual([file.count for file in result.bundle.files], [1, 1])
            self.assertEqual(rows(result.bundle.files[0])[0]['phone'], '9181234567')
            self.assertEqual(rows(result.bundle.files[0])[0]['email'], 'first@example.com')
            self.assertEqual(len(app.get('download_button')), 3)

    def test_b2_only_database_explains_zero_rows_and_never_offers_empty_downloads(self):
        _, grid = fixtures()
        wrong_raw = workbook_data(HEADERS, [
            ["B2-ONLY", "B2-CARD", 2, "MC", None, None, "09171234567", "b2@example.com", "B2 Person"],
        ])
        uploads = {"b4_database_upload": Upload("daily database.xlsx", wrong_raw),
                   "b4_grid_upload": Upload("NEGATIVE B4 DB.xlsx", grid)}
        with patch("streamlit.file_uploader", side_effect=lambda label, **kw: uploads.get(kw["key"])):
            app = AppTest.from_file(str(ROOT / "pages" / "NEGATIVE_B2.py")).run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertTrue(any("1 B2 accounts and no B4 accounts" in item.value for item in app.error))
            self.assertEqual(len(app.get("download_button")), 0)
            result = app.session_state["b4_result"]["matched"]
            self.assertEqual(result.stats["Database B2 rows"], 1)
            self.assertEqual(result.stats["Database B4 rows"], 0)
            self.assertEqual([file.count for file in result.bundle.files], [0, 0])
            uploads["b4_database_upload"] = Upload("daily database.xlsx", fixtures()[0])
            app.run(timeout=20)
            self.assertEqual(len(app.error), 0)
            self.assertEqual([file.count for file in app.session_state["b4_result"]["matched"].bundle.files], [2, 1])
            self.assertEqual(len(app.get("download_button")), 3)
            uploads["b4_database_upload"] = Upload("daily database.xlsx", wrong_raw)
            app.run(timeout=20)
            self.assertEqual(len(app.get("download_button")), 0)

    def test_unrelated_b4_database_reports_account_mismatch(self):
        _, grid = fixtures()
        raw = workbook_data(HEADERS, [
            ["OTHER", "OTHER-CARD", 4, "MC", None, None, "09171234567", "other@example.com", "Other Person"],
        ])
        uploads = {"b4_database_upload": Upload("B4.xlsx", raw), "b4_grid_upload": Upload("list.xlsx", grid)}
        with patch("streamlit.file_uploader", side_effect=lambda label, **kw: uploads.get(kw["key"])):
            app = AppTest.from_file(str(ROOT / "pages" / "NEGATIVE_B2.py")).run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertTrue(any("No account numbers from this B4 list match" in item.value for item in app.error))
            self.assertEqual(len(app.get("download_button")), 0)

    def test_uploads_automatically_prepare_once_with_visible_details_and_collapsed_options(self):
        raw, grid = fixtures()
        uploads = {"b4_database_upload": Upload("B4.xlsx", raw), "b4_grid_upload": Upload("list.xlsx", grid)}
        with patch("streamlit.file_uploader", side_effect=lambda label, **kw: uploads.get(kw["key"])), \
                patch("b2_page.build_b4_exports", wraps=build_b4_exports) as prepare:
            app = AppTest.from_file(str(ROOT / "pages" / "NEGATIVE_B2.py")).run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(prepare.call_count, 1)
            self.assertEqual(len(app.get("download_button")), 3)
            self.assertFalse(any(widget.key == "b4_prepare" for widget in app.button))
            self.assertTrue(all(not panel.proto.expanded for panel in app.expander))
            self.assertFalse(any(panel.label == "B4 details" for panel in app.expander))
            self.assertTrue(any(heading.value == "B4 details" for heading in app.subheader))
            self.assertEqual(len(app.metric), 0)
            app.run(timeout=20)
            self.assertEqual(prepare.call_count, 1)

    def test_b4_uploads_prepare_on_negative_b2_page_and_clear_independently(self):
        raw, grid = fixtures()
        uploads = {"b4_database_upload": Upload("endorsement.xlsx", raw), "b4_grid_upload": Upload("NEGATIVE B4 DB.xlsx", grid)}
        with patch("streamlit.file_uploader", side_effect=lambda label, **kw: uploads.get(kw["key"])):
            app = AppTest.from_file(str(ROOT / "pages" / "NEGATIVE_B2.py")).run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(len(app.error), 0)
            self.assertEqual(len(app.error), 0)
            result = app.session_state["b4_result"]
            self.assertEqual([file.count for file in result["matched"].bundle.files], [2, 1])
            uploads["b4_database_upload"] = Upload("endorsement.xlsx", fixtures("updated@example.com")[0])
            app.run(timeout=20)
            self.assertIn("b4_result", app.session_state)
            self.assertEqual(rows(app.session_state["b4_result"]["matched"].bundle.files[0])[0]["email"], "updated@example.com")
            app.button(key="clear_all_files").click().run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertNotIn("b4_result", app.session_state)
            self.assertEqual(app.session_state["b4_source_info"], {})
            app.run(timeout=20)
            self.assertNotIn("b4_result", app.session_state)


if __name__ == "__main__":
    unittest.main()
