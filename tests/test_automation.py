from __future__ import annotations

import csv
import unittest
from datetime import datetime
from io import BytesIO, StringIO
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from streamlit.testing.v1 import AppTest

from automation import (
    FIELD_MAPPING,
    TARGET_HEADERS,
    AutomationError,
    create_default_template,
    default_source_sheet,
    fill_csv_template,
    fill_excel_template,
    inspect_workbook,
    read_source,
)
from export_config import CHANNEL_SUFFIXES, OUTPUT_FILENAMES
from exports import (
    build_negative_exports, build_sheet_exports, channel_filename,
    export_filename, separate_channels, separate_negative,
)
from folder_output import save_files_to_folder


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"


def save(workbook: Workbook) -> bytes:
    stream = BytesIO()
    workbook.save(stream)
    workbook.close()
    return stream.getvalue()


def sample_row(**updates):
    row = {
        "debtor_email": "ana@example.test",
        "OLD I.C": "000123",
        "Account no.": "0000000456",
        "financing/Card no.": "000123456789012345",
        "Name": "Ana Cruz",
        "endo date": datetime(2026, 10, 26),
        "level": 2,
        "product": "Card",
        "OB (Masterlist)": 1250.75,
        "phone_1": "09171234567",
        "Christmas Amnesty": None,
    }
    row.update(updates)
    return row


def source_file(rows=None, *, headers=None, title_rows=0, other_sheet=True):
    headers = list(FIELD_MAPPING) if headers is None else headers
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Summary" if other_sheet else "NEGATIVE"
    if other_sheet:
        sheet.append(["Not debtor data"])
        sheet = workbook.create_sheet("NEGATIVE")
    for _ in range(title_rows):
        sheet.append(["Masterlist export"])
    sheet.append(headers)
    for row in [sample_row()] if rows is None else rows:
        sheet.append([row.get(header) for header in headers])
    return save(workbook)


def loaded(data):
    return load_workbook(BytesIO(data))


class MappingTests(unittest.TestCase):
    def test_negative_default_and_detect_header_below_title(self):
        sheets = inspect_workbook(source_file(title_rows=3))
        self.assertEqual(default_source_sheet(sheets), "NEGATIVE")
        self.assertEqual(sheets[1].header_row, 4)
        self.assertEqual(sheets[1].matched_fields, len(FIELD_MAPPING))

    def test_all_mappings_reordered_source_and_template(self):
        source = source_file(headers=list(reversed(FIELD_MAPPING)))
        records = read_source(source, "NEGATIVE", 1).records
        template = Workbook()
        template.active.title = "Destination"
        headers = ["Extra"] + list(reversed(TARGET_HEADERS))
        template.active.append(headers)
        result = loaded(fill_excel_template(save(template), records, "Destination", 1))
        self.addCleanup(result.close)
        sheet = result.active
        self.assertEqual([cell.value for cell in sheet[1]], headers)
        actual = dict(zip(headers, [cell.value for cell in sheet[2]]))
        expected = {target: sample_row()[source] for source, target in FIELD_MAPPING.items()}
        expected["Extra"] = None
        self.assertEqual(actual, expected)
        for column in ("primary_key", "cardnumber", "phone", "indexoneuid"):
            self.assertEqual(sheet.cell(2, headers.index(column) + 1).data_type, "s")
            self.assertEqual(sheet.cell(2, headers.index(column) + 1).number_format, "@")

    def test_header_variations_and_ob_alias(self):
        workbook = loaded(source_file())
        sheet = workbook["NEGATIVE"]
        for cell in sheet[1]:
            cell.value = "  " + str(cell.value).upper().replace(" ", "\n") + "  "
        sheet.cell(1, 9, "OB")
        data = read_source(save(workbook), "NEGATIVE", 1)
        self.assertEqual(data.records[0]["OB"], 1250.75)
        self.assertEqual(data.records[0]["email"], "ana@example.test")

    def test_sample_product_type_header_and_sixteen_column_template(self):
        workbook = loaded(source_file())
        workbook["NEGATIVE"].cell(1, 8, "Product Type")
        source = save(workbook)
        self.assertEqual(inspect_workbook(source)[1].matched_fields, len(FIELD_MAPPING))
        records = read_source(source, "NEGATIVE", 1).records
        template_headers = [
            "email", "primary_key", "indexoneuid", "cardnumber", "FullName", "phone", "OB",
            "Monthly", "Terms", "doclass", "Level", "Product", "EndoDate", "AREA",
            "ChristmasAmnesty", "NewEndo",
        ]
        workbook = Workbook()
        workbook.active.append(template_headers)
        result = loaded(fill_excel_template(save(workbook), records, "Sheet", 1))
        self.addCleanup(result.close)
        self.assertEqual([cell.value for cell in result.active[1]], template_headers)
        output = dict(zip(template_headers, [cell.value for cell in result.active[2]]))
        self.assertEqual(output["Product"], "Card")
        self.assertEqual(output["Level"], 2)
        self.assertEqual(output["EndoDate"], datetime(2026, 10, 26))
        for header in ("Monthly", "Terms", "doclass", "AREA", "ChristmasAmnesty", "NewEndo"):
            self.assertIsNone(output[header])

    def test_numeric_identifiers_keep_zero_display_masks(self):
        workbook = loaded(source_file())
        sheet = workbook["NEGATIVE"]
        sheet.cell(2, 3, 456).number_format = "0000000000"
        sheet.cell(2, 10, 9171234567).number_format = "00000000000"
        row = read_source(save(workbook), "NEGATIVE", 1).records[0]
        self.assertEqual(row["primary_key"], "0000000456")
        self.assertEqual(row["phone"], "09171234567")

    def test_blank_rows_skipped_duplicates_and_missing_email_kept(self):
        source = source_file([sample_row(), {}, sample_row(), sample_row(debtor_email=None)])
        result = read_source(source, "NEGATIVE", 1)
        self.assertEqual(len(result.records), 3)
        self.assertEqual(result.source_rows, [2, 4, 5])
        self.assertEqual(result.blank_rows_skipped, 1)
        self.assertIsNone(result.records[-1]["email"])

    def test_remove_only_accounts_missing_both_email_and_phone(self):
        contacts = [
            (None, None), (None, "  "), ("a@example.test", None),
            (None, "09171234567"), ("a@example.test", "09171234567"),
            ("  ", "\t"), ("", "09171234567"), ("a@example.test", "  "),
        ]
        source = source_file([sample_row(debtor_email=email, phone_1=phone) for email, phone in contacts])
        result = read_source(source, "NEGATIVE", 1)
        self.assertEqual(result.no_contact_rows_skipped, 3)
        self.assertEqual(result.source_rows, [4, 5, 6, 8, 9])
        self.assertEqual(len(result.records), 5)
        self.assertEqual(result.blank_rows_skipped, 0)

    def test_empty_rows_and_contactless_accounts_counted_separately(self):
        result = read_source(source_file([{}, sample_row(debtor_email=None, phone_1=None)]), "NEGATIVE", 1)
        self.assertEqual(result.records, [])
        self.assertEqual(result.blank_rows_skipped, 1)
        self.assertEqual(result.no_contact_rows_skipped, 1)

    def test_hidden_rows_are_included(self):
        workbook = loaded(source_file())
        workbook["NEGATIVE"].row_dimensions[2].hidden = True
        self.assertEqual(len(read_source(save(workbook), "NEGATIVE", 1).records), 1)

    def test_missing_and_duplicate_headers_fail_with_names(self):
        with self.assertRaisesRegex(AutomationError, "phone_1"):
            read_source(source_file(headers=[header for header in FIELD_MAPPING if header != "phone_1"]), "NEGATIVE", 1)
        workbook = loaded(source_file())
        workbook["NEGATIVE"].cell(1, len(FIELD_MAPPING) + 1, "PHONE 1")
        with self.assertRaisesRegex(AutomationError, "ambiguous.*phone_1"):
            read_source(save(workbook), "NEGATIVE", 1)

    def test_christmas_amnesty_qualified_or_blank(self):
        statuses = ["QUALIFIED", " qualified ", None, "", "NOT QUALIFIED"]
        source = source_file([sample_row(**{"Christmas Amnesty": status}) for status in statuses])
        records = read_source(source, "NEGATIVE", 1).records
        self.assertEqual([row["ChristmasAmnesty"] for row in records], ["QUALIFIED", "QUALIFIED", None, None, None])

    def test_missing_amnesty_header_cannot_silently_become_blank(self):
        source = source_file(headers=[header for header in FIELD_MAPPING if header != "Christmas Amnesty"])
        with self.assertRaisesRegex(AutomationError, "Christmas Amnesty"):
            read_source(source, "NEGATIVE", 1)

    def test_missing_formula_cache_and_excel_errors_fail(self):
        for value, message in [("=1+1", "formula has no saved value"), ("#N/A", "Excel error")]:
            with self.subTest(value=value):
                with self.assertRaisesRegex(AutomationError, message):
                    read_source(source_file([sample_row(**{"OB (Masterlist)": value})]), "NEGATIVE", 1)

    def test_selected_sheet_only(self):
        workbook = loaded(source_file())
        sheet = workbook.create_sheet("POSITIVE")
        sheet.append(list(FIELD_MAPPING))
        sheet.append(list(sample_row(Name="Ben Santos").values()))
        source = save(workbook)
        self.assertEqual(read_source(source, "NEGATIVE", 1).records[0]["FullName"], "Ana Cruz")
        self.assertEqual(read_source(source, "POSITIVE", 1).records[0]["FullName"], "Ben Santos")

    def test_bad_workbook_is_actionable(self):
        with self.assertRaisesRegex(AutomationError, "could not be read"):
            inspect_workbook(b"not an Excel workbook")


class TemplateTests(unittest.TestCase):
    def setUp(self):
        self.records = read_source(source_file(), "NEGATIVE", 1).records

    def test_template_styles_other_sheets_and_stale_rows(self):
        workbook = loaded(create_default_template())
        sheet = workbook.active
        sheet.insert_rows(1)
        sheet.cell(1, 1, "Template title")
        sheet.cell(3, 5, "Old debtor").fill = PatternFill("solid", fgColor="FFE699")
        sheet.cell(4, 5, "Stale debtor")
        sheet.cell(4, 1).hyperlink = "https://example.test/old"
        sheet.row_dimensions[3].height = 25
        sheet.row_dimensions[3].hidden = True
        sheet.auto_filter.ref = "A2:J4"
        workbook.create_sheet("Instructions").cell(1, 1, "Keep this")
        result = loaded(fill_excel_template(save(workbook), self.records, "Template", 2))
        self.addCleanup(result.close)
        sheet = result["Template"]
        self.assertEqual(sheet.cell(1, 1).value, "Template title")
        self.assertEqual(sheet.cell(3, 5).value, "Ana Cruz")
        self.assertEqual(sheet.cell(3, 5).fill.fgColor.rgb, "00FFE699")
        self.assertEqual(sheet.row_dimensions[3].height, 25)
        self.assertFalse(sheet.row_dimensions[3].hidden)
        self.assertIsNone(sheet.cell(4, 5).value)
        self.assertIsNone(sheet.cell(4, 1).hyperlink)
        self.assertEqual(sheet.auto_filter.ref, "A2:K3")
        self.assertEqual(result["Instructions"].cell(1, 1).value, "Keep this")

    def test_table_range_grows_to_fit_output(self):
        workbook = loaded(create_default_template())
        sheet = workbook.active
        sheet.append(["old"] * 10)
        table = Table(displayName="Debtors", ref="A1:J2")
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        sheet.add_table(table)
        result = loaded(fill_excel_template(save(workbook), self.records * 3, "Template", 1))
        self.addCleanup(result.close)
        self.assertEqual(result.active.tables["Debtors"].ref, "A1:J4")

    def test_untrusted_string_remains_text_in_excel(self):
        self.records[0]["FullName"] = "=HYPERLINK(\"https://example.test\")"
        result = loaded(fill_excel_template(create_default_template(), self.records, "Template", 1))
        self.addCleanup(result.close)
        self.assertEqual(result.active.cell(2, 5).data_type, "s")
        self.assertEqual(result.active.cell(2, 5).value, self.records[0]["FullName"])

    def test_csv_keeps_order_quoting_unicode_and_leading_zeros(self):
        headers = ["Extra"] + list(reversed(TARGET_HEADERS))
        template = (";".join(headers) + "\r\n").encode("utf-8-sig")
        self.records[0]["FullName"] = 'Peña, Ana "A"'
        output = fill_csv_template(template, self.records)
        rows = list(csv.reader(StringIO(output.decode("utf-8-sig")), delimiter=";"))
        self.assertEqual(rows[0], headers)
        row = dict(zip(headers, rows[1]))
        self.assertEqual(row["primary_key"], "0000000456")
        self.assertEqual(row["FullName"], 'Peña, Ana "A"')
        self.assertEqual(row["Extra"], "")
        self.assertEqual(row["endodate"], "2026-10-26 00:00:00")

    def test_missing_template_headers_and_merged_data_fail(self):
        workbook = Workbook()
        workbook.active.append(["email"])
        with self.assertRaisesRegex(AutomationError, "Template headers are missing"):
            fill_excel_template(save(workbook), self.records, "Sheet", 1)
        workbook = loaded(create_default_template())
        workbook.active.merge_cells("A2:B2")
        with self.assertRaisesRegex(AutomationError, "merged cells"):
            fill_excel_template(save(workbook), self.records, "Template", 1)


class FakeUpload(BytesIO):
    def __init__(self, name, data):
        super().__init__(data)
        self.name = name


class SeparationTests(unittest.TestCase):
    def records(self, cases):
        rows = [sample_row(**{"product": product, "level": level, "Christmas Amnesty": status})
                for status, product, level in cases]
        return read_source(source_file(rows), "NEGATIVE", 1).records

    def test_status_product_and_level_boundaries_account_for_every_row(self):
        records = self.records([
            ("QUALIFIED", "CARDS", "LEVEL 1"),
            (" qualified ", " cards ", "Level 6"),
            ("QUALIFIED", "CARDS", 1),
            ("QUALIFIED", "CARDS", 6.0),
            ("QUALIFIED", "BEL", "Level 3"),
            ("QUALIFIED", "BEL", None),
            (None, "CARDS", "Level 1"),
            ("NOT QUALIFIED", "BEL", "Level 6"),
            ("QUALIFIED", "CARDS", "Level 2"),
            ("QUALIFIED", "SALAD", "Level 1"),
            (None, "CARDS", "Level 6"),
            (None, "SLA", "Level 1"),
            (None, "PL", None),
            (None, "CARDS", "Level 2"),
        ])
        groups = separate_negative(records)
        self.assertEqual({key: len(rows) for key, rows in groups.items()}, {
            "qualified_cards_level_1": 2, "qualified_cards_level_6": 2,
            "qualified_bel": 2, "qualified_review": 2,
            "blank_cards_level_1": 1, "blank_cards_level_6": 1,
            "blank_bel": 1, "blank_other_products": 2, "blank_review": 1,
        })
        self.assertEqual(sum(map(len, groups.values())), len(records))
        self.assertTrue(all(row["ChristmasAmnesty"] is None
                            for key, rows in groups.items() if key.startswith("blank_") for row in rows))
        self.assertTrue(all(row["ChristmasAmnesty"] == "QUALIFIED" for row in groups["qualified_bel"]))

    def test_zip_has_separate_templates_and_configured_filenames(self):
        records = self.records([
            ("QUALIFIED", "CARDS", "Level 1"),
            ("QUALIFIED", "CARDS", "Level 6"),
            ("QUALIFIED", "BEL", "Level 4"),
            (None, "PL", "Level 6"),
        ])
        with patch.dict(OUTPUT_FILENAMES, {"qualified_cards_level_1": "My custom file"}):
            bundle = build_negative_exports(create_default_template(), records, "Template", 1)
        self.assertEqual(len(bundle.files), 8)
        self.assertEqual(bundle.files[0].filename, "My custom file_Email.xlsx")
        self.assertEqual(bundle.files[1].filename, "My custom file_Viber.xlsx")
        with ZipFile(BytesIO(bundle.zip_data)) as archive:
            self.assertEqual(archive.namelist(), [file.filename for file in bundle.files])
            for file in bundle.files:
                workbook = loaded(archive.read(file.filename))
                self.addCleanup(workbook.close)
                headers = [cell.value for cell in workbook.active[1]]
                actual = dict(zip(headers, [cell.value for cell in workbook.active[2]]))
                self.assertEqual(workbook.active.max_row, 2)
                self.assertEqual(actual["ChristmasAmnesty"], None if file.group_key.startswith("blank_") else "QUALIFIED")
                self.assertEqual(actual["primary_key"], "0000000456")
                self.assertIsNone(actual["phone"] if file.channel == "email" else actual["email"])

    def test_empty_groups_omitted_and_missing_emails_kept(self):
        records = self.records([("QUALIFIED", "BEL", "Level 6")])
        records[0]["email"] = None
        bundle = build_negative_exports(create_default_template(), records, "Template", 1)
        self.assertEqual([file.key for file in bundle.files], ["qualified_bel_email", "qualified_bel_viber"])
        self.assertEqual([file.count for file in bundle.files], [0, 1])
        self.assertEqual(bundle.files[1].missing_emails, 1)
        self.assertEqual(bundle.counts["blank_cards_level_1"], 0)

    def test_direct_grouped_exports_exclude_contactless_accounts_from_both_statuses(self):
        records = self.records([
            ("QUALIFIED", "CARDS", "Level 1"),
            ("QUALIFIED", "BEL", "Level 6"),
            (None, "CARDS", "Level 6"),
            (None, "SLA", "Level 6"),
        ])
        records[0].update(email=None, phone="  ")
        records[1]["phone"] = None  # Email alone must suffice.
        records[2]["email"] = None  # Phone alone must suffice.
        records[3].update(email="\t", phone=None)
        bundle = build_negative_exports(create_default_template(), records, "Template", 1)
        self.assertEqual(bundle.no_contact_rows_skipped, 2)
        self.assertEqual(sum(bundle.counts.values()), 2)
        self.assertEqual([file.key for file in bundle.files], [
            "qualified_bel_email", "qualified_bel_viber", "blank_cards_level_6_email", "blank_cards_level_6_viber",
        ])
        self.assertEqual([file.count for file in bundle.files], [1, 0, 0, 1])
        self.assertEqual(bundle.files[3].missing_emails, 1)

    def test_grouped_csv_has_status_and_template_order(self):
        records = self.records([("QUALIFIED", "CARDS", "Level 1"), (None, "BEL", "Level 6")])
        headers = list(reversed(TARGET_HEADERS))
        template = (",".join(headers) + "\n").encode()
        bundle = build_negative_exports(template, records, "", 1, ".csv")
        for file in bundle.files:
            rows = list(csv.reader(StringIO(file.data.decode("utf-8-sig"))))
            self.assertEqual(rows[0], headers)
            self.assertEqual(rows[1][headers.index("ChristmasAmnesty")], "" if file.group_key.startswith("blank_") else "QUALIFIED")
            self.assertEqual(rows[1][headers.index("phone" if file.channel == "email" else "email")], "")
            self.assertTrue(file.filename.endswith(".csv"))

    def test_invalid_or_duplicate_filenames_fail_before_creating_zip(self):
        for name in ("", "../bad", "file/name", "CON", "bad:name"):
            with self.subTest(name=name), self.assertRaises(AutomationError):
                export_filename(name, ".xlsx")
        records = self.records([("QUALIFIED", "BEL", "Level 6")])
        with patch.dict(OUTPUT_FILENAMES, {"blank_other_products": OUTPUT_FILENAMES["qualified_bel"]}):
            with self.assertRaisesRegex(AutomationError, "unique"):
                build_negative_exports(create_default_template(), records, "Template", 1)


class ChannelTests(unittest.TestCase):
    def records(self):
        base = read_source(source_file(), "NEGATIVE", 1).records[0]
        return [
            {**base, "primary_key": "0001", "email": "email-only@example.test", "phone": None},
            {**base, "primary_key": "0002", "email": None, "phone": "09171234567"},
            {**base, "primary_key": "0003", "email": "both@example.test", "phone": "09181234567"},
            {**base, "primary_key": "0004", "email": " \t", "phone": None},
        ]

    def test_email_only_phone_only_both_and_neither_routing_without_source_mutation(self):
        records = self.records()
        original = [dict(record) for record in records]
        channels = separate_channels(records)
        self.assertEqual([row["primary_key"] for row in channels["email"]], ["0001", "0003"])
        self.assertEqual([row["primary_key"] for row in channels["viber"]], ["0002", "0003"])
        self.assertTrue(all(row["phone"] is None for row in channels["email"]))
        self.assertTrue(all(row["email"] is None for row in channels["viber"]))
        self.assertEqual(records, original)

    def test_channel_workbooks_keep_all_sixteen_headers_and_other_values(self):
        headers = [
            "email", "primary_key", "indexoneuid", "cardnumber", "FullName", "phone", "OB",
            "Monthly", "Terms", "doclass", "Level", "Product", "EndoDate", "AREA",
            "ChristmasAmnesty", "NewEndo",
        ]
        template = Workbook()
        template.active.append(headers)
        template.create_sheet("Instructions").cell(1, 1, "Keep these instructions")
        records = self.records()
        bundle = build_sheet_exports(
            save(template), records, "Sheet", 1, base_name="Campaign.xlsx", label="Test sheet",
        )
        self.assertEqual(bundle.counts, {"all": 3})
        self.assertEqual(bundle.no_contact_rows_skipped, 1)
        self.assertEqual([file.filename for file in bundle.files], ["Campaign_Email.xlsx", "Campaign_Viber.xlsx"])
        self.assertEqual([file.count for file in bundle.files], [2, 2])
        expected_records = {record["primary_key"]: record for record in records}
        for file in bundle.files:
            workbook = loaded(file.data)
            self.addCleanup(workbook.close)
            sheet = workbook["Sheet"]
            self.assertEqual([cell.value for cell in sheet[1]], headers)
            for row in sheet.iter_rows(min_row=2, values_only=True):
                values = dict(zip(headers, row))
                expected = expected_records[values["primary_key"]]
                for source_header, target in FIELD_MAPPING.items():
                    actual_header = next(header for header in headers if header.casefold() == target.casefold())
                    value = None if target == ("phone" if file.channel == "email" else "email") else expected[target]
                    self.assertEqual(values[actual_header], value)
                for header in ("Monthly", "Terms", "doclass", "AREA", "NewEndo"):
                    self.assertIsNone(values[header])
            self.assertEqual(workbook["Instructions"].cell(1, 1).value, "Keep these instructions")

    def test_each_nonempty_group_keeps_both_files_even_when_one_channel_has_no_rows(self):
        records = [self.records()[0]]
        bundle = build_sheet_exports(
            create_default_template(), records, "Template", 1, base_name="Email only", label="Test",
        )
        self.assertEqual([file.count for file in bundle.files], [1, 0])
        workbook = loaded(bundle.files[1].data)
        self.addCleanup(workbook.close)
        self.assertEqual(tuple(cell.value for cell in workbook.active[1]), TARGET_HEADERS)
        data = [row for row in workbook.active.iter_rows(min_row=2, values_only=True) if any(value is not None for value in row)]
        self.assertEqual(data, [])

    def test_every_group_including_review_files_gets_both_channels(self):
        cases = [
            ("QUALIFIED", "CARDS", "Level 1"), ("QUALIFIED", "CARDS", "Level 6"),
            ("QUALIFIED", "BEL", "Level 3"), ("QUALIFIED", "SLA", "Level 6"),
            (None, "CARDS", "Level 1"), (None, "CARDS", "Level 6"),
            (None, "BEL", "Level 5"), (None, "PL", "Level 6"), (None, "CARDS", "Level 2"),
        ]
        rows = [sample_row(product=product, level=level, **{"Christmas Amnesty": status})
                for status, product, level in cases]
        records = read_source(source_file(rows), "NEGATIVE", 1).records
        bundle = build_negative_exports(create_default_template(), records, "Template", 1)
        self.assertEqual(len(bundle.files), 18)
        self.assertEqual(sum(bundle.counts.values()), 9)
        self.assertEqual(sum(file.count for file in bundle.files), 18)
        for key in bundle.counts:
            self.assertEqual({file.channel for file in bundle.files if file.group_key == key}, {"email", "viber"})

    def test_channel_csv_routing_and_header_only_file(self):
        headers = list(reversed(TARGET_HEADERS))
        template = (";".join(headers) + "\n").encode()
        bundle = build_sheet_exports(
            template, self.records(), "", 1, base_name="Campaign.csv", label="Test", extension=".csv",
        )
        for file in bundle.files:
            rows = list(csv.reader(StringIO(file.data.decode("utf-8-sig")), delimiter=";"))
            self.assertEqual(rows[0], headers)
            expected_keys = ["0001", "0003"] if file.channel == "email" else ["0002", "0003"]
            self.assertEqual([row[headers.index("primary_key")] for row in rows[1:]], expected_keys)
            self.assertTrue(all(row[headers.index("phone" if file.channel == "email" else "email")] == "" for row in rows[1:]))
        bundle = build_sheet_exports(
            template, [self.records()[0]], "", 1, base_name="Only email", label="Test", extension=".csv",
        )
        self.assertEqual(list(csv.reader(StringIO(bundle.files[1].data.decode("utf-8-sig")), delimiter=";")), [headers])

    def test_channel_suffix_names_can_be_edited_and_must_be_unique(self):
        self.assertEqual(channel_filename("My campaign.xlsx", "email", ".xlsx"), "My campaign_Email.xlsx")
        with patch.dict(CHANNEL_SUFFIXES, {"viber": "_Phone"}):
            self.assertEqual(channel_filename("My campaign", "viber", ".xlsx"), "My campaign_Phone.xlsx")
        with patch.dict(CHANNEL_SUFFIXES, {"viber": "_Email"}):
            with self.assertRaisesRegex(AutomationError, "unique"):
                build_sheet_exports(create_default_template(), self.records(), "Template", 1, base_name="Test", label="Test")


class FolderOutputTests(unittest.TestCase):
    def test_all_files_saved_without_zip_and_previous_exports_kept(self):
        with TemporaryDirectory(prefix="blasting-folder-test-") as directory:
            parent = Path(directory).resolve()
            files = [("first.xlsx", create_default_template()), ("second.csv", b"email\r\n")]
            with patch("folder_output.datetime") as test_clock:
                test_clock.now.return_value = datetime(2026, 10, 26, 21, 0)
                first = save_files_to_folder(files, parent, batch_name="NEGATIVE")
                (first / "second.csv").write_bytes(b"previous export")
                second = save_files_to_folder(files, parent, batch_name="NEGATIVE")
            self.assertNotEqual(first, second)
            self.assertEqual(first.parent, parent)
            self.assertEqual(second.parent, parent)
            self.assertEqual((first / "second.csv").read_bytes(), b"previous export")
            self.assertEqual({file.name for file in second.iterdir()}, {name for name, _ in files})
            for name, data in files:
                self.assertEqual((second / name).read_bytes(), data)
            workbook = loaded((second / "first.xlsx").read_bytes())
            self.addCleanup(workbook.close)
            self.assertEqual(tuple(cell.value for cell in workbook.active[1]), TARGET_HEADERS)

    def test_relative_output_folder_uses_project_directory(self):
        with TemporaryDirectory(prefix="blasting-folder-test-") as directory:
            project = Path(directory).resolve()
            with patch("folder_output.PROJECT_DIRECTORY", project):
                folder = save_files_to_folder(
                    [("data.csv", b"email\r\n")], "output", batch_name="../NEGATIVE",
                )
            self.assertEqual(folder.parent, project / "output")
            self.assertTrue(folder.is_relative_to(project))

    def test_invalid_and_duplicate_names_fail_before_writing(self):
        with TemporaryDirectory(prefix="blasting-folder-test-") as directory:
            parent = Path(directory).resolve() / "not_created"
            invalid_batches = [
                [], [("../outside.xlsx", b"data")], [("bad:name.xlsx", b"data")],
                [("same.xlsx", b"one"), ("SAME.xlsx", b"two")],
            ]
            for files in invalid_batches:
                with self.subTest(files=[name for name, _ in files]):
                    with self.assertRaises(AutomationError):
                        save_files_to_folder(files, parent, batch_name="NEGATIVE")
                    self.assertFalse(parent.exists())

    def test_existing_file_as_destination_is_not_changed(self):
        with TemporaryDirectory(prefix="blasting-folder-test-") as directory:
            destination = Path(directory).resolve() / "already_a_file"
            destination.write_bytes(b"keep this")
            with self.assertRaisesRegex(AutomationError, "Could not save the files"):
                save_files_to_folder([("data.csv", b"email\r\n")], destination, batch_name="NEGATIVE")
            self.assertEqual(destination.read_bytes(), b"keep this")


class AppTests(unittest.TestCase):
    def test_initial_screen(self):
        app = AppTest.from_file(str(APP_PATH)).run(timeout=20)
        self.assertEqual(len(app.exception), 0)
        self.assertIn("Upload a source workbook", app.info[0].value)

    def test_upload_preview_download_and_invalidation(self):
        workbook = loaded(source_file())
        sheet = workbook.create_sheet("NEXT")
        sheet.append(list(FIELD_MAPPING))
        sheet.append(list(sample_row(Name="Next person").values()))
        upload = FakeUpload("source.xlsx", save(workbook))

        def uploader(label, **kwargs):
            return upload if kwargs.get("key") == "source_upload" else None

        with patch("streamlit.file_uploader", side_effect=uploader):
            app = AppTest.from_file(str(APP_PATH)).run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(app.selectbox[0].value, "NEGATIVE")
            app.radio[0].set_value("All accounts in one group").run(timeout=20)
            app.button[0].click().run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(len(app.error), 0)
            self.assertEqual(app.session_state["result"]["count"], 1)
            self.assertEqual(len(app.get("download_button")), 4)
            output = loaded(app.session_state["result"]["files"][0].data)
            self.addCleanup(output.close)
            self.assertEqual(output.active.cell(2, 5).value, "Ana Cruz")
            # A routine rerun must retain the prepared download.
            app.run(timeout=20)
            self.assertEqual(app.session_state["result"]["count"], 1)
            app.selectbox[0].select("NEXT").run(timeout=20)
            self.assertEqual(len(app.success), 0)
            self.assertEqual(len(app.get("download_button")), 1)

    def test_uploaded_csv_template_and_missing_header_error(self):
        upload = FakeUpload("source.xlsx", source_file())
        template = FakeUpload("template.csv", (",".join(TARGET_HEADERS) + "\n").encode())

        def uploader(label, **kwargs):
            return upload if kwargs.get("key") == "source_upload" else template

        with patch("streamlit.file_uploader", side_effect=uploader):
            app = AppTest.from_file(str(APP_PATH)).run(timeout=20)
            app.radio[0].set_value("All accounts in one group").run(timeout=20)
            app.button[0].click().run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual([file.filename for file in app.session_state["result"]["files"]], [
                "source_NEGATIVE_filled_Email.csv", "source_NEGATIVE_filled_Viber.csv",
            ])
            upload = FakeUpload("bad.xlsx", source_file(headers=[header for header in FIELD_MAPPING if header != "phone_1"]))
            app.run(timeout=20)
            app.radio[0].set_value("All accounts in one group").run(timeout=20)
            app.button[0].click().run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertIn("phone_1", app.error[0].value)
            self.assertEqual(len(app.get("download_button")), 1)

    def test_grouped_exports_default_preview_and_mode_change(self):
        rows = [sample_row(**{"product": product, "level": level, "Christmas Amnesty": status})
                for status, product, level in [
                    ("QUALIFIED", "CARDS", "Level 1"),
                    ("QUALIFIED", "CARDS", "Level 6"),
                    ("QUALIFIED", "BEL", "Level 3"),
                    (None, "CARDS", "Level 1"),
                    (None, "CARDS", "Level 6"),
                    (None, "BEL", "Level 5"),
                    (None, "PL", "Level 6"),
                ]]
        rows.extend([
            sample_row(debtor_email=None, phone_1=None, **{"Christmas Amnesty": "QUALIFIED"}),
            sample_row(debtor_email=" ", phone_1="\t"),
        ])
        upload = FakeUpload("source.xlsx", source_file(rows))

        def uploader(label, **kwargs):
            return upload if kwargs.get("key") == "source_upload" else None

        with patch("streamlit.file_uploader", side_effect=uploader):
            app = AppTest.from_file(str(APP_PATH)).run(timeout=20)
            self.assertEqual(app.radio[0].value, "Separate Christmas Amnesty groups")
            self.assertIn("OUTPUT_FILENAMES", app.code[0].value)
            app.button[0].click().run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(len(app.error), 0)
            self.assertEqual(len(app.session_state["result"]["files"]), 14)
            self.assertEqual(app.session_state["result"]["no_contact_rows"], 2)
            self.assertEqual(app.metric[2].value, "2")
            self.assertEqual(len(app.get("download_button")), 16)
            app.selectbox[-1].select(OUTPUT_FILENAMES["blank_other_products"] + "_Email.xlsx").run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(app.session_state["result"]["count"], 7)
            app.radio[0].set_value("All accounts in one group").run(timeout=20)
            self.assertEqual(len(app.success), 0)
            self.assertEqual(len(app.get("download_button")), 1)

    def test_all_accounts_without_contacts_show_exclusions_and_no_download(self):
        upload = FakeUpload("source.xlsx", source_file([sample_row(debtor_email=None, phone_1=None)]))

        def uploader(label, **kwargs):
            return upload if kwargs.get("key") == "source_upload" else None

        with patch("streamlit.file_uploader", side_effect=uploader):
            app = AppTest.from_file(str(APP_PATH)).run(timeout=20)
            for mode in ("Separate Christmas Amnesty groups", "All accounts in one group"):
                app.radio[0].set_value(mode).run(timeout=20)
                app.button[0].click().run(timeout=20)
                self.assertEqual(len(app.exception), 0)
                self.assertEqual(len(app.error), 0)
                self.assertEqual(app.session_state["result"]["count"], 0)
                self.assertEqual(app.session_state["result"]["no_contact_rows"], 1)
                self.assertEqual(len(app.get("download_button")), 1)
                self.assertIn("No accounts remain", app.warning[0].value)

    def test_grouped_files_saved_to_folder_only_when_clicked(self):
        rows = [
            sample_row(product="CARDS", level="Level 1", **{"Christmas Amnesty": "QUALIFIED"}),
            sample_row(product="PL", level="Level 6"),
        ]
        upload = FakeUpload("source.xlsx", source_file(rows))

        def uploader(label, **kwargs):
            return upload if kwargs.get("key") == "source_upload" else None

        with TemporaryDirectory(prefix="blasting-app-folder-test-") as directory:
            parent = Path(directory).resolve()
            with patch("streamlit.file_uploader", side_effect=uploader):
                app = AppTest.from_file(str(APP_PATH)).run(timeout=20)
                app.button[0].click().run(timeout=20)
                app.text_input[0].set_value(str(parent)).run(timeout=20)
                self.assertEqual(list(parent.iterdir()), [])
                app.button(key="save_to_folder").click().run(timeout=20)
                self.assertEqual(len(app.exception), 0)
                self.assertEqual(len(app.error), 0)
                saved = Path(app.session_state["result"]["saved_folder"])
                files = app.session_state["result"]["files"]
                self.assertEqual({path.name for path in saved.iterdir()}, {file.filename for file in files})
                for file in files:
                    self.assertEqual((saved / file.filename).read_bytes(), file.data)
                app.run(timeout=20)
                self.assertEqual(list(parent.iterdir()), [saved])

    def test_single_csv_saved_to_folder(self):
        upload = FakeUpload("source.xlsx", source_file())
        template = FakeUpload("template.csv", (",".join(TARGET_HEADERS) + "\n").encode())

        def uploader(label, **kwargs):
            return upload if kwargs.get("key") == "source_upload" else template

        with TemporaryDirectory(prefix="blasting-app-folder-test-") as directory:
            with patch("streamlit.file_uploader", side_effect=uploader):
                app = AppTest.from_file(str(APP_PATH)).run(timeout=20)
                app.radio[0].set_value("All accounts in one group").run(timeout=20)
                app.button[0].click().run(timeout=20)
                app.text_input[0].set_value(directory).run(timeout=20)
                app.button(key="save_to_folder").click().run(timeout=20)
                self.assertEqual(len(app.exception), 0)
                self.assertEqual(len(app.error), 0)
                result = app.session_state["result"]
                saved = Path(result["saved_folder"])
                self.assertEqual({path.name for path in saved.iterdir()}, {file.filename for file in result["files"]})
                self.assertEqual(len(result["files"]), 2)
                for file in result["files"]:
                    self.assertEqual((saved / file.filename).read_bytes(), file.data)


if __name__ == "__main__":
    unittest.main()
