"""Read a selected source sheet and fill a fresh copy of a spreadsheet template."""

from __future__ import annotations

import csv
import re
import unicodedata
from copy import copy
from dataclasses import dataclass
from datetime import date, datetime
from io import BytesIO, StringIO
from pathlib import Path
from typing import Any
from zipfile import BadZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries
from openpyxl.utils.exceptions import InvalidFileException

from email_validation import validated_email


FIELD_MAPPING = {
    "debtor_email": "email",
    "OLD I.C": "indexoneuid",
    "Account no.": "primary_key",
    "financing/Card no.": "cardnumber",
    "Name": "FullName",
    "endo date": "endodate",
    "level": "level",
    "product": "product",
    "OB (Masterlist)": "OB",
    "phone_1": "phone",
    "Christmas Amnesty": "ChristmasAmnesty",
}
TARGET_HEADERS = tuple(FIELD_MAPPING.values())
IDENTIFIER_FIELDS = frozenset({"indexoneuid", "primary_key", "cardnumber", "phone"})
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


class AutomationError(ValueError):
    """An actionable input problem that can be displayed in the app."""


@dataclass(frozen=True)
class SheetInfo:
    name: str
    header_row: int
    matched_fields: int


@dataclass
class MappedData:
    records: list[dict[str, Any]]
    source_rows: list[int]
    blank_rows_skipped: int = 0
    no_contact_rows_skipped: int = 0
    invalid_email_rows: int = 0


def normalize_header(value: Any) -> str:
    """Match headers despite capitalization, extra whitespace, and punctuation."""
    text = unicodedata.normalize("NFKC", str(value or "")).casefold()
    return "".join(character for character in text if character.isalnum())


def _source_aliases(source: str) -> set[str]:
    labels = {normalize_header(source)}
    if source == "OB (Masterlist)":
        labels.add("ob")
    elif source == "product":
        labels.add("producttype")
    return labels


def _load(data: bytes, *, read_only: bool = True, data_only: bool = False):
    if not data:
        raise AutomationError("The workbook is empty. Upload an .xlsx file.")
    try:
        return load_workbook(BytesIO(data), read_only=read_only, data_only=data_only)
    except (BadZipFile, InvalidFileException, KeyError, OSError, ValueError) as exc:
        raise AutomationError(
            "This file could not be read as an Excel workbook. "
            "Open it in Excel and save an unencrypted .xlsx copy."
        ) from exc


def _check_header_row(row: int) -> None:
    if not isinstance(row, int) or row < 1 or row >= 1_048_576:
        raise AutomationError("The header row must be between 1 and 1,048,575.")


def _header_groups(template: bool) -> list[set[str]]:
    if template:
        return [{normalize_header(header)} for header in TARGET_HEADERS]
    return [_source_aliases(source) for source in FIELD_MAPPING]


def _detect_header(sheet, *, template: bool = False) -> tuple[int, int]:
    groups = _header_groups(template)
    best_row, best_score = 1, 0
    for row_number, row in enumerate(
        sheet.iter_rows(max_row=min(sheet.max_row or 50, 50), values_only=True), 1
    ):
        labels = {normalize_header(value) for value in row if value is not None}
        score = sum(bool(labels & group) for group in groups)
        if score > best_score:
            best_row, best_score = row_number, score
        if score == len(groups):
            break
    return best_row, best_score


def inspect_workbook(data: bytes, *, template: bool = False) -> list[SheetInfo]:
    """Find likely header rows without loading all the worksheet data into memory."""
    workbook = _load(data)
    try:
        return [
            SheetInfo(sheet.title, *_detect_header(sheet, template=template))
            for sheet in workbook.worksheets
        ]
    finally:
        workbook.close()


def default_source_sheet(sheets: list[SheetInfo]) -> str:
    if not sheets:
        raise AutomationError("The source workbook has no worksheets.")
    return next(
        (sheet.name for sheet in sheets if normalize_header(sheet.name) == "negative"),
        sheets[0].name,
    )


def _worksheet(workbook, name: str):
    if name not in workbook.sheetnames:
        raise AutomationError(f'The worksheet "{name}" was not found.')
    return workbook[name]


def _headers(sheet, row: int) -> list[Any]:
    _check_header_row(row)
    return list(next(sheet.iter_rows(min_row=row, max_row=row, values_only=True), ()))


def _column_lookup(headers: list[Any], *, template: bool = False) -> dict[str, int]:
    """Resolve required fields by column position; reject ambiguous headers."""
    normalized = [normalize_header(value) for value in headers]
    positions = {}
    missing = []
    for source, target in FIELD_MAPPING.items():
        aliases = {normalize_header(target)} if template else _source_aliases(source)
        matches = [index for index, value in enumerate(normalized) if value in aliases]
        label = target if template else source
        if not matches:
            missing.append(label)
        elif len(matches) > 1:
            raise AutomationError(
                f'Duplicate or ambiguous header "{label}". '
                "Give that column a unique header before processing."
            )
        else:
            positions[target] = matches[0]
    if missing:
        kind = "Template" if template else "Source"
        raise AutomationError(
            f"{kind} headers are missing: {', '.join(missing)}. "
            "Check the selected sheet and header row."
        )
    return positions


def _is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def has_contact_value(value: Any) -> bool:
    return not _is_blank(value)


def has_contact_details(record: dict[str, Any]) -> bool:
    """Keep accounts with a valid email address or a nonempty phone field."""
    return validated_email(record.get("email")) is not None or has_contact_value(record.get("phone"))


def _identifier(value: Any, number_format: str) -> str | None:
    if _is_blank(value):
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        text = str(int(value)) if float(value).is_integer() else str(value)
        # Excel may store 123 with a display mask of 000000. Retain 000123.
        mask = number_format.split(";")[0].strip()
        if re.fullmatch(r"0+", mask) and float(value).is_integer() and value >= 0:
            text = text.zfill(len(mask))
        return text
    return str(value).strip()


def read_source(data: bytes, sheet_name: str, header_row: int) -> MappedData:
    """Read all mapped values, including hidden/filtered rows, in source order."""
    _check_header_row(header_row)
    formulas = _load(data)
    cached = None
    try:
        source = _worksheet(formulas, sheet_name)
        positions = _column_lookup(_headers(source, header_row))
        cached = _load(data, data_only=True)
        values = _worksheet(cached, sheet_name)
        result = MappedData([], [])
        max_column = max(positions.values()) + 1
        source_iterator = source.iter_rows(min_row=header_row + 1, max_col=max_column)
        value_iterator = values.iter_rows(min_row=header_row + 1, max_col=max_column)
        for row_number, (source_row, value_row) in enumerate(
            zip(source_iterator, value_iterator), header_row + 1
        ):
            record = {}
            invalid_email = False
            for target, column in positions.items():
                original, cell = source_row[column], value_row[column]
                value = cell.value
                if original.data_type == "f" and value is None:
                    raise AutomationError(
                        f'{sheet_name}, row {row_number}, {target}: a formula has no saved value. '
                        "Open the source in Excel, recalculate and save it, then upload it again."
                    )
                if cell.data_type == "e":
                    raise AutomationError(
                        f'{sheet_name}, row {row_number}, {target}: Excel error {value}. '
                        "Fix that source cell before processing."
                    )
                if target in IDENTIFIER_FIELDS:
                    value = _identifier(value, cell.number_format)
                elif target == "email":
                    email = validated_email(value)
                    if not _is_blank(value) and email is None:
                        invalid_email = True
                        result.invalid_email_rows += 1
                    value = email
                elif _is_blank(value):
                    value = None
                if target == "ChristmasAmnesty":
                    value = "QUALIFIED" if normalize_header(value) == "qualified" else None
                record[target] = value
            if not invalid_email and all(_is_blank(value) for value in record.values()):
                result.blank_rows_skipped += 1
                continue
            if not has_contact_details(record):
                result.no_contact_rows_skipped += 1
                continue
            result.records.append(record)
            result.source_rows.append(row_number)
        return result
    finally:
        formulas.close()
        if cached is not None:
            cached.close()


def create_default_template() -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Template"
    sheet.append(TARGET_HEADERS)
    sheet.freeze_panes = "A2"
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="167D78")
        sheet.column_dimensions[cell.column_letter].width = 21
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def _write_excel_value(cell, value: Any, target: str | None = None) -> None:
    cell.value = value
    # Values supplied by the source are data, even when they begin with '='.
    if isinstance(value, str):
        cell.data_type = "s"
    if target in IDENTIFIER_FIELDS:
        cell.number_format = "@"
    elif isinstance(value, (date, datetime)) and cell.number_format == "General":
        cell.number_format = "yyyy-mm-dd"


def fill_excel_template(
    template: bytes,
    records: list[dict[str, Any]],
    sheet_name: str,
    header_row: int,
) -> bytes:
    """Keep template headers, column order, styles, and other sheets in a new file.

    The chosen sheet's data area is replaced. Non-mapped template columns are
    left blank; this prevents old example data from leaking into a fresh output.
    """
    _check_header_row(header_row)
    if header_row + len(records) > 1_048_576:
        raise AutomationError("The output exceeds Excel's worksheet row limit.")
    workbook = _load(template, read_only=False)
    try:
        sheet = _worksheet(workbook, sheet_name)
        headers = _headers(sheet, header_row)
        positions = _column_lookup(headers, template=True)
        if any(area.max_row > header_row for area in sheet.merged_cells.ranges):
            raise AutomationError(
                "The template has merged cells in its data area. "
                "Unmerge them before filling the template."
            )
        first_data_row = header_row + 1
        styles = [copy(sheet.cell(first_data_row, column)._style)
                  for column in range(1, len(headers) + 1)]
        row_height = sheet.row_dimensions[first_data_row].height
        for row in sheet.iter_rows(min_row=first_data_row):
            for cell in row:
                cell.value = None
                cell.hyperlink = None
                cell.comment = None
        for offset, record in enumerate(records):
            row_number = first_data_row + offset
            if row_height is not None:
                sheet.row_dimensions[row_number].height = row_height
            # Imported records should remain visible even if the template was filtered.
            sheet.row_dimensions[row_number].hidden = False
            for column, style in enumerate(styles, 1):
                sheet.cell(row_number, column)._style = copy(style)
            for target, column in positions.items():
                _write_excel_value(sheet.cell(row_number, column + 1), record.get(target), target)
        last_row = max(header_row + len(records), first_data_row)
        if sheet.auto_filter.ref:
            sheet.auto_filter.ref = f"A{header_row}:{get_column_letter(len(headers))}{last_row}"
            sheet.auto_filter.filterColumn = []
            sheet.auto_filter.sortState = None
        for table in sheet.tables.values():
            min_column, min_row, max_column, _ = range_boundaries(table.ref)
            if min_row != header_row or table.totalsRowCount:
                raise AutomationError(
                    "The template contains a table outside its header/data area or a totals row. "
                    "Use a template with a single header row and no totals row."
                )
            table.ref = (
                f"{get_column_letter(min_column)}{min_row}:"
                f"{get_column_letter(max_column)}{last_row}"
            )
            if table.autoFilter is not None:
                table.autoFilter.ref = table.ref
                table.autoFilter.filterColumn = []
                table.autoFilter.sortState = None
            for column in table.tableColumns:
                column.calculatedColumnFormula = None
        output = BytesIO()
        workbook.save(output)
        return output.getvalue()
    finally:
        workbook.close()


def _csv_contents(data: bytes) -> tuple[list[str], str]:
    try:
        text = data.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = data.decode("cp1252")
    try:
        delimiter = csv.Sniffer().sniff(text[:8192], delimiters=",;\t|").delimiter
    except csv.Error:
        delimiter = ","
    try:
        headers = next(csv.reader(StringIO(text), delimiter=delimiter), [])
    except csv.Error as exc:
        raise AutomationError("The CSV template could not be read.") from exc
    _column_lookup(headers, template=True)
    return headers, delimiter


def csv_template_headers(data: bytes) -> list[str]:
    return _csv_contents(data)[0]


def _csv_header(headers: list[Any] | tuple[str, ...]) -> bytes:
    output = StringIO(newline="")
    csv.writer(output).writerow(headers)
    return output.getvalue().encode("utf-8-sig")


def create_default_csv_template() -> bytes:
    return _csv_header(TARGET_HEADERS)


def excel_template_to_csv(template: bytes, sheet_name: str, header_row: int) -> bytes:
    """Use only the selected Excel header row as the CSV column layout."""
    workbook = _load(template)
    try:
        headers = _headers(_worksheet(workbook, sheet_name), header_row)
        _column_lookup(headers, template=True)
        return _csv_header(headers)
    finally:
        workbook.close()


def fill_csv_template(template: bytes, records: list[dict[str, Any]]) -> bytes:
    headers, delimiter = _csv_contents(template)
    positions = _column_lookup(headers, template=True)
    output = StringIO(newline="")
    writer = csv.writer(output, delimiter=delimiter)
    writer.writerow(headers)
    for record in records:
        row = [""] * len(headers)
        for target, position in positions.items():
            value = record.get(target)
            if isinstance(value, (date, datetime)):
                value = value.isoformat(sep=" ") if isinstance(value, datetime) else value.isoformat()
            row[position] = "" if value is None else value
        writer.writerow(row)
    return output.getvalue().encode("utf-8-sig")


def output_filename(source_name: str, sheet_name: str, extension: str = ".csv") -> str:
    safe_sheet = re.sub(r'[^\w .-]+', "_", sheet_name).strip(" .") or "sheet"
    safe_source = re.sub(r'[^\w .-]+', "_", Path(source_name).stem).strip(" .") or "source"
    return f"{safe_source}_{safe_sheet}_filled{extension}"
