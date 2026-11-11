"""Filter insurance DataGrid statuses before deduplicating financing accounts."""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
from itertools import chain
import re
from typing import Any, Iterable

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell

from automation import (
    AutomationError, DEFAULT_TEMPLATE_HEADERS, IDENTIFIER_FIELDS, TARGET_HEADERS,
    _check_header_row, _headers, _identifier, _is_blank, _load, _worksheet,
    _write_excel_value, create_default_csv_template, fill_csv_template,
    normalize_header, source_column_positions,
)
from exports import ExportFile


SBC_INSURANCE_VERSION = 1
INSURANCE_DOCLASS = "SBC INSURANCE"


def status_key(value: Any) -> str:
    """Ignore case and incidental whitespace while retaining the status wording."""
    return " ".join(str(value or "").split()).upper()


def status_label(value: str) -> str:
    return value or "(Blank status)"


@dataclass
class InsuranceRow:
    number: int
    values: tuple[Any, ...]
    financing: str | None
    status: str
    record: dict[str, Any]


@dataclass
class InsuranceSource:
    sheet_name: str
    headers: list[str]
    rows: list[InsuranceRow]
    identifier_columns: set[int]
    blank_rows: int
    non_account_rows: int

    @property
    def status_counts(self) -> Counter[str]:
        return Counter(row.status for row in self.rows)


@dataclass
class InsuranceResult:
    source: InsuranceSource
    rows: list[InsuranceRow]
    stats: dict[str, int]
    file: ExportFile


def _column(headers: list[str], *aliases: str, required: bool = False) -> int | None:
    labels = [normalize_header(value) for value in headers]
    matches = [index for index, label in enumerate(labels)
               if label in {normalize_header(alias) for alias in aliases}]
    if len(matches) > 1:
        raise AutomationError(f"Ambiguous {aliases[0]} column. Keep one column with this header.")
    if matches:
        return matches[0]
    if required:
        raise AutomationError(f"Missing {aliases[0]}. Choose the insurance DataGrid sheet and header row.")
    return None


def read_insurance_source(data: bytes, sheet_name: str, header_row: int) -> InsuranceSource:
    """Read actual worksheet rows, independently of Excel's declared dimensions."""
    _check_header_row(header_row)
    original = _load(data)
    cached = None
    try:
        sheet = _worksheet(original, sheet_name)
        headers = [str(value or "") for value in _headers(sheet, header_row)]
        status_column = _column(headers, "Status Code", "Status", required=True)
        financing_column = _column(headers, "Financing/Card No.", "Contract Number", "cardnumber", required=True)
        account_column = _column(headers, "Account No.", "Account Number", "primary_key", required=True)
        positions = source_column_positions(headers)
        positions.update({"primary_key": account_column, "cardnumber": financing_column})
        # DataGrid's Amount Referred is the campaign's original outstanding balance.
        amount = _column(headers, "Amount Referred")
        if amount is not None:
            positions["OB"] = amount
        collector = _column(headers, "Collector", "CHCode")
        if collector is not None:
            positions["Terms"] = collector
        identifier_columns = {index for target, index in positions.items() if target in IDENTIFIER_FIELDS}
        identifier_columns.update(index for index, header in enumerate(headers)
                                  if normalize_header(header) in {"newic", "debtorid", "clientreferenceno"})
        cached = _load(data, data_only=True)
        options = {"min_row": header_row + 1, "max_col": len(headers)}
        rows, blank_rows, non_account_rows = [], 0, 0
        for number, (formulas, values) in enumerate(zip(
            sheet.iter_rows(**options), _worksheet(cached, sheet_name).iter_rows(**options),
        ), header_row + 1):
            if all(_is_blank(cell.value) for cell in formulas):
                blank_rows += 1
                continue
            cells = []
            for index, (formula, cell) in enumerate(zip(formulas, values)):
                if formula.data_type == "f" and cell.value is None:
                    raise AutomationError(
                        f"{sheet_name}, row {number}, {headers[index]}: a formula has no saved value. "
                        "Recalculate and save the workbook in Excel."
                    )
                if cell.data_type == "e":
                    raise AutomationError(f"{sheet_name}, row {number}, {headers[index]}: fix Excel error {cell.value}.")
                value = _identifier(cell.value, cell.number_format) if index in identifier_columns else cell.value
                if index in {account_column, financing_column} and value is not None and re.fullmatch(
                    r"[+-]?\d+(?:\.\d+)?[eE][+-]?\d+", str(value),
                ):
                    raise AutomationError(
                        f"{sheet_name}, row {number}: an account is written in scientific notation. "
                        "Restore its original digits as text before processing."
                    )
                cells.append(value)
            if not cells[account_column] and not cells[financing_column]:
                non_account_rows += 1
                continue
            record = dict.fromkeys((*TARGET_HEADERS, "Monthly", "Terms"))
            record.update({target: cells[index] for target, index in positions.items()})
            # These DataGrids do not supply campaign contacts. Do not infer them from call history.
            record.update({"email": None, "phone": None})
            rows.append(InsuranceRow(number, tuple(cells), cells[financing_column],
                                     status_key(cells[status_column]), record))
        return InsuranceSource(sheet_name, headers, rows, identifier_columns, blank_rows, non_account_rows)
    finally:
        original.close()
        if cached is not None:
            cached.close()


def prepare_insurance(source: InsuranceSource, included_statuses: Iterable[str], *, encoding: str = "utf-8-sig") -> InsuranceResult:
    """Filter individual rows first; keep the first eligible row per financing number."""
    included = {status_key(status) for status in included_statuses}
    if not included:
        raise AutomationError("Select at least one Status Code value to include.")
    selected, seen = [], set()
    excluded, missing_financing, duplicates = 0, 0, 0
    for row in source.rows:
        if row.status not in included:
            excluded += 1
            continue
        if not row.financing:
            missing_financing += 1
        # Excel's Remove Duplicates groups blank financing numbers together too,
        # as in the supplied filtered DataGrid. Retain the first eligible one.
        identity = row.financing.casefold() if row.financing else ""
        if identity in seen:
            duplicates += 1
            continue
        seen.add(identity)
        selected.append(row)
    records = [row.record for row in selected]
    template = create_default_csv_template(encoding=encoding)
    file = ExportFile(
        key="sbc_insurance", group_key="all", channel="insurance", label=INSURANCE_DOCLASS,
        filename="SBC INSURANCE.csv",
        data=fill_csv_template(template, records, encoding=encoding, doclass=INSURANCE_DOCLASS),
        count=len(records), missing_emails=len(records), headers=list(DEFAULT_TEMPLATE_HEADERS),
        preview=records[:100], records=records, doclass=INSURANCE_DOCLASS,
    )
    stats = {
        "Source account rows": len(source.rows),
        "Removed by status": excluded,
        "Selected rows with blank Financing/Card No.": missing_financing,
        "Duplicate financing/card accounts removed": duplicates,
        "Unique accounts included": len(selected),
        "Blank rows skipped": source.blank_rows,
        "Totals or rows without account identifiers skipped": source.non_account_rows,
    }
    return InsuranceResult(source, selected, stats, file)


def build_filtered_datagrid(result: InsuranceResult) -> bytes:
    """Keep the selected source rows and headers for reviewing the cleaning step."""
    workbook = Workbook(write_only=True)
    sheet = workbook.create_sheet("SBC INSURANCE")
    sheet.freeze_panes = "A2"
    for values in chain([result.source.headers], (row.values for row in result.rows)):
        cells = []
        for index, value in enumerate(values):
            cell = WriteOnlyCell(sheet)
            _write_excel_value(cell, value, "primary_key" if index in result.source.identifier_columns else None)
            if isinstance(value, datetime):
                cell.number_format = "yyyy-mm-dd hh:mm:ss"
            cells.append(cell)
        sheet.append(cells)
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()
