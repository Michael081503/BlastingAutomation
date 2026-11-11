"""Filter insurance DataGrid statuses before deduplicating financing accounts."""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from io import BytesIO
from itertools import chain
import re
from typing import Any, Iterable
import warnings

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell

from automation import (
    AutomationError, DEFAULT_TEMPLATE_HEADERS, IDENTIFIER_FIELDS, TARGET_HEADERS,
    _check_header_row, _headers, _identifier, _is_blank, _load, _worksheet,
    _write_excel_value, create_default_csv_template, fill_csv_template,
    normalize_header, source_column_positions,
)
from exports import ExportFile
from autostat import _validated_candidates, validated_phone
from email_validation import validated_email


SBC_INSURANCE_VERSION = 2
SBC_INSURANCE_MASTERLIST_VERSION = 1
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
class InsuranceContact:
    emails: list[str] = field(default_factory=list)
    phones: list[str] = field(default_factory=list)
    addresses: list[tuple[str | None, str | None, str | None]] = field(default_factory=list)
    source_rows: int = 0

    @property
    def email(self) -> str | None:
        return self.emails[0] if self.emails else None

    @property
    def phone(self) -> str | None:
        return self.phones[0] if self.phones else None

    @property
    def address(self) -> tuple[str | None, str | None, str | None]:
        return self.addresses[0] if self.addresses else (None, None, None)


@dataclass
class InsuranceReference:
    contact: InsuranceContact = field(default_factory=InsuranceContact)
    financing: set[str] = field(default_factory=set)


@dataclass
class InsuranceMasterlist:
    accounts: dict[str, InsuranceContact]
    stats: dict[str, int]
    references: dict[str, InsuranceReference] = field(default_factory=dict)


@dataclass(frozen=True)
class InsuranceMasterSheet:
    name: str
    header_row: int


@dataclass
class InsuranceResult:
    source: InsuranceSource
    rows: list[InsuranceRow]
    stats: dict[str, int]
    file: ExportFile
    contacts: list[InsuranceContact | None] | None = None


def _column(headers: list[str], *aliases: str, required: bool = False) -> int | None:
    labels = [normalize_header(value) for value in headers]
    matches = [index for index, label in enumerate(labels)
               if label in {normalize_header(alias) for alias in aliases}]
    if len(matches) > 1:
        raise AutomationError(f"Ambiguous {aliases[0]} column. Keep one column with this header.")
    if matches:
        return matches[0]
    if required:
        raise AutomationError(f"Missing {aliases[0]}. Check the selected worksheet and header row.")
    return None


def inspect_insurance_masterlist(data: bytes) -> list[InsuranceMasterSheet]:
    """Find contact sheets without selecting summaries or combining old snapshots."""
    workbook = _load(data)
    try:
        result = []
        for sheet in workbook.worksheets:
            for number, row in enumerate(sheet.iter_rows(max_row=50, values_only=True), 1):
                labels = {normalize_header(value) for value in row}
                if "pnnumber" in labels and labels.intersection({"email", "emailaddress"}):
                    result.append(InsuranceMasterSheet(sheet.title, number))
                    break
        if not result:
            raise AutomationError("The insurance masterlist needs PN_Number and EMAIL headers. Upload the masterfile, not the DataGrid.")
        return result
    finally:
        workbook.close()


def _contact_text(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return None if text.casefold() in {"", "-", "n/a", "#n/a", "none", "null", "nan"} else text


def _master_identifier(original, cached) -> str | None:
    if cached.data_type == "e" or original.data_type == "f" and cached.value is None:
        return None
    value = _identifier(cached.value, cached.number_format)
    if not _contact_text(value) or re.fullmatch(r"[+-]?\d+(?:\.\d+)?[eE][+-]?\d+", value):
        return None
    if isinstance(cached.value, (int, float)) and abs(cached.value) >= 10**15:
        return None
    return value.casefold()


def _merge_contact(destination: InsuranceContact, incoming: InsuranceContact) -> None:
    destination.source_rows += incoming.source_rows
    for target, values in ((destination.emails, incoming.emails), (destination.phones, incoming.phones)):
        for value in values:
            if value.casefold() not in {previous.casefold() for previous in target}:
                target.append(value)
    for address in incoming.addresses:
        if tuple((v or "").casefold() for v in address) not in {
            tuple((v or "").casefold() for v in previous) for previous in destination.addresses
        }:
            destination.addresses.append(address)


def read_insurance_masterlist(data: bytes, sheet_name: str, header_row: int) -> InsuranceMasterlist:
    """Index contacts by exact IS Ref No# and PN_Number from one masterlist sheet."""
    _check_header_row(header_row)
    original = _load(data)
    cached = None
    try:
        sheet = _worksheet(original, sheet_name)
        headers = [str(value or "") for value in _headers(sheet, header_row)]
        pn_column = _column(headers, "PN_Number", required=True)
        reference_columns = [index for index, header in enumerate(headers)
                             if normalize_header(header) in {"isrefno", "isrefnumber"}]
        columns = {
            "email": _column(headers, "EMAIL", "Email Address", required=True),
            "phone": _column(headers, "MOBILE", "Mobile Number", "phone"),
            "address1": _column(headers, "ADDRESS1", "ADDRESS"),
            "address2": _column(headers, "ADDRESS2"),
            "address3": _column(headers, "ADDRESS3"),
        }
        cached = _load(data, data_only=True)
        options = {"min_row": header_row + 1, "max_col": max(pn_column, *reference_columns, *(c for c in columns.values() if c is not None)) + 1}
        accounts: dict[str, InsuranceContact] = {}
        references: dict[str, InsuranceReference] = {}
        stats = Counter({
            "Masterlist account rows": 0,
            "Masterlist rows with unusable PN_Number": 0,
            "Masterlist rows with conflicting reference columns": 0,
            "Masterlist contact cells with Excel errors or uncached formulas": 0,
            "Invalid masterlist email entries": 0,
            "Invalid masterlist mobile entries": 0,
        })
        with warnings.catch_warnings():
            # Unused policy/date columns in the provided file have invalid date
            # formats. Contact errors are counted explicitly below instead.
            warnings.filterwarnings("ignore", message=r"Cell .* is marked as a date but the serial value .* is outside the limits for dates.*", category=UserWarning)
            for formulas, values in zip(sheet.iter_rows(**options), _worksheet(cached, sheet_name).iter_rows(**options)):
                key = _master_identifier(formulas[pn_column], values[pn_column])
                if key is None and not _is_blank(formulas[pn_column].value):
                    stats["Masterlist rows with unusable PN_Number"] += 1
                reference_values = {_master_identifier(formulas[index], values[index]) for index in reference_columns}
                reference_values.discard(None)
                reference_key = next(iter(reference_values)) if len(reference_values) == 1 else None
                if len(reference_values) > 1:
                    stats["Masterlist rows with conflicting reference columns"] += 1
                if key is None and reference_key is None:
                    continue
                contact = InsuranceContact(source_rows=1)
                stats["Masterlist account rows"] += 1
                cells = {}
                for target, index in columns.items():
                    if index is None:
                        cells[target] = None
                    elif values[index].data_type == "e" or formulas[index].data_type == "f" and values[index].value is None:
                        stats["Masterlist contact cells with Excel errors or uncached formulas"] += 1
                        cells[target] = None
                    else:
                        cells[target] = values[index].value
                for target, validator, destination, invalid in (
                    ("email", validated_email, contact.emails, "Invalid masterlist email entries"),
                    ("phone", validated_phone, contact.phones, "Invalid masterlist mobile entries"),
                ):
                    if _contact_text(cells[target]) is None:
                        continue
                    for _, value in _validated_candidates(cells[target], validator):
                        if value is None:
                            stats[invalid] += 1
                        elif value.casefold() not in {previous.casefold() for previous in destination}:
                            destination.append(value)
                address = tuple(_contact_text(cells[f"address{number}"]) for number in range(1, 4))
                if any(address) and tuple((v or "").casefold() for v in address) not in {
                    tuple((v or "").casefold() for v in previous) for previous in contact.addresses
                }:
                    contact.addresses.append(address)
                if key is not None:
                    _merge_contact(accounts.setdefault(key, InsuranceContact()), contact)
                if reference_key is not None:
                    reference = references.setdefault(reference_key, InsuranceReference())
                    reference.financing.add(key or "")
                    _merge_contact(reference.contact, contact)
        stats["Distinct masterlist PN_Number accounts"] = len(accounts)
        stats["Masterlist PN_Number accounts with repeated rows"] = sum(c.source_rows > 1 for c in accounts.values())
        stats["Distinct masterlist IS Ref No# accounts"] = len(references)
        return InsuranceMasterlist(accounts, dict(stats), references)
    finally:
        original.close()
        if cached is not None:
            cached.close()


def _matching_master_contact(row: InsuranceRow, masterlist: InsuranceMasterlist) -> tuple[InsuranceContact | None, str]:
    account = _contact_text(row.record.get("primary_key"))
    reference = masterlist.references.get(account.casefold()) if account else None
    if reference is not None:
        financing = {key for key in reference.financing if key}
        if len(financing) > 1 or row.financing and financing and row.financing.casefold() not in financing:
            return None, "conflict"
        return reference.contact, "reference"
    contact = masterlist.accounts.get(row.financing.casefold()) if row.financing else None
    return contact, "financing" if contact is not None else "unmatched"


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


def prepare_insurance(
    source: InsuranceSource, included_statuses: Iterable[str], *, encoding: str = "utf-8-sig",
    masterlist: InsuranceMasterlist | None = None,
) -> InsuranceResult:
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
    records = [dict(row.record) for row in selected]
    contacts = None
    match_methods = Counter()
    if masterlist is not None:
        matches = [_matching_master_contact(row, masterlist) for row in selected]
        contacts = [contact for contact, _ in matches]
        match_methods.update(method for _, method in matches)
        for record, contact in zip(records, contacts):
            if contact is not None:
                record.update({"email": contact.email, "phone": contact.phone})
    template = create_default_csv_template(encoding=encoding)
    file = ExportFile(
        key="sbc_insurance", group_key="all", channel="insurance", label=INSURANCE_DOCLASS,
        filename="SBC INSURANCE.csv",
        data=fill_csv_template(template, records, encoding=encoding, doclass=INSURANCE_DOCLASS),
        count=len(records), missing_emails=sum(not record.get("email") for record in records), headers=list(DEFAULT_TEMPLATE_HEADERS),
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
    if contacts is not None:
        stats.update(masterlist.stats)
        stats.update({
            "Accounts matched to masterlist": sum(contact is not None for contact in contacts),
            "Accounts matched by IS Ref No#": match_methods["reference"],
            "Accounts matched by PN_Number fallback": match_methods["financing"],
            "Accounts with conflicting masterlist identifiers": match_methods["conflict"],
            "Accounts without a masterlist match": sum(contact is None for contact in contacts),
            "Accounts with valid masterlist email": sum(bool(c and c.email) for c in contacts),
            "Accounts with valid masterlist mobile": sum(bool(c and c.phone) for c in contacts),
            "Accounts with a postal address": sum(bool(c and c.addresses) for c in contacts),
            "Matched accounts with alternative valid emails": sum(bool(c and len(c.emails) > 1) for c in contacts),
            "Matched accounts with alternative valid mobiles": sum(bool(c and len(c.phones) > 1) for c in contacts),
            "Matched accounts with alternative postal addresses": sum(bool(c and len(c.addresses) > 1) for c in contacts),
        })
    return InsuranceResult(source, selected, stats, file, contacts)


def build_filtered_datagrid(result: InsuranceResult) -> bytes:
    """Keep selected source rows and append matched contacts and postal addresses."""
    workbook = Workbook(write_only=True)
    sheet = workbook.create_sheet("SBC INSURANCE")
    sheet.freeze_panes = "A2"
    headers = list(result.source.headers)
    identifier_columns = set(result.source.identifier_columns)
    if result.contacts is not None:
        identifier_columns.add(len(headers) + 1)  # The appended masterlist MOBILE.
        headers.extend(["Masterlist EMAIL", "Masterlist MOBILE", "Masterlist ADDRESS1", "Masterlist ADDRESS2", "Masterlist ADDRESS3"])

    def export_rows():
        for index, row in enumerate(result.rows):
            values = row.values
            if result.contacts is not None:
                contact = result.contacts[index]
                values = (*values, contact.email if contact else None, contact.phone if contact else None,
                          *(contact.address if contact else (None, None, None)))
            yield values

    for values in chain([headers], export_rows()):
        cells = []
        for index, value in enumerate(values):
            cell = WriteOnlyCell(sheet)
            _write_excel_value(cell, value, "primary_key" if index in identifier_columns else None)
            if isinstance(value, datetime):
                cell.number_format = "yyyy-mm-dd hh:mm:ss"
            cells.append(cell)
        sheet.append(cells)
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()
