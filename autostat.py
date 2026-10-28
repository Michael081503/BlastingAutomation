"""Extract account contacts and prepare the user's AUTOSTAT activity logs."""

from __future__ import annotations

import csv
import math
import re
from dataclasses import dataclass, field, replace
from datetime import datetime
from decimal import Decimal, InvalidOperation
from io import BytesIO, StringIO
from itertools import islice
from pathlib import Path
from typing import Any
from zipfile import BadZipFile, ZIP_DEFLATED, ZipFile

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from automation import AutomationError, _encode_csv, _identifier, normalize_header
from email_validation import validated_email
from export_config import AUTOSTAT_BEL_FILENAMES
from exports import export_filename


AUTOSTAT_VERSION = 5
ACCOUNT_DEDUP_VERSION = 1
PREFERRED_ACCOUNT_HEADER = "cardnumber"
REMARK_PREFIX = "SPMA | 08 With SMS / email / DL without response"
# Column labels, statuses and remark text come from the supplied templates.
# The sample account, email and historical timestamp are deliberately omitted.
EMAIL_HEADERS = (
    "Time", "Account No.", "Status", "Remark", "Remark By", "EMAIL/ NUMBER",
    "PTP Date", "Claim Paid Amount", "Claim Paid Date", "Dialed Number", REMARK_PREFIX,
)
VIBER_HEADERS = ("Time", "Account No.", "Status", "Remark", "Remark By")
ACCOUNT_ALIASES = (
    PREFERRED_ACCOUNT_HEADER, "financing/Card no.", "Contract Number", "primary_key", "Account No.", "Account Key", "account", "account number", "account #",
    "acct no", "acct number", "accountno", "account id",
)
MIXED_ALIASES = {"emailnumber", "emailphone", "contact", "contacts", "contactdetails"}


def validated_phone(value: Any) -> str | None:
    """Keep the rightmost ten digits only when there are ten and they start at 9."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (not math.isfinite(value) or not value.is_integer()):
            return None
        text = str(int(value))
    elif isinstance(value, str):
        text = value.strip()
        # CSV exports sometimes render integral Excel numbers as .0 or E notation.
        if re.fullmatch(r"\+?\d+(?:\.\d+)?[eE][+-]?\d{1,3}|\+?\d+\.0+", text):
            try:
                number = Decimal(text)
                if number != number.to_integral_value() or not 0 <= number < 10**30:
                    return None
                text = str(int(number))
            except InvalidOperation:
                return None
    else:
        return None
    if not re.fullmatch(r"\+?[0-9(][0-9\s().-]*", text):
        return None
    digits = re.sub(r"[^0-9]", "", text)
    if len(digits) < 10:
        return None
    phone = digits[-10:]
    return phone if phone.startswith("9") else None


def _blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def _contact_kind(header: Any) -> tuple[bool, bool]:
    label = normalize_header(header)
    mixed = label in MIXED_ALIASES
    email = mixed or bool(re.fullmatch(r"(?:debtor)?emails?(?:address)?\d*", label))
    phone = mixed or bool(re.fullmatch(
        r"(?:debtor|home|business|office|work)?(?:phone|mobile|cellphone|cell|telephone|tel|contact|viber|sms)"
        r"(?:number|no)?\d*", label,
    ))
    return email, phone


@dataclass(frozen=True)
class ContactSheet:
    name: str
    header_row: int
    headers: tuple[str, ...]
    account_column: int | None
    email_columns: tuple[int, ...]
    phone_columns: tuple[int, ...]

    @property
    def ready(self) -> bool:
        return self.account_column is not None and bool(self.email_columns or self.phone_columns)


def _header_score(row) -> int:
    labels = {normalize_header(value) for value in row}
    account = any(normalize_header(alias) in labels for alias in ACCOUNT_ALIASES)
    kinds = [_contact_kind(value) for value in row]
    return 4 * account + any(kind[0] for kind in kinds) + any(kind[1] for kind in kinds)


def _sheet_info(name: str, rows: list[tuple[Any, ...]], header_row: int | None) -> ContactSheet:
    if header_row is None:
        header_row = max(range(len(rows)), key=lambda index: _header_score(rows[index])) + 1 if rows else 1
        headers = rows[header_row - 1] if rows else ()
    else:
        headers = rows[0] if rows else ()
    labels = [normalize_header(value) for value in headers]
    account_column = next((labels.index(normalize_header(alias)) for alias in ACCOUNT_ALIASES
                           if normalize_header(alias) in labels), None)
    kinds = [_contact_kind(value) for value in headers]
    return ContactSheet(
        name, header_row, tuple("" if value is None else str(value) for value in headers),
        account_column, tuple(i for i, kind in enumerate(kinds) if kind[0]),
        tuple(i for i, kind in enumerate(kinds) if kind[1]),
    )


def _csv_rows(data: bytes):
    try:
        text = data.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = data.decode("cp1252")
    try:
        delimiter = csv.Sniffer().sniff(text[:8192], delimiters=",;\t|").delimiter
    except csv.Error:
        # A title before the header can defeat Sniffer's equal-width heuristic.
        scores = {
            candidate: max((_header_score(row) for row in islice(
                csv.reader(StringIO(text, newline=""), delimiter=candidate), 50,
            )), default=0)
            for candidate in (",", ";", "\t", "|")
        }
        delimiter = max(scores, key=scores.get)
    return csv.reader(StringIO(text, newline=""), delimiter=delimiter)


def _workbook(data: bytes, *, data_only: bool = False):
    try:
        workbook = load_workbook(BytesIO(data), read_only=True, data_only=data_only)
        for sheet in workbook.worksheets:
            sheet.reset_dimensions()
        return workbook
    except (BadZipFile, InvalidFileException, KeyError, OSError, ValueError) as exc:
        raise AutomationError("Cannot read this workbook. Save an unencrypted .xlsx copy in Excel and try again.") from exc


def inspect_contact_source(
    data: bytes, extension: str, *, sheet_name: str | None = None, header_row: int | None = None,
) -> list[ContactSheet]:
    """Find account/contact columns without requiring any campaign fields."""
    if not data:
        raise AutomationError("The source file is empty.")
    if header_row is not None and not 1 <= header_row <= 1_048_575:
        raise AutomationError("The header row must be between 1 and 1,048,575.")
    if extension == ".csv":
        rows = []
        for number, row in enumerate(_csv_rows(data), 1):
            if header_row is None or number == header_row:
                rows.append(tuple(row))
            if number >= (header_row or 50):
                break
        return [_sheet_info("CSV", rows, header_row)]
    if extension != ".xlsx":
        raise AutomationError("Upload an .xlsx workbook or .csv source file.")
    workbook = _workbook(data)
    try:
        sheets = workbook.worksheets if sheet_name is None else [workbook[sheet_name]]
        return [_sheet_info(sheet.title, list(sheet.iter_rows(
            min_row=header_row or 1, max_row=header_row or min(sheet.max_row or 50, 50),
            values_only=True,
        )), header_row) for sheet in sheets]
    finally:
        workbook.close()


@dataclass(frozen=True)
class SourceSelection:
    name: str
    data: bytes
    sheet: ContactSheet


@dataclass
class ContactRows:
    emails: list[tuple[str, str]] = field(default_factory=list)
    phones: list[tuple[str, str]] = field(default_factory=list)
    checked_rows: int = 0
    blank_rows: int = 0
    missing_accounts: int = 0
    no_contacts: int = 0
    invalid_emails: int = 0
    invalid_phones: int = 0
    duplicate_email_rows: int = 0
    duplicate_phone_rows: int = 0


def _candidates(value: Any, validator) -> list[Any]:
    if _blank(value):
        return []
    if validator(value) is not None or not isinstance(value, str):
        return [value]
    delimiter = r"[;,|\r\n/]+" if validator is validated_phone else r"[;,|\r\n]+"
    return [part.strip() for part in re.split(delimiter, value) if part.strip()]


def _read_rows(source: SourceSelection, *, extra_columns: tuple[int, ...] = ()):
    """Yield selected values; do not evaluate source formulas or follow links."""
    sheet = source.sheet
    columns = tuple(dict.fromkeys((sheet.account_column, *sheet.email_columns, *sheet.phone_columns, *extra_columns)))
    if sheet.account_column is None or not (sheet.email_columns or sheet.phone_columns):
        raise AutomationError(f"{source.name} / {sheet.name}: select an account column and at least one contact column.")
    if any(index is None or index < 0 or index >= len(sheet.headers) for index in columns):
        raise AutomationError("The selected columns do not match the source header row.")
    if sheet.account_column in (*sheet.email_columns, *sheet.phone_columns):
        raise AutomationError("Choose separate columns for account numbers and contacts.")
    if Path(source.name).suffix.lower() == ".csv":
        for number, row in enumerate(_csv_rows(source.data), 1):
            if number > sheet.header_row:
                yield {index: (row[index] if index < len(row) else None) for index in columns}
        return
    original = _workbook(source.data)
    cached = None
    try:
        cached = _workbook(source.data, data_only=True)
        options = {"min_row": sheet.header_row + 1, "max_col": max(columns) + 1}
        for number, (formulas, values) in enumerate(zip(
            original[sheet.name].iter_rows(**options), cached[sheet.name].iter_rows(**options),
        ), sheet.header_row + 1):
            row = {}
            for index in columns:
                cell = values[index]
                if formulas[index].data_type == "f" and cell.value is None:
                    raise AutomationError(
                        f"{source.name} / {sheet.name}, row {number}: a formula has no saved value. "
                        "Recalculate and save the source in Excel before uploading again."
                    )
                if cell.data_type == "e":
                    raise AutomationError(f"{source.name} / {sheet.name}, row {number}: fix Excel error {cell.value}.")
                row[index] = (_identifier(cell.value, cell.number_format)
                              if index == sheet.account_column else cell.value)
            yield row
    finally:
        original.close()
        if cached is not None:
            cached.close()


def _add_contact_row(result: ContactRows, sheet: ContactSheet, row: dict[int, Any]) -> None:
    if all(_blank(value) for value in row.values()):
        result.blank_rows += 1
        return
    result.checked_rows += 1
    emails, phones = [], []
    for columns, validator, contacts, invalid_field in (
        (sheet.email_columns, validated_email, emails, "invalid_emails"),
        (sheet.phone_columns, validated_phone, phones, "invalid_phones"),
    ):
        for column in columns:
            # A shared EMAIL/NUMBER column may legitimately contain either kind.
            other_validator = validated_phone if validator is validated_email else validated_email
            mixed = column in sheet.email_columns and column in sheet.phone_columns
            for value in _candidates(row[column], validator):
                contact = validator(value)
                if contact is not None:
                    if contact not in contacts:
                        contacts.append(contact)
                elif not (mixed and other_validator(value) is not None):
                    setattr(result, invalid_field, getattr(result, invalid_field) + 1)
    account = row[sheet.account_column]
    if _blank(account):
        result.missing_accounts += 1
        return
    account = str(account).strip()
    if not emails and not phones:
        result.no_contacts += 1
    result.emails.extend((account, email) for email in emails)
    result.phones.extend((account, phone) for phone in phones)


def extract_contacts(source: SourceSelection) -> ContactRows:
    """Keep source order and one log row per distinct valid contact in each row."""
    result = ContactRows()
    for row in _read_rows(source):
        _add_contact_row(result, source.sheet, row)
    return result


def unique_account_contacts(
    contacts: ContactRows, *, seen: dict[str, set[str]] | None = None,
) -> ContactRows:
    """Keep the first valid contact per account independently in each channel."""
    if seen is None:
        seen = {"email": set(), "viber": set()}
    unique = {}
    removed = {}
    for channel, pairs in (("email", contacts.emails), ("viber", contacts.phones)):
        unique[channel] = []
        removed[channel] = 0
        for account, contact in pairs:
            key = account.strip().casefold()
            if key in seen[channel]:
                removed[channel] += 1
            else:
                seen[channel].add(key)
                unique[channel].append((account, contact))
    return replace(
        contacts, emails=unique["email"], phones=unique["viber"],
        duplicate_email_rows=contacts.duplicate_email_rows + removed["email"],
        duplicate_phone_rows=contacts.duplicate_phone_rows + removed["viber"],
    )


def split_bel_contacts(source: SourceSelection) -> dict[str, ContactRows]:
    """Separate rows whose primary_key contains BEL, independently of the account column."""
    positions = [index for index, header in enumerate(source.sheet.headers)
                 if normalize_header(header) == "primarykey"]
    if len(positions) != 1:
        raise AutomationError(
            f"{source.name} / {source.sheet.name}: BEL separation requires one primary_key column."
        )
    primary_key = positions[0]
    groups = {"BEL": ContactRows(), "Other": ContactRows()}
    for row in _read_rows(source, extra_columns=(primary_key,)):
        key = "BEL" if "BEL" in str(row[primary_key] or "").upper() else "Other"
        _add_contact_row(groups[key], source.sheet, row)
    return groups


def has_bel_primary_keys(source: SourceSelection) -> bool:
    """Detect BEL anywhere in the selected sheet, without treating account IDs as contacts."""
    positions = [index for index, header in enumerate(source.sheet.headers)
                 if normalize_header(header) == "primarykey"]
    if len(positions) != 1:
        return False
    column = positions[0]
    if Path(source.name).suffix.lower() == ".csv":
        return any(number > source.sheet.header_row and len(row) > column
                   and "BEL" in str(row[column] or "").upper()
                   for number, row in enumerate(_csv_rows(source.data), 1))
    workbook = _workbook(source.data, data_only=True)
    try:
        return any("BEL" in str(row[0] or "").upper() for row in workbook[source.sheet.name].iter_rows(
            min_row=source.sheet.header_row + 1, min_col=column + 1, max_col=column + 1,
            values_only=True,
        ))
    finally:
        workbook.close()


@dataclass
class AutostatFile:
    filename: str
    data: bytes
    count: int
    headers: tuple[str, ...]
    preview: list[dict[str, str]]


@dataclass
class AutostatBundle:
    files: list[AutostatFile]
    summary: list[dict[str, Any]]
    zip_data: bytes


def _contact_summary(source: SourceSelection, contacts: ContactRows) -> dict[str, Any]:
    return {
        "File": source.name, "Sheet": source.sheet.name, "Rows checked": contacts.checked_rows,
        "Email rows": len(contacts.emails), "Viber rows": len(contacts.phones),
        "Invalid emails": contacts.invalid_emails, "Invalid phones": contacts.invalid_phones,
        "Missing account": contacts.missing_accounts, "No valid contact": contacts.no_contacts,
        "Empty rows": contacts.blank_rows,
        "Duplicate Email rows removed": contacts.duplicate_email_rows,
        "Duplicate Viber rows removed": contacts.duplicate_phone_rows,
    }


def _activity_rows(pairs, *, channel: str, headers: tuple[str, ...], time_text: str, remark_by: str):
    status = "EMAIL BLAST SENT - WAITING FOR REPLY" if channel == "Email" else "SMS BLAST SENT - WAITING FOR REPLY"
    ending = "FOLLOW UP EMAIL SENT" if channel == "Email" else "VIBER SENT - BLASTING"
    for account, contact in pairs:
        row = dict.fromkeys(headers, "")
        row.update({
            "Time": time_text, "Account No.": account, "Status": status,
            "Remark": f"{REMARK_PREFIX} - {contact} - {ending}", "Remark By": remark_by.strip(),
        })
        if "EMAIL/ NUMBER" in row:
            row["EMAIL/ NUMBER"] = contact
        yield row


def _log_file(filename: str, headers: tuple[str, ...], rows, encoding: str) -> AutostatFile:
    output = StringIO(newline="")
    writer = csv.DictWriter(output, fieldnames=headers, lineterminator="\r\n")
    writer.writeheader()
    count, preview = 0, []
    for row in rows:
        writer.writerow(row)
        count += 1
        if len(preview) < 100:
            preview.append(row)
    return AutostatFile(filename, _encode_csv(output.getvalue(), encoding), count, headers, preview)


def _bundle(files: list[AutostatFile], summaries: list[dict[str, Any]]) -> AutostatBundle:
    output = BytesIO()
    with ZipFile(output, "w", ZIP_DEFLATED) as archive:
        for file in files:
            archive.writestr(file.filename, file.data)
    return AutostatBundle(files, summaries, output.getvalue())


def build_bel_exports(
    sources: list[SourceSelection], *, timestamp: datetime, remark_by: str = "MBURGOS",
    encoding: str = "cp1252",
) -> AutostatBundle:
    """Create exactly two group logs, retaining Email and Viber status/remark text."""
    if not sources:
        raise AutomationError("Select at least one source sheet.")
    if not remark_by.strip():
        raise AutomationError("Enter Remark By before preparing the files.")
    if set(AUTOSTAT_BEL_FILENAMES) != {"BEL", "Other"}:
        raise AutomationError("Set both BEL and Other filenames in export_config.py.")
    names = {group: export_filename(name, ".csv") for group, name in AUTOSTAT_BEL_FILENAMES.items()}
    if len({name.casefold() for name in names.values()}) != 2:
        raise AutomationError("BEL and CARDS output filenames must be unique. Check export_config.py.")
    time_text = timestamp.strftime("%Y-%m-%d %I:%M:%S %p")
    rows, summaries = {"BEL": [], "Other": []}, []
    seen = {"email": set(), "viber": set()}
    for source in sources:
        for group, contacts in split_bel_contacts(source).items():
            contacts = unique_account_contacts(contacts, seen=seen)
            summaries.append({"Group": "BEL" if group == "BEL" else "CARDS", **_contact_summary(source, contacts)})
            for channel, pairs in (("Email", contacts.emails), ("Viber", contacts.phones)):
                rows[group].extend(_activity_rows(
                    pairs, channel=channel, headers=EMAIL_HEADERS, time_text=time_text, remark_by=remark_by,
                ))
    files = [_log_file(names[group], EMAIL_HEADERS, rows[group], encoding) for group in rows]
    return _bundle(files, summaries)


def build_autostat_exports(
    sources: list[SourceSelection], *, timestamp: datetime, remark_by: str = "MBURGOS",
    encoding: str = "cp1252",
) -> AutostatBundle:
    if not sources:
        raise AutomationError("Select at least one source sheet.")
    if not remark_by.strip():
        raise AutomationError("Enter Remark By before preparing the files.")
    time_text = timestamp.strftime("%Y-%m-%d %I:%M:%S %p")
    files, summaries, used_names = [], [], set()
    seen = {"email": set(), "viber": set()}
    for source in sources:
        contacts = unique_account_contacts(extract_contacts(source), seen=seen)
        summaries.append(_contact_summary(source, contacts))
        label = f"{Path(source.name).stem}_{source.sheet.name}"
        label = re.sub(r'[^\w .-]+', "_", label).strip(" .")[:110].rstrip(" .") or "source"
        base, suffix = label, 2
        while base.casefold() in used_names:
            base = f"{label}_{suffix}"
            suffix += 1
        used_names.add(base.casefold())
        for channel, pairs, headers in (
            ("Email", contacts.emails, EMAIL_HEADERS),
            ("Viber", contacts.phones, VIBER_HEADERS),
        ):
            files.append(_log_file(
                f"AUTOSTAT_{base}_{channel}.csv", headers,
                _activity_rows(pairs, channel=channel, headers=headers, time_text=time_text, remark_by=remark_by),
                encoding,
            ))
    return _bundle(files, summaries)
