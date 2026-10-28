"""Match the DataGrid account list to SBC B2 contact records for AUTOSTAT."""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, replace
from datetime import datetime
from typing import Any

from automation import AutomationError, _identifier, normalize_header
from autostat import (
    EMAIL_HEADERS, AutostatBundle, ContactRows, ContactSheet, SourceSelection,
    _activity_rows, _add_contact_row, _bundle, _contact_summary, _log_file, _workbook,
    unique_account_contacts,
)
from export_config import AUTOSTAT_B2_FILENAMES
from exports import export_filename


B2_MATCHING_VERSION = 1
RAW_FIELDS = {
    "account_key": ("Account Key",),
    "contract": ("Contract Number",),
    "bucket": ("Bucket Code",),
    "category": ("Product Category Code",),
}
GRID_FIELDS = {
    "account_key": ("Account No.", "Account Number", "primary_key"),
    "contract": ("Financing/Card No.", "cardnumber", "Contract Number"),
}


def _key(value: Any) -> str:
    text = "" if value is None else str(value).strip().casefold()
    return "" if text in {"", "-", "n/a", "#n/a", "null", "none"} else text


def matching_sheet_score(sheet: ContactSheet, *, database: bool) -> int:
    labels = {normalize_header(value) for value in sheet.headers}
    groups = RAW_FIELDS if database else GRID_FIELDS
    return sum(bool(labels & {normalize_header(alias) for alias in aliases}) for aliases in groups.values())


def _columns(source: SourceSelection, *, database: bool) -> dict[str, int]:
    labels = [normalize_header(header) for header in source.sheet.headers]
    result = {}
    for field, aliases in (RAW_FIELDS if database else GRID_FIELDS).items():
        matches = [index for index, label in enumerate(labels)
                   if label in {normalize_header(alias) for alias in aliases}]
        if len(matches) > 1:
            raise AutomationError(f"{source.name} / {source.sheet.name}: ambiguous {aliases[0]} column.")
        if matches:
            result[field] = matches[0]
        elif database:
            raise AutomationError(f"{source.name} / {source.sheet.name}: missing {aliases[0]}. Choose the full SBC database sheet.")
    if not database and not result:
        raise AutomationError("The DataGrid sheet needs Account No. or Financing/Card No. for matching.")
    if database and not (source.sheet.email_columns or source.sheet.phone_columns):
        raise AutomationError("The SBC database sheet has no recognized email or phone columns.")
    return result


def _rows(
    source: SourceSelection, columns: set[int], identifier_columns: set[int],
    *, contact_columns: set[int] | None = None,
):
    """Read actual XML rows, checking saved formula values and identifier masks."""
    original = _workbook(source.data)
    cached = None
    try:
        cached = _workbook(source.data, data_only=True)
        options = {"min_row": source.sheet.header_row + 1, "max_col": max(columns) + 1}
        for number, (formulas, values) in enumerate(zip(
            original[source.sheet.name].iter_rows(**options), cached[source.sheet.name].iter_rows(**options),
        ), source.sheet.header_row + 1):
            row = {}
            for index in columns:
                original_cell, cell = formulas[index], values[index]
                if original_cell.data_type == "f" and cell.value is None:
                    if contact_columns and index in contact_columns:
                        row[index] = "[Formula without saved value]"
                        continue
                    raise AutomationError(
                        f"{source.name} / {source.sheet.name}, row {number}: a formula has no saved value. "
                        "Recalculate and save the workbook in Excel."
                    )
                if cell.data_type == "e":
                    if contact_columns and index in contact_columns:
                        row[index] = str(cell.value)
                        continue
                    raise AutomationError(f"{source.name} / {source.sheet.name}, row {number}: fix Excel error {cell.value}.")
                row[index] = _identifier(cell.value, cell.number_format) if index in identifier_columns else cell.value
            yield number, row
    finally:
        original.close()
        if cached is not None:
            cached.close()


@dataclass
class B2MatchResult:
    bundle: AutostatBundle
    stats: dict[str, int]
    audit: list[dict[str, Any]]


def build_b2_exports(
    database: SourceSelection, datagrid: SourceSelection, *, timestamp: datetime,
    remark_by: str = "MBURGOS", encoding: str = "cp1252",
) -> B2MatchResult:
    """Export only verified DataGrid accounts, using contacts from the full SBC sheet."""
    if not remark_by.strip():
        raise AutomationError("Enter Remark By before preparing the files.")
    if set(AUTOSTAT_B2_FILENAMES) != {"CARDS", "BEL"}:
        raise AutomationError("Set both CARDS and BEL B2 filenames in export_config.py.")
    names = {group: export_filename(name, ".csv") for group, name in AUTOSTAT_B2_FILENAMES.items()}
    if len({name.casefold() for name in names.values()}) != 2:
        raise AutomationError("The B2 output filenames must be unique. Check export_config.py.")
    raw_columns = _columns(database, database=True)
    grid_columns = _columns(datagrid, database=False)
    contacts_sheet = replace(database.sheet, account_column=raw_columns["contract"])
    selected_raw = set(raw_columns.values()) | set(contacts_sheet.email_columns) | set(contacts_sheet.phone_columns)
    raw_rows = list(_rows(
        database, selected_raw, {raw_columns["account_key"], raw_columns["contract"]},
        contact_columns=set(contacts_sheet.email_columns) | set(contacts_sheet.phone_columns),
    ))
    indexes = {field: defaultdict(set) for field in ("account_key", "contract")}
    for index, (_, row) in enumerate(raw_rows):
        for field in indexes:
            key = _key(row[raw_columns[field]])
            if key:
                indexes[field][key].add(index)

    stats = dict.fromkeys((
        "Database rows", "DataGrid rows", "Matched rows", "Unique B2 accounts",
        "Unmatched rows", "Missing account numbers", "Conflicting or ambiguous keys",
        "Other or missing bucket", "Unsupported product", "Duplicate account rows removed",
    ), 0)
    stats["Database rows"] = len(raw_rows)
    groups = {"CARDS": ContactRows(), "BEL": ContactRows()}
    audit, seen_accounts = [], set()
    for number, row in _rows(datagrid, set(grid_columns.values()), set(grid_columns.values())):
        stats["DataGrid rows"] += 1
        keys = {field: _key(row[column]) for field, column in grid_columns.items()}
        keys = {field: value for field, value in keys.items() if value}
        reason = None
        if not keys:
            reason = "Missing account numbers"
            candidates = set()
        else:
            candidates = set.intersection(*(indexes[field].get(value, set()) for field, value in keys.items()))
            if not candidates:
                reason = ("Conflicting or ambiguous keys" if any(indexes[field].get(value) for field, value in keys.items())
                          else "Unmatched rows")
            elif len(candidates) != 1:
                reason = "Conflicting or ambiguous keys"
        if reason is None:
            _, matched = raw_rows[next(iter(candidates))]
            stats["Matched rows"] += 1
            contract = _key(matched[raw_columns["contract"]])
            if _key(matched[raw_columns["bucket"]]) not in {"2", "2.0", "b2"}:
                reason = "Other or missing bucket"
            else:
                category = _key(matched[raw_columns["category"]])
                group = "BEL" if category == "bel" else "CARDS" if category in {"mc", "visa", "cards", "card", "mastercard"} else None
                if group is None:
                    reason = "Unsupported product"
                elif not contract:
                    reason = "Missing account numbers"
                elif contract in seen_accounts:
                    reason = "Duplicate account rows removed"
                else:
                    seen_accounts.add(contract)
                    stats["Unique B2 accounts"] += 1
                    _add_contact_row(groups[group], contacts_sheet, matched)
        if reason:
            stats[reason] += 1
            audit.append({
                "DataGrid row": number,
                "Account No.": row.get(grid_columns.get("account_key"), ""),
                "Financing/Card No.": row.get(grid_columns.get("contract"), ""),
                "Reason": reason,
            })

    files, summaries = [], []
    time_text = timestamp.strftime("%Y-%m-%d %I:%M:%S %p")
    for group, contacts in groups.items():
        contacts = unique_account_contacts(contacts)
        summaries.append({"Group": group, **_contact_summary(database, contacts)})
        rows = []
        for channel, pairs in (("Email", contacts.emails), ("Viber", contacts.phones)):
            rows.extend(_activity_rows(pairs, channel=channel, headers=EMAIL_HEADERS,
                                       time_text=time_text, remark_by=remark_by))
        files.append(_log_file(names[group], EMAIL_HEADERS, rows, encoding))
    return B2MatchResult(_bundle(files, summaries), stats, audit)
