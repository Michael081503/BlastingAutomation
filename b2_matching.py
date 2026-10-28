"""Match DataGrid accounts to SBC contacts and fill the B2 blasting template."""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, replace
from datetime import date, datetime
from typing import Any, Sequence

from automation import AutomationError, DEFAULT_TEMPLATE_HEADERS, _identifier, normalize_header
from autostat import (
    AutostatBundle, ContactRows, ContactSheet, SourceSelection,
    _add_contact_row, _bundle, _contact_summary, _log_file, _workbook,
    unique_account_contacts,
)
from export_config import AUTOSTAT_B2_FILENAMES
from exports import export_filename


B2_MATCHING_VERSION = 2
B2_TEMPLATE_VERSION = 1
# Exact header order from 1EMAIL_Blasting - Template.xlsx; no sample rows are used.
B2_TEMPLATE_HEADERS = DEFAULT_TEMPLATE_HEADERS
DATAGRID_ROUTING_VERSION = 1
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
RAW_DETAILS = {
    "FullName": ("Name", "FullName"),
    "OB": ("Total Outstanding", "OB (Masterlist)", "OB"),
    "EndoDate": ("Assignment Date", "EndoDate", "endo date"),
    "AREA": ("Area",),
    "Level": ("Level",),
    "ChristmasAmnesty": ("ChristmasAmnesty",),
    "NewEndo": ("NewEndo",),
}
GRID_DETAILS = {
    "indexoneuid": ("Old I.C.", "indexoneuid"),
    "FullName": ("Name", "FullName"),
    "OB": ("Amount Referred", "OB (Masterlist)", "OB"),
    "EndoDate": ("EndoDate", "endo date"),
    "AREA": ("Area",),
    "Level": ("Level",),
    "ChristmasAmnesty": ("ChristmasAmnesty",),
    "NewEndo": ("NewEndo",),
}


def _key(value: Any) -> str:
    text = "" if value is None else str(value).strip().casefold()
    return "" if text in {"", "-", "n/a", "#n/a", "null", "none"} else text


def is_datagrid_sheet(headers: Sequence[str]) -> bool:
    """Recognize the account-list export by its columns, independent of its name."""
    labels = {normalize_header(header) for header in headers}
    return {"accountno", "financingcardno"} <= labels and (
        bool(labels & {"statuscode", "clientname"})
        or {"no", "name", "newic"} <= labels
    )


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


def _detail_columns(source: SourceSelection, fields: dict[str, tuple[str, ...]]) -> dict[str, int]:
    labels = [normalize_header(header) for header in source.sheet.headers]
    result = {}
    for target, aliases in fields.items():
        for alias in aliases:
            matches = [index for index, label in enumerate(labels) if label == normalize_header(alias)]
            if len(matches) > 1:
                raise AutomationError(f"{source.name} / {source.sheet.name}: ambiguous {alias} column.")
            if matches:
                result[target] = matches[0]
                break
    return result


def _campaign_record(raw: dict, grid: dict, raw_columns: dict, raw_details: dict, grid_details: dict, group: str) -> dict:
    record = dict.fromkeys(B2_TEMPLATE_HEADERS, "")
    for target in set(raw_details) | set(grid_details):
        value = raw.get(raw_details.get(target))
        if value is None or isinstance(value, str) and not value.strip():
            value = grid.get(grid_details.get(target))
        record[target] = value
    record.update({
        "primary_key": raw[raw_columns["account_key"]],
        "cardnumber": raw[raw_columns["contract"]],
        "doclass": f"SBC {group} B2",
        "Product": group,
        "Monthly": "",
        "Terms": "",
    })
    return record


def _template_text(value: Any) -> str:
    if isinstance(value, datetime):
        return value.isoformat(sep=" ")
    if isinstance(value, date):
        return value.isoformat()
    return "" if value is None else str(value)


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
    database: SourceSelection, datagrid: SourceSelection, *, encoding: str = "cp1252",
) -> B2MatchResult:
    """Fill the 16-column blasting layout once per verified DataGrid account."""
    if set(AUTOSTAT_B2_FILENAMES) != {"CARDS", "BEL"}:
        raise AutomationError("Set both CARDS and BEL B2 filenames in export_config.py.")
    names = {group: export_filename(name, ".csv") for group, name in AUTOSTAT_B2_FILENAMES.items()}
    if len({name.casefold() for name in names.values()}) != 2:
        raise AutomationError("The B2 output filenames must be unique. Check export_config.py.")
    raw_columns = _columns(database, database=True)
    grid_columns = _columns(datagrid, database=False)
    raw_details = _detail_columns(database, RAW_DETAILS)
    grid_details = _detail_columns(datagrid, GRID_DETAILS)
    contacts_sheet = replace(database.sheet, account_column=raw_columns["contract"])
    selected_raw = (set(raw_columns.values()) | set(raw_details.values())
                    | set(contacts_sheet.email_columns) | set(contacts_sheet.phone_columns))
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
    records = {"CARDS": [], "BEL": []}
    audit, seen_accounts = [], set()
    grid_identifiers = set(grid_columns.values())
    if "indexoneuid" in grid_details:
        grid_identifiers.add(grid_details["indexoneuid"])
    for number, row in _rows(datagrid, set(grid_columns.values()) | set(grid_details.values()), grid_identifiers):
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
                    records[group].append(_campaign_record(matched, row, raw_columns, raw_details, grid_details, group))
        if reason:
            stats[reason] += 1
            audit.append({
                "DataGrid row": number,
                "Account No.": row.get(grid_columns.get("account_key"), ""),
                "Financing/Card No.": row.get(grid_columns.get("contract"), ""),
                "Reason": reason,
            })

    files, summaries = [], []
    for group, contacts in groups.items():
        contacts = unique_account_contacts(contacts)
        emails = {_key(account): value for account, value in contacts.emails}
        phones = {_key(account): value for account, value in contacts.phones}
        rows = []
        for record in records[group]:
            account = _key(record["cardnumber"])
            if account not in emails and account not in phones:
                continue
            record.update({"email": emails.get(account, ""), "phone": phones.get(account, "")})
            rows.append({header: _template_text(record[header]) for header in B2_TEMPLATE_HEADERS})
        summaries.append({"Group": group, "Accounts exported": len(rows), **_contact_summary(database, contacts)})
        files.append(_log_file(names[group], B2_TEMPLATE_HEADERS, rows, encoding))
    return B2MatchResult(_bundle(files, summaries), stats, audit)
