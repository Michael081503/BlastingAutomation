"""Build B2 activity logs using the supplied Email, SMS and Viber layouts."""

from __future__ import annotations

from dataclasses import dataclass, replace
from datetime import datetime

from automation import AutomationError, normalize_header, ob_sort_key, source_column_positions
from autostat import (
    AutostatBundle, ContactRows, ContactSheet, SourceSelection,
    _add_contact_row, _bundle, _contact_summary, _log_file, _read_rows,
    account_number_text, unique_account_contacts,
)


B2_AUTOSTAT_VERSION = 4
B2_PRIMARY_KEY_VERSION = 1
B2_COLLECTOR_TERMS_VERSION = 1
B2_EMAIL_HEADERS = (
    "Name", "Account No.", "STATUS CODE", "REMARKS", "REMARKS BY", "REMARKS DATE", "Email", "Collector",
)
B2_PHONE_HEADERS = (
    "Account number", "Name", "STATUS CODE", "REMARKS", "REMARKS BY", "REMARKS DATE", "Mobile Number",
)
B2_CHANNELS = ("Email", "SMS", "Viber")


@dataclass(frozen=True)
class B2ActivitySource:
    source: SourceSelection
    name_column: int | None = None
    collector_column: int | None = None


def b2_account_sheet(sheet: ContactSheet) -> ContactSheet:
    """Use primary_key for B2 accounts whenever that source header is present."""
    columns = [index for index, header in enumerate(sheet.headers) if normalize_header(header) == "primarykey"]
    if len(columns) > 1:
        raise AutomationError(f"{sheet.name}: more than one primary_key column was found. Keep one unambiguous primary_key header.")
    return replace(sheet, account_column=columns[0]) if columns else sheet


def b2_metadata_columns(headers: tuple[str, ...]) -> tuple[int | None, int | None]:
    labels = [normalize_header(header) for header in headers]

    def find(aliases):
        return next((labels.index(alias) for alias in aliases if labels.count(alias) == 1), None)

    return find(("fullname", "name", "debtorname", "customername")), find(("terms", "collector", "collectory", "collectorcode"))


def _text(value) -> str:
    return "" if value is None else str(value).strip()


def build_b2_autostat_exports(
    sources: list[B2ActivitySource], *, timestamp: datetime | None,
    remark_by: str = "PBBALUYOT", default_collector: str = "", encoding: str = "cp1252",
) -> AutostatBundle:
    """Fill template fields from current uploads, keeping one contact per account and identity."""
    if not sources:
        raise AutomationError("Select at least one B2 source sheet.")
    if not remark_by.strip():
        raise AutomationError("Enter Remarks By before preparing the B2 files.")
    time_text = "" if timestamp is None else timestamp.strftime("%Y-%m-%d %H:%M:%S")
    channels = {"Email": [], "Phone": []}
    seen = {"email": set(), "viber": set()}
    summaries = []
    for selection in sources:
        source = replace(selection.source, sheet=b2_account_sheet(selection.source.sheet))
        labels = [normalize_header(header) for header in source.sheet.headers]
        identity_column = next((labels.index(alias) for alias in ("indexoneuid", "oldic")
                                if labels.count(alias) == 1), None)
        ob_column = source_column_positions(source.sheet.headers).get("OB")
        extra_columns = tuple(dict.fromkeys(column for column in (
            selection.name_column, selection.collector_column, identity_column, ob_column,
        ) if column is not None))
        if source.sheet.account_column in (selection.name_column, selection.collector_column):
            raise AutomationError("Choose separate columns for the account number, name and collector.")
        contacts = ContactRows()
        metadata = {"Email": {}, "Phone": {}}
        account_issues = []
        for row in _read_rows(
            source, extra_columns=extra_columns, account_issues=account_issues,
            row_errors=contacts.source_error_rows,
        ):
            account = row[source.sheet.account_column]
            row[source.sheet.account_column] = account_number_text(account)
            email_start, phone_start = len(contacts.emails), len(contacts.phones)
            _add_contact_row(contacts, source.sheet, row, identity=row.get(identity_column), ob=row.get(ob_column))
            details = (_text(row.get(selection.name_column)),
                       _text(row.get(selection.collector_column)) or default_collector.strip())
            for channel, keys in (
                ("Email", contacts.email_account_keys[email_start:]),
                ("Phone", contacts.phone_account_keys[phone_start:]),
            ):
                for key in keys:
                    metadata[channel].setdefault(key, details)
        contacts.account_precision_issues = len(account_issues)
        contacts = unique_account_contacts(contacts, seen=seen)
        summaries.append(_contact_summary(source, contacts))
        for channel, pairs, keys, balances in (
            ("Email", contacts.emails, contacts.email_account_keys, contacts.email_ob_values),
            ("Phone", contacts.phones, contacts.phone_account_keys, contacts.phone_ob_values),
        ):
            channels[channel].extend((account, contact, *metadata[channel][key], balance)
                                     for (account, contact), key, balance in zip(pairs, keys, balances))
    for entries in channels.values():
        entries.sort(key=lambda item: ob_sort_key(item[-1]), reverse=True)

    def activity_rows(channel):
        entries = channels["Email" if channel == "Email" else "Phone"]
        for account, contact, name, collector, _ in entries:
            row = {
                "Name": name, "STATUS CODE": f"{channel.upper()} BLAST SENT - WAITING FOR REPLY",
                "REMARKS BY": remark_by.strip(), "REMARKS DATE": time_text,
            }
            if channel == "Email":
                row.update({
                    "Account No.": account, "Email": contact, "Collector": collector,
                    "REMARKS": f"{contact} | EMAIL SENT | {collector}",
                })
            else:
                # The B2 templates use local eleven-digit mobile numbers.
                mobile = "0" + contact
                ending = "  SMS SENT" if channel == "SMS" else " VIBER SENT"
                row.update({"Account number": account, "Mobile Number": mobile,
                            "REMARKS": f"{mobile} |{ending}"})
            yield row

    return _bundle([
        _log_file(f"AUTOSTAT_B2_{channel}.csv", B2_EMAIL_HEADERS if channel == "Email" else B2_PHONE_HEADERS,
                  activity_rows(channel), encoding)
        for channel in B2_CHANNELS
    ], summaries)
