"""Resolve inbound campaigns by email or phone from active-account workbooks."""

from __future__ import annotations

import csv
import os
import re
from dataclasses import dataclass, replace
from collections import Counter
from hashlib import sha256
from io import BytesIO, StringIO
from pathlib import Path
from tempfile import NamedTemporaryFile
from zipfile import BadZipFile, ZipFile

from inbound_email import EmailSource, header_key, normalized_email
from inbound_reply import (
    PL_RECOVERY_CAMPAIGN, SBLEGAL_CAMPAIGN, InboundReplyError, ReplyRecord, ReplySource, excluded_inbound_campaign,
    inbound_campaign_label, inbound_csv_delimiter, client_account_key, normalized_campaign, normalized_phone,
)


EMAIL_CAMPAIGN_LOOKUP_VERSION = 9
PHONE_CAMPAIGN_LOOKUP_VERSION = 2
BUCKET_CAMPAIGN_LOOKUP_VERSION = 1
B4_PRIMARY_EMAIL_LOOKUP_VERSION = 1
LOOKUP_CAMPAIGN_LABELS_VERSION = 1
REGULAR_LEVEL_GROUPING_VERSION = 1
SAVED_CAMPAIGN_LOOKUP_PATH = Path(__file__).resolve().parent / "uploads" / "inbound-campaign-databases.zip"


@dataclass(frozen=True)
class EmailCampaignSource:
    filename: str
    sheets: tuple[str, ...]
    campaigns_by_email: dict[str, tuple[str, ...]]
    invalid_email_rows: int
    missing_campaign_rows: int
    campaigns_by_phone: dict[str, tuple[str, ...]]
    invalid_phone_rows: int
    campaign_label: str = ""
    campaigns: tuple[str, ...] = ()


@dataclass(frozen=True)
class EmailCampaignLookup:
    signature: tuple
    sources: tuple[EmailCampaignSource, ...]
    campaigns_by_email: dict[str, tuple[str, ...]]
    campaigns_by_phone: dict[str, tuple[str, ...]]

    @property
    def campaigns(self) -> tuple[str, ...]:
        labels = {normalized_campaign(label): label for source in self.sources for label in source.campaigns}
        return tuple(sorted(labels.values(), key=str.casefold))


def _bucket_campaign(filename: str) -> str:
    buckets = set(re.findall(r"(?<![a-z0-9])b([24])(?![a-z0-9])", Path(filename).stem, re.IGNORECASE))
    if len(buckets) > 1:
        raise InboundReplyError(f"{filename}: the lookup filename must identify either B2 or B4.")
    return f"SBC B{next(iter(buckets))}" if buckets else ""


def _positions(row, fixed_campaign: str = "") -> dict | None:
    keys = [header_key(value) for value in row]
    emails = [i for i, key in enumerate(keys) if key in {"debtoremail", "primaryemail"}
              or re.fullmatch(r"email(?:address)?\d*", key)]
    b4_emails = ([i for i, key in enumerate(keys) if key == "email01"]
                 or [i for i, key in enumerate(keys) if key == "email1"])
    if fixed_campaign == "SBC B4":
        emails = b4_emails
    phones = [i for i, key in enumerate(keys) if re.fullmatch(r"(?:phone|mobile)(?:number)?\d*", key)]
    if not (emails or phones):
        return None
    contacts = {"emails": emails, "b4_emails": b4_emails, "phones": phones}
    if fixed_campaign:
        return {**contacts, "category": None, "kind": "fixed"}
    # GEO has both placement and Area: Area determines its campaign.
    for field, aliases in (("bucket", {"bucketcode", "bucket"}), ("area", {"area"}), ("level", {"level", "placement"})):
        matches = [i for i, key in enumerate(keys) if key in aliases]
        if len(matches) == 1:
            return {**contacts, "category": matches[0], "kind": field}
        if matches:
            return None
    return None


def _campaign(value: object, kind: str) -> str:
    text = " ".join(str(value if value is not None else "").split())
    if not text:
        return ""
    if kind == "bucket":
        match = re.fullmatch(r"(?:b)?([24])(?:\.0)?", text, re.IGNORECASE)
        return f"SBC B{match[1]}" if match else ""
    if kind == "level":
        match = re.fullmatch(r"(?:level\s*)?(\d+)(?:\.0)?", text, re.IGNORECASE)
        if not match or int(match[1]) <= 0:
            return ""
        level = int(match[1])
        return f"SBC Level {1 if level in (2, 3, 4, 5) else level}"
    match = re.fullmatch(r"region[\s_-]*(\d+)", text, re.IGNORECASE)
    area = f"Region {int(match[1])}" if match else ("NCR" if text.casefold() == "ncr" else text.title())
    return inbound_campaign_label(f"SBC {area}")


def _read_rows(rows, index: dict, phone_index: dict, counts: Counter, *, fixed_campaign: str = "",
               campaign_labels: dict | None = None) -> bool:
    rows = iter(rows)
    positions = None
    for _ in range(50):
        row = next(rows, None)
        if row is None:
            break
        positions = _positions(row, fixed_campaign)
        if positions is not None:
            break
    if positions is None:
        return False
    for row in rows:
        if not any(value is not None and str(value).strip() for value in row):
            continue
        category = positions["category"]
        campaign = fixed_campaign or _campaign(row[category] if category is not None and category < len(row) else None,
                                               positions["kind"])
        if campaign and campaign_labels is not None:
            campaign_labels.setdefault(normalized_campaign(campaign), campaign)
        email_columns = positions["b4_emails"] if campaign == "SBC B4" else positions["emails"]
        addresses = {address for column in email_columns if column < len(row)
                     if (address := normalized_email(row[column]))}
        phones = {phone for column in positions["phones"] if column < len(row)
                  if (phone := normalized_phone(str(row[column] if row[column] is not None else "")))}
        if email_columns and not addresses:
            counts["invalid_email"] += 1
        if positions["phones"] and not phones:
            counts["invalid_phone"] += 1
        if not addresses and not phones:
            continue
        if not campaign:
            counts["missing_campaign"] += 1
            continue
        for address in addresses:
            index.setdefault(address, {})[normalized_campaign(campaign)] = campaign
        for phone in phones:
            phone_index.setdefault(phone, {})[normalized_campaign(campaign)] = campaign
    return True


def _csv_rows(filename: str, data: bytes):
    encodings = ("utf-16",) if data.startswith((b"\xff\xfe", b"\xfe\xff")) else ("utf-8-sig", "cp1252")
    for encoding in encodings:
        try:
            text = data.decode(encoding)
            break
        except UnicodeError:
            continue
    else:
        raise InboundReplyError(f"{filename}: save the lookup file as CSV UTF-8 and upload it again.")
    delimiter = inbound_csv_delimiter(text)
    return csv.reader(StringIO(text, newline=""), delimiter=delimiter, strict=True)


def read_email_campaign_source(filename: str, data: bytes) -> EmailCampaignSource:
    """Use OVERALL when available; otherwise merge eligible sheets without duplicate memberships."""
    extension = Path(filename).suffix.casefold()
    fixed_campaign = _bucket_campaign(filename)
    index, phone_index, counts, sheets, labels = {}, {}, Counter(), [], {}
    if extension == ".csv":
        try:
            if _read_rows(_csv_rows(filename, data), index, phone_index, counts, fixed_campaign=fixed_campaign,
                          campaign_labels=labels):
                sheets.append("CSV")
        except csv.Error as error:
            raise InboundReplyError(f"{filename}: the lookup CSV contains an incomplete quoted field.") from error
    elif extension == ".xlsx":
        from openpyxl import load_workbook
        try:
            book = load_workbook(BytesIO(data), read_only=True, data_only=True)
        except Exception as error:
            raise InboundReplyError(f"{filename}: the lookup workbook could not be opened.") from error
        try:
            overall = [sheet for sheet in book if header_key(sheet.title) == "overall"]
            fallback_sheets = list(book)
            if fixed_campaign:
                # Bucket exports also contain field-reference and Lite sheets.
                # Prefer contact tables with account headers when present.
                account_sheets = []
                for sheet in book:
                    sheet.reset_dimensions()
                    if any(_positions(row, fixed_campaign) is not None and
                           {"accountkey", "contractnumber", "accountno", "accountnumber"}.intersection(header_key(v) for v in row)
                           for row in sheet.iter_rows(min_row=1, max_row=50, values_only=True)):
                        account_sheets.append(sheet)
                fallback_sheets = account_sheets or fallback_sheets
            for sheet in overall or fallback_sheets:
                sheet.reset_dimensions()
                if _read_rows(sheet.iter_rows(values_only=True), index, phone_index, counts, fixed_campaign=fixed_campaign,
                              campaign_labels=labels):
                    sheets.append(sheet.title)
            # An unrelated OVERALL sheet must not hide eligible data sheets.
            if overall and not sheets:
                for sheet in fallback_sheets:
                    if sheet in overall:
                        continue
                    sheet.reset_dimensions()
                    if _read_rows(sheet.iter_rows(values_only=True), index, phone_index, counts, fixed_campaign=fixed_campaign,
                                  campaign_labels=labels):
                        sheets.append(sheet.title)
        finally:
            book.close()
    else:
        raise InboundReplyError(f"{filename}: upload an .xlsx or .csv lookup file.")
    if not sheets:
        raise InboundReplyError(f"{filename}: expected email or phone columns and Area, LEVEL, placement or Bucket Code, "
                                "or a filename containing B2 or B4.")
    return EmailCampaignSource(filename, tuple(sheets),
                               {email: tuple(sorted(labels.values(), key=str.casefold)) for email, labels in index.items()},
                               counts["invalid_email"], counts["missing_campaign"],
                               {phone: tuple(sorted(labels.values(), key=str.casefold)) for phone, labels in phone_index.items()},
                               counts["invalid_phone"], fixed_campaign, tuple(sorted(labels.values(), key=str.casefold)))


def _resolved_campaigns(index: dict) -> dict[str, tuple[str, ...]]:
    result = {}
    for contact, labels in index.items():
        # The user's B2-else-B4 rule is independent of upload order.
        selected = [label for key, label in labels.items() if key != "sbc b4" or "sbc b2" not in labels]
        result[contact] = tuple(sorted(selected, key=str.casefold))
    return result


def build_email_campaign_lookup(files: list[tuple[str, bytes]]) -> EmailCampaignLookup:
    sources = tuple(read_email_campaign_source(name, data) for name, data in files)
    index, phone_index = {}, {}
    for source in sources:
        for email, campaigns in source.campaigns_by_email.items():
            for campaign in campaigns:
                index.setdefault(email, {})[normalized_campaign(campaign)] = campaign
        for phone, campaigns in source.campaigns_by_phone.items():
            for campaign in campaigns:
                phone_index.setdefault(phone, {})[normalized_campaign(campaign)] = campaign
    signature = (EMAIL_CAMPAIGN_LOOKUP_VERSION, tuple((name, sha256(data).hexdigest()) for name, data in files))
    return EmailCampaignLookup(signature, sources,
                               _resolved_campaigns(index), _resolved_campaigns(phone_index))


def read_saved_campaign_lookup_files(path: Path) -> list[tuple[str, bytes]]:
    """Read the original filenames and contents of the saved lookup files."""
    try:
        with ZipFile(path) as archive:
            files = [(entry.filename.split("/", 1)[-1], archive.read(entry))
                     for entry in archive.infolist() if not entry.is_dir()]
        if not files:
            raise BadZipFile("No saved lookup files")
        return files
    except (BadZipFile, OSError, RuntimeError) as error:
        raise InboundReplyError("Could not read the saved campaign databases. Upload the files again to replace them.") from error


def save_campaign_lookup_files(path: Path, files: list[tuple[str, bytes]]) -> None:
    """Replace the saved set only after all lookup files have been written."""
    temporary = None
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with NamedTemporaryFile(dir=path.parent, delete=False) as output:
            temporary = Path(output.name)
            with ZipFile(output, "w") as archive:
                for index, (name, contents) in enumerate(files):
                    archive.writestr(f"{index}/{name}", contents)
        os.replace(temporary, path)
    except OSError as error:
        raise InboundReplyError("Could not save the campaign databases. Check write access to the app's uploads folder.") from error
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)


def _workbook_email_placements(source: EmailSource) -> dict[str, tuple[str, ...]] | None:
    headers = tuple(header_key(header) for header in source.headers)
    if headers.count("reportingcampaign") != 1 or headers.count("lookupstatus") != 1:
        return None
    campaign_col, status_col = (headers.index(key) for key in ("reportingcampaign", "lookupstatus"))
    index: dict[str, dict[str, str]] = {}
    for row in source.records:
        values = row.source_values
        if max(campaign_col, status_col) >= len(values):
            continue
        if normalized_campaign(str(values[status_col] or "")) not in {"matched", "multiple placements"}:
            continue
        for value in str(values[campaign_col] or "").split(";"):
            campaign = inbound_campaign_label(value.strip())
            if campaign:
                campaigns = index.setdefault(row.email, {})
                if not excluded_inbound_campaign(campaign):
                    campaigns.setdefault(normalized_campaign(campaign), campaign)
    return {email: tuple(campaigns.values()) for email, campaigns in index.items()}


def apply_email_campaign_lookup(source: EmailSource, lookup: EmailCampaignLookup | None) -> EmailSource:
    """Resolve email campaigns from the live lookup or placements carried in the workbook."""
    campaigns_by_email = (_workbook_email_placements(source) if lookup is None else lookup.campaigns_by_email)
    if campaigns_by_email is None:
        return source
    records, unmatched = [], {}
    header = "Lookup Campaign"
    while header in source.headers:
        header += " (lookup)"
    for row in source.records:
        if row.campaign == PL_RECOVERY_CAMPAIGN:
            campaigns = (PL_RECOVERY_CAMPAIGN,)
        else:
            campaigns = campaigns_by_email.get(row.email)
            if lookup is None and campaigns == ():
                continue  # The workbook assigned only excluded campaigns.
        if not campaigns:
            unmatched.setdefault(row.email, row)
            campaigns = (SBLEGAL_CAMPAIGN,)
        values = row.source_values[:len(source.headers)] + (None,) * max(0, len(source.headers) - len(row.source_values))
        for campaign in campaigns:
            records.append(replace(row, campaign=campaign, source_values=(*values, campaign)))
    return replace(source, headers=(*source.headers, header), records=tuple(records),
                   lookup_applied=True, unmatched_records=tuple(unmatched.values()), workbook_placements=lookup is None)


def apply_phone_campaign_lookup(source: ReplySource, lookup: EmailCampaignLookup) -> ReplySource:
    """Resolve campaigns for non-PL RECOVERY SMS and Viber rows."""
    records, unmatched = [], {}
    excluded = source.excluded_rows
    header = "Lookup Campaign"
    while header in source.headers:
        header += " (lookup)"
    unique_clients = source.layout == "client"
    for row in source.records:
        phone = normalized_phone(row.phone)
        if unique_clients:
            row = _resolve_client_placement(row, phone, lookup)
            if row.placement_issue:
                unmatched.setdefault(client_account_key(row), row)
            campaigns = (SBLEGAL_CAMPAIGN if row.placement_issue else row.campaign,)
        else:
            if not phone:
                excluded += 1
                continue
            campaigns = ((PL_RECOVERY_CAMPAIGN,) if row.campaign == PL_RECOVERY_CAMPAIGN
                         else lookup.campaigns_by_phone.get(phone))
            if not campaigns:
                unmatched.setdefault(phone, row)
                campaigns = (SBLEGAL_CAMPAIGN,)
        values = row.source_values[:len(source.headers)] + (None,) * max(0, len(source.headers) - len(row.source_values))
        for campaign in campaigns:
            records.append(replace(row, campaign=campaign, source_values=(*values, campaign)))
    return replace(source, headers=(*source.headers, header), records=tuple(records),
                   excluded_rows=excluded, lookup_applied=True, unmatched_records=tuple(unmatched.values()))


def _resolve_client_placement(row: ReplyRecord, phone: str, lookup: EmailCampaignLookup) -> ReplyRecord:
    if row.campaign == PL_RECOVERY_CAMPAIGN:
        return row
    matches = {}
    for candidate in row.lookup_phones or (phone,):
        for campaign in lookup.campaigns_by_phone.get(candidate, ()):
            matches.setdefault(campaign, candidate)
    if not matches:
        reason = "Phone not found in placement databases" if row.lookup_phones or phone else "No usable phone number"
        return replace(row, placement_issue=reason)
    if len(matches) > 1:
        original = _client_campaign_key(row.campaign)
        confirmed = {label: number for label, number in matches.items() if _client_campaign_key(label) == original}
        if len(confirmed) != 1:
            return replace(row, placement_issue="Conflicting database placements — " + ", ".join(sorted(matches)))
        matches = confirmed
    campaign, matched_phone = next(iter(matches.items()))
    return replace(row, campaign=campaign, placement_phone=matched_phone)


def _client_campaign_key(label: str) -> str:
    key = normalized_campaign(label).replace("sbc cards ", "sbc ")
    level = re.fullmatch(r"sbc (?:l|level\s+)([1-6])", key)
    return f"sbc level {6 if level[1] == '6' else 1}" if level else key
