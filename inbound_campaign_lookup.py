"""Resolve inbound campaigns by email or phone from active-account workbooks."""

from __future__ import annotations

import csv
import re
from dataclasses import dataclass, replace
from collections import Counter
from hashlib import sha256
from io import BytesIO, StringIO
from pathlib import Path

from inbound_email import EmailSource, header_key, normalized_email
from inbound_reply import InboundReplyError, ReplySource, inbound_csv_delimiter, normalized_campaign, normalized_phone


EMAIL_CAMPAIGN_LOOKUP_VERSION = 2
PHONE_CAMPAIGN_LOOKUP_VERSION = 1
UNMATCHED_CAMPAIGN = "Unmatched campaign"


@dataclass(frozen=True)
class EmailCampaignSource:
    filename: str
    sheets: tuple[str, ...]
    campaigns_by_email: dict[str, tuple[str, ...]]
    invalid_email_rows: int
    missing_campaign_rows: int
    campaigns_by_phone: dict[str, tuple[str, ...]]
    invalid_phone_rows: int


@dataclass(frozen=True)
class EmailCampaignLookup:
    signature: tuple
    sources: tuple[EmailCampaignSource, ...]
    campaigns_by_email: dict[str, tuple[str, ...]]
    campaigns_by_phone: dict[str, tuple[str, ...]]


def _positions(row) -> dict | None:
    keys = [header_key(value) for value in row]
    emails = [i for i, key in enumerate(keys) if key in {"email", "debtoremail", "emailaddress", "primaryemail"}]
    phones = [i for i, key in enumerate(keys) if re.fullmatch(r"(?:phone|mobile)(?:number)?\d*", key)]
    if len(emails) > 1 or not (emails or phones):
        return None
    # GEO has both placement and Area: Area determines its campaign.
    for field, aliases in (("area", {"area"}), ("level", {"level", "placement"})):
        matches = [i for i, key in enumerate(keys) if key in aliases]
        if len(matches) == 1:
            return {"email": emails[0] if emails else None, "phones": phones, "category": matches[0], "kind": field}
        if matches:
            return None
    return None


def _campaign(value: object, kind: str) -> str:
    text = " ".join(str(value if value is not None else "").split())
    if not text:
        return ""
    if kind == "level":
        match = re.fullmatch(r"(?:level\s*)?(\d+)(?:\.0)?", text, re.IGNORECASE)
        return f"SBC Level {int(match[1])}" if match and int(match[1]) > 0 else ""
    match = re.fullmatch(r"region[\s_-]*(\d+)", text, re.IGNORECASE)
    area = f"Region {int(match[1])}" if match else ("NCR" if text.casefold() == "ncr" else text.title())
    return f"SBC {area}"


def _read_rows(rows, index: dict, phone_index: dict, counts: Counter) -> bool:
    rows = iter(rows)
    positions = None
    for _ in range(50):
        row = next(rows, None)
        if row is None:
            break
        positions = _positions(row)
        if positions is not None:
            break
    if positions is None:
        return False
    for row in rows:
        if not any(value is not None and str(value).strip() for value in row):
            continue
        email_column = positions["email"]
        address = normalized_email(row[email_column] if email_column is not None and email_column < len(row) else None)
        phones = {phone for column in positions["phones"] if column < len(row)
                  if (phone := normalized_phone(str(row[column] if row[column] is not None else "")))}
        if email_column is not None and not address:
            counts["invalid_email"] += 1
        if positions["phones"] and not phones:
            counts["invalid_phone"] += 1
        if not address and not phones:
            continue
        campaign = _campaign(row[positions["category"]] if positions["category"] < len(row) else None,
                             positions["kind"])
        if not campaign:
            counts["missing_campaign"] += 1
            continue
        if address:
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
    index, phone_index, counts, sheets = {}, {}, Counter(), []
    if extension == ".csv":
        try:
            if _read_rows(_csv_rows(filename, data), index, phone_index, counts):
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
            for sheet in overall or book:
                sheet.reset_dimensions()
                if _read_rows(sheet.iter_rows(values_only=True), index, phone_index, counts):
                    sheets.append(sheet.title)
            # An unrelated OVERALL sheet must not hide eligible data sheets.
            if overall and not sheets:
                for sheet in book:
                    if sheet in overall:
                        continue
                    sheet.reset_dimensions()
                    if _read_rows(sheet.iter_rows(values_only=True), index, phone_index, counts):
                        sheets.append(sheet.title)
        finally:
            book.close()
    else:
        raise InboundReplyError(f"{filename}: upload an .xlsx or .csv lookup file.")
    if not sheets:
        raise InboundReplyError(f"{filename}: expected email, debtor_email or phone columns and Area, LEVEL or placement.")
    return EmailCampaignSource(filename, tuple(sheets),
                               {email: tuple(sorted(labels.values(), key=str.casefold)) for email, labels in index.items()},
                               counts["invalid_email"], counts["missing_campaign"],
                               {phone: tuple(sorted(labels.values(), key=str.casefold)) for phone, labels in phone_index.items()},
                               counts["invalid_phone"])


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
                               {email: tuple(sorted(labels.values(), key=str.casefold)) for email, labels in index.items()},
                               {phone: tuple(sorted(labels.values(), key=str.casefold)) for phone, labels in phone_index.items()})


def apply_email_campaign_lookup(source: EmailSource, lookup: EmailCampaignLookup) -> EmailSource:
    """Keep every matching campaign per address and retain unmatched accounts explicitly."""
    records, unmatched = [], set()
    header = "Lookup Campaign"
    while header in source.headers:
        header += " (lookup)"
    for row in source.records:
        campaigns = lookup.campaigns_by_email.get(row.email)
        if not campaigns:
            unmatched.add(row.email)
        values = row.source_values[:len(source.headers)] + (None,) * max(0, len(source.headers) - len(row.source_values))
        for campaign in campaigns or (UNMATCHED_CAMPAIGN,):
            records.append(replace(row, campaign=campaign, source_values=(*values, campaign)))
    return replace(source, headers=(*source.headers, header), records=tuple(records),
                   lookup_applied=True, unmatched_addresses=len(unmatched))


def apply_phone_campaign_lookup(source: ReplySource, lookup: EmailCampaignLookup) -> ReplySource:
    """Resolve SMS/Viber campaigns using all phone columns from the same lookup files."""
    records, unmatched = [], set()
    header = "Lookup Campaign"
    while header in source.headers:
        header += " (lookup)"
    for row in source.records:
        phone = normalized_phone(row.phone)
        campaigns = lookup.campaigns_by_phone.get(phone)
        if phone and not campaigns:
            unmatched.add(phone)
        values = row.source_values[:len(source.headers)] + (None,) * max(0, len(source.headers) - len(row.source_values))
        for campaign in campaigns or (UNMATCHED_CAMPAIGN,):
            records.append(replace(row, campaign=campaign, source_values=(*values, campaign)))
    return replace(source, headers=(*source.headers, header), records=tuple(records),
                   lookup_applied=True, unmatched_addresses=len(unmatched))
