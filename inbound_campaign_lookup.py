"""Resolve inbound email campaigns from uploaded active-account workbooks."""

from __future__ import annotations

import csv
import re
from dataclasses import dataclass, replace
from hashlib import sha256
from io import BytesIO, StringIO
from pathlib import Path

from inbound_email import EmailSource, header_key, normalized_email
from inbound_reply import InboundReplyError, inbound_csv_delimiter, normalized_campaign


EMAIL_CAMPAIGN_LOOKUP_VERSION = 1
UNMATCHED_CAMPAIGN = "Unmatched campaign"


@dataclass(frozen=True)
class EmailCampaignSource:
    filename: str
    sheets: tuple[str, ...]
    campaigns_by_email: dict[str, tuple[str, ...]]
    invalid_email_rows: int
    missing_campaign_rows: int


@dataclass(frozen=True)
class EmailCampaignLookup:
    signature: tuple
    sources: tuple[EmailCampaignSource, ...]
    campaigns_by_email: dict[str, tuple[str, ...]]


def _positions(row) -> dict | None:
    keys = [header_key(value) for value in row]
    emails = [i for i, key in enumerate(keys) if key in {"email", "debtoremail", "emailaddress", "primaryemail"}]
    if len(emails) != 1:
        return None
    # GEO has both placement and Area: Area determines its campaign.
    for field, aliases in (("area", {"area"}), ("level", {"level", "placement"})):
        matches = [i for i, key in enumerate(keys) if key in aliases]
        if len(matches) == 1:
            return {"email": emails[0], "category": matches[0], "kind": field}
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


def _read_rows(rows, index: dict, counts: list[int]) -> bool:
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
        address = normalized_email(row[positions["email"]] if positions["email"] < len(row) else None)
        if not address:
            counts[0] += 1
            continue
        campaign = _campaign(row[positions["category"]] if positions["category"] < len(row) else None,
                             positions["kind"])
        if not campaign:
            counts[1] += 1
            continue
        index.setdefault(address, {})[normalized_campaign(campaign)] = campaign
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
    index, counts, sheets = {}, [0, 0], []
    if extension == ".csv":
        try:
            if _read_rows(_csv_rows(filename, data), index, counts):
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
                if _read_rows(sheet.iter_rows(values_only=True), index, counts):
                    sheets.append(sheet.title)
            # An unrelated OVERALL sheet must not hide eligible data sheets.
            if overall and not sheets:
                for sheet in book:
                    if sheet in overall:
                        continue
                    sheet.reset_dimensions()
                    if _read_rows(sheet.iter_rows(values_only=True), index, counts):
                        sheets.append(sheet.title)
        finally:
            book.close()
    else:
        raise InboundReplyError(f"{filename}: upload an .xlsx or .csv lookup file.")
    if not sheets:
        raise InboundReplyError(f"{filename}: expected an email or debtor_email column and Area, LEVEL or placement.")
    return EmailCampaignSource(filename, tuple(sheets),
                               {email: tuple(sorted(labels.values(), key=str.casefold)) for email, labels in index.items()},
                               *counts)


def build_email_campaign_lookup(files: list[tuple[str, bytes]]) -> EmailCampaignLookup:
    sources = tuple(read_email_campaign_source(name, data) for name, data in files)
    index = {}
    for source in sources:
        for email, campaigns in source.campaigns_by_email.items():
            for campaign in campaigns:
                index.setdefault(email, {})[normalized_campaign(campaign)] = campaign
    signature = (EMAIL_CAMPAIGN_LOOKUP_VERSION, tuple((name, sha256(data).hexdigest()) for name, data in files))
    return EmailCampaignLookup(signature, sources,
                               {email: tuple(sorted(labels.values(), key=str.casefold)) for email, labels in index.items()})


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
