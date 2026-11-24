"""Read Email account and agent-history exports for inbound reply matching."""

from __future__ import annotations

import csv
from dataclasses import dataclass
from io import BytesIO, StringIO
from pathlib import Path

from email_validation import validated_email
from inbound_reply import (
    CAMPAIGN_HEADER_ALIASES, CampaignReplyCounts, InboundReplyError, build_reply_workbook,
    excluded_inbound_campaign, inbound_campaign_label, inbound_csv_delimiter, normalized_campaign,
)


EMAIL_REPLY_VERSION = 7
EMAIL_CAMPAIGN_LOOKUP_VERSION = 1


def normalized_email(value: object) -> str:
    address = validated_email(value)
    return address.casefold() if address else ""


def header_key(value: object) -> str:
    return "".join(char for char in str(value or "").casefold() if char.isalnum())


@dataclass(frozen=True)
class EmailRecord:
    email: str
    campaign: str
    sent: bool
    agent: str
    source_values: tuple[object, ...]


@dataclass(frozen=True)
class EmailSource:
    filename: str
    headers: tuple[str, ...]
    records: tuple[EmailRecord, ...]
    excluded_rows: int
    sheet: str
    lookup_applied: bool = False
    unmatched_records: tuple[EmailRecord, ...] = ()
    workbook_placements: bool = False

    @property
    def unmatched_addresses(self) -> int:
        return len(self.unmatched_records)


@dataclass(frozen=True)
class EmailReplyReport:
    dc: EmailSource
    history: EmailSource | None
    accounts: tuple[EmailRecord, ...]
    replied_emails: frozenset[str]
    campaigns: tuple[CampaignReplyCounts, ...]

    @property
    def totals(self) -> CampaignReplyCounts:
        addresses = {row.email for row in self.accounts}
        history_emails = {row.email for row in self.history.records} if self.history is not None else None
        return _reply_counts("All campaigns", addresses, self.replied_emails, history_emails)

    @property
    def history_only_addresses(self) -> int | None:
        if self.history is None:
            return None
        return len({row.email for row in self.history.records} - {row.email for row in self.accounts})


def _positions(row, kind: str, *, require_campaign: bool = True) -> dict[str, int] | None:
    keys = [header_key(value) for value in row]
    # The DC Checking "Email" column is a workflow label, not the sender's address.
    aliases = ({"email": {"sender", "from", "senderemail"}, "campaign": CAMPAIGN_HEADER_ALIASES}
               if kind == "dc" else {"email": {"email", "emailaddress"}, "sent": {"sent"}})
    if kind == "dc" and not require_campaign:
        aliases.pop("campaign")
    result = {}
    for field, names in aliases.items():
        matches = [index for index, key in enumerate(keys) if key in names]
        if len(matches) != 1:
            return None
        result[field] = matches[0]
    for field, names in {"agent": {"person", "agentlark"}, "campaign": CAMPAIGN_HEADER_ALIASES}.items():
        matches = [index for index, key in enumerate(keys) if key in names]
        if len(matches) == 1:
            result[field] = matches[0]
    return result


def _read_rows(rows, filename: str, kind: str, sheet: str, *, require_campaign: bool = True) -> EmailSource:
    rows = iter(rows)
    positions = None
    for _ in range(50):
        row = next(rows, None)
        if row is None:
            break
        positions = _positions(row, kind, require_campaign=require_campaign)
        if positions is not None:
            break
    if positions is None:
        expected = ("sender and Campaign" if require_campaign else "sender") if kind == "dc" else "email and sent"
        raise InboundReplyError(f"{filename}: expected {expected} columns for this Email upload.")
    headers = tuple(str(value) if value is not None else "" for value in row)
    records = []
    excluded = 0
    for row in rows:
        if not any(value is not None and str(value).strip() for value in row):
            continue
        values = {field: row[index] if index < len(row) else None for field, index in positions.items()}
        campaign = str(values.get("campaign") or "").strip() or "Unspecified campaign"
        if excluded_inbound_campaign(campaign):
            continue
        email = normalized_email(values["email"])
        if not email:
            excluded += 1
            continue
        records.append(EmailRecord(
            email, inbound_campaign_label(campaign),
            bool(str(values.get("sent") or "").strip()), str(values.get("agent") or "").strip(), tuple(row),
        ))
    return EmailSource(filename, headers, tuple(records), excluded, sheet)


def read_email_source(filename: str, data: bytes, *, kind: str, require_campaign: bool = True) -> EmailSource:
    if kind not in {"dc", "history"}:
        raise InboundReplyError("Choose DC Checking or Email History for this upload.")
    extension = Path(filename).suffix.casefold()
    if extension == ".csv":
        encodings = ("utf-16",) if data.startswith((b"\xff\xfe", b"\xfe\xff")) else ("utf-8-sig", "cp1252")
        text = None
        for encoding in encodings:
            try:
                text = data.decode(encoding)
                break
            except UnicodeError:
                continue
        if text is None:
            raise InboundReplyError(f"{filename}: save the file as CSV UTF-8 and upload it again.")
        try:
            delimiter = inbound_csv_delimiter(text)
            return _read_rows(csv.reader(StringIO(text, newline=""), delimiter=delimiter, strict=True), filename, kind, "CSV",
                              require_campaign=require_campaign)
        except csv.Error as error:
            raise InboundReplyError(f"{filename}: the CSV contains an incomplete quoted field.") from error
    if extension != ".xlsx":
        raise InboundReplyError(f"{filename}: upload a .xlsx or .csv file.")
    from openpyxl import load_workbook
    try:
        workbook = load_workbook(BytesIO(data), read_only=True, data_only=True)
    except Exception as error:
        raise InboundReplyError(f"{filename}: the Excel workbook could not be opened.") from error
    try:
        candidates, with_campaign = [], []
        for sheet in workbook:
            sheet.reset_dimensions()
            for row in sheet.iter_rows(min_row=1, max_row=50, values_only=True):
                positions = _positions(row, kind, require_campaign=require_campaign)
                if positions is not None:
                    candidates.append(sheet)
                    if "campaign" in positions:
                        with_campaign.append(sheet)
                    break
        # Sender-only review sheets can accompany a complete DC Checking table.
        if kind == "dc" and not require_campaign and with_campaign:
            candidates = with_campaign
        if len(candidates) != 1:
            raise InboundReplyError(f"{filename}: expected one matching data worksheet for this Email upload.")
        sheet = candidates[0]
        return _read_rows(sheet.iter_rows(values_only=True), filename, kind, sheet.title,
                          require_campaign=require_campaign)
    finally:
        workbook.close()


def _reply_counts(label: str, addresses: set[str], replied_emails: frozenset[str],
                  history_emails: set[str] | None) -> CampaignReplyCounts:
    if history_emails is None:
        return CampaignReplyCounts(label, len(addresses), None, None, None)
    replied = len(addresses & replied_emails)
    return CampaignReplyCounts(label, len(addresses), replied, len(addresses) - replied,
                               len(addresses & history_emails))


def match_email_sources(dc: EmailSource, history: EmailSource | None) -> EmailReplyReport:
    """DC contains incoming customer emails; sent history entries identify agent responses."""
    history_rows = history.records if history is not None else ()
    replied_emails = frozenset(row.email for row in history_rows if row.sent)
    history_emails = {row.email for row in history_rows} if history is not None else None
    accounts = {}
    labels = {}
    for row in dc.records:
        key = normalized_campaign(row.campaign)
        labels.setdefault(key, row.campaign)
        accounts.setdefault((key, row.email), row)
    campaigns = []
    for key, label in sorted(labels.items(), key=lambda item: item[1].casefold()):
        addresses = {email for campaign, email in accounts if campaign == key}
        campaigns.append(_reply_counts(label, addresses, replied_emails, history_emails))
    return EmailReplyReport(dc, history, tuple(accounts.values()), replied_emails, tuple(campaigns))


def build_email_unreplied_workbook(report: EmailReplyReport) -> bytes:
    """Retain every source DC field for each unreplied address within a campaign."""
    if report.history is None:
        raise InboundReplyError("Upload Email History to prepare the Email unreplied report.")
    details = []
    headers = report.dc.headers
    for row in sorted(report.accounts, key=lambda row: (normalized_campaign(row.campaign), row.email)):
        if row.email in report.replied_emails:
            continue
        values = row.source_values[:len(headers)] + (None,) * max(0, len(headers) - len(row.source_values))
        details.append((*values, row.email, "Not replied"))
    return build_reply_workbook(report.campaigns, headers, details, matched_header="Matched Email")
