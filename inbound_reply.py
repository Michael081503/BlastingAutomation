"""Read Viber exports and match people across the two reply files."""

from __future__ import annotations

import csv
import re
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from io import BytesIO, StringIO
from pathlib import Path
from typing import Iterable


VIBER_REPLY_VERSION = 1
SOURCE_NAMES = ("message_logs", "campaign_replies")
MATCH_MODES = ("name", "phone")
HEADER_ALIASES = {
    "name": {"clientname", "contactname", "fullname", "name"},
    "phone": {"phone", "phonenumber", "mobile", "mobilenumber"},
    "campaign": {"bankcampaign", "campaign"},
    "message_type": {"messagetype"},
}


class InboundReplyError(ValueError):
    """An input problem that the page can explain to the user."""


def normalized_name(value: str) -> str:
    text = unicodedata.normalize("NFKC", value).casefold()
    return " ".join("".join(char if char.isalnum() else " " for char in text).split())


def normalized_phone(value: str) -> str:
    value = value.strip()
    if re.fullmatch(r"\d+\.0+", value):
        value = value.split(".")[0]
    if not re.fullmatch(r"[+\d\s().\-]+", value):
        return ""
    digits = re.sub(r"\D", "", value)
    if len(digits) == 12 and digits.startswith("63"):
        digits = "0" + digits[2:]
    elif len(digits) == 10 and digits.startswith("9"):
        digits = "0" + digits
    return digits if 7 <= len(digits) <= 15 else ""


@dataclass(frozen=True)
class ReplyRecord:
    name: str
    phone: str
    campaign: str
    message_type: str = ""


@dataclass(frozen=True)
class ReplySource:
    filename: str
    records: tuple[ReplyRecord, ...]
    excluded_rows: int = 0
    sheet: str = "CSV"

    @property
    def message_types(self) -> dict[str, int]:
        return dict(Counter(row.message_type or "Not specified" for row in self.records))


@dataclass
class MatchedPerson:
    name: str
    phones: set[str] = field(default_factory=set)
    # Campaign -> original label and row counts for each file. Retain all campaigns.
    campaigns: dict[str, dict] = field(default_factory=dict)

    def source_rows(self, source: str, campaign: str | None = None) -> int:
        groups = [self.campaigns[campaign]] if campaign is not None else self.campaigns.values()
        return sum(group[source] for group in groups)


@dataclass(frozen=True)
class CampaignReplyCounts:
    campaign: str
    unique_people: int
    agent_replied: int
    awaiting_reply: int
    matched_both_files: int

    @property
    def agent_percent(self) -> float | None:
        return self.agent_replied / self.unique_people if self.unique_people else None


@dataclass(frozen=True)
class ViberReplyReport:
    people: tuple[MatchedPerson, ...]
    campaigns: tuple[CampaignReplyCounts, ...]
    agent_replied_source: str
    match_mode: str
    message_logs: ReplySource
    campaign_replies: ReplySource

    @property
    def totals(self) -> CampaignReplyCounts:
        return _counts("All campaigns", self.people, self.agent_replied_source)


def _header_positions(row: Iterable) -> dict[str, int] | None:
    keys = ["".join(char for char in str(cell or "").casefold() if char.isalnum()) for cell in row]
    positions = {}
    for field_name, aliases in HEADER_ALIASES.items():
        matches = [index for index, key in enumerate(keys) if key in aliases]
        if len(matches) > 1:
            return None
        if matches:
            positions[field_name] = matches[0]
    return positions if {"name", "phone", "campaign"} <= positions.keys() else None


def _read_rows(rows: Iterable, filename: str, sheet: str) -> ReplySource:
    rows = iter(rows)
    positions = None
    for _ in range(50):
        row = next(rows, None)
        if row is None:
            break
        positions = _header_positions(row)
        if positions is not None:
            break
    if positions is None:
        raise InboundReplyError(
            f"{filename}: expected Client Name or Contact Name, Phone, and Bank Campaign headers."
        )
    records = []
    excluded = 0
    for row in rows:
        if not any(value is not None and str(value).strip() for value in row):
            continue
        values = {name: str(row[index]).strip() if index < len(row) and row[index] is not None else ""
                  for name, index in positions.items()}
        if not normalized_name(values["name"]) and not normalized_phone(values["phone"]):
            excluded += 1
            continue
        records.append(ReplyRecord(
            values["name"], values["phone"], values["campaign"] or "Unspecified campaign",
            values.get("message_type", ""),
        ))
    return ReplySource(filename, tuple(records), excluded, sheet)


def read_reply_source(filename: str, data: bytes) -> ReplySource:
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
            try:
                delimiter = csv.Sniffer().sniff(text[:16384], delimiters=",;\t|").delimiter
            except csv.Error:
                delimiter = ","
            return _read_rows(csv.reader(StringIO(text, newline=""), delimiter=delimiter, strict=True), filename, "CSV")
        except csv.Error as error:
            raise InboundReplyError(f"{filename}: the CSV has an incomplete or invalid quoted field.") from error
    if extension == ".xlsx":
        from openpyxl import load_workbook
        try:
            workbook = load_workbook(BytesIO(data), read_only=True, data_only=True)
        except Exception as error:
            raise InboundReplyError(f"{filename}: this Excel workbook could not be opened.") from error
        try:
            candidates = []
            for sheet in workbook:
                sheet.reset_dimensions()
                for row in sheet.iter_rows(min_row=1, max_row=50, values_only=True):
                    if _header_positions(row) is not None:
                        candidates.append(sheet)
                        break
            if len(candidates) != 1:
                raise InboundReplyError(
                    f"{filename}: expected one sheet with Name, Phone and Bank Campaign columns. "
                    "Save the required worksheet as CSV and upload that file."
                )
            sheet = candidates[0]
            return _read_rows(sheet.iter_rows(values_only=True), filename, sheet.title)
        finally:
            workbook.close()
    raise InboundReplyError(f"{filename}: upload a .csv or .xlsx file.")


def _counts(label: str, people: Iterable[MatchedPerson], replied_source: str,
            campaign: str | None = None) -> CampaignReplyCounts:
    people = tuple(people)
    replied = sum(person.source_rows(replied_source, campaign) > 0 for person in people)
    matched = sum(all(person.source_rows(source, campaign) > 0 for source in SOURCE_NAMES) for person in people)
    return CampaignReplyCounts(label, len(people), replied, len(people) - replied, matched)


def match_viber_sources(message_logs: ReplySource, campaign_replies: ReplySource, *,
                        agent_replied_source: str, match_mode: str) -> ViberReplyReport:
    """Count one person once; an agent reply takes priority over a pending row."""
    if agent_replied_source not in SOURCE_NAMES or match_mode not in MATCH_MODES:
        raise InboundReplyError("Choose the agent-replied file and a person matching rule.")
    sources = dict(zip(SOURCE_NAMES, (message_logs, campaign_replies)))
    # Only use the alternate field for a missing identity when it is unambiguous.
    alternate_ids = defaultdict(set)
    for source in sources.values():
        for row in source.records:
            name, phone = normalized_name(row.name), normalized_phone(row.phone)
            primary, alternate = (name, phone) if match_mode == "name" else (phone, name)
            if primary and alternate:
                alternate_ids[alternate].add(primary)
    people = {}
    campaign_labels = {}
    for source_name, source in sources.items():
        for row in source.records:
            name, phone = normalized_name(row.name), normalized_phone(row.phone)
            primary, alternate = (name, phone) if match_mode == "name" else (phone, name)
            if not primary and len(alternate_ids.get(alternate, ())) == 1:
                primary = next(iter(alternate_ids[alternate]))
            identity = ("primary", primary) if primary else ("fallback", alternate)
            person = people.setdefault(identity, MatchedPerson(row.name or phone))
            if phone:
                person.phones.add(phone)
            campaign_key = " ".join(unicodedata.normalize("NFKC", row.campaign).casefold().split())
            campaign_labels.setdefault(campaign_key, row.campaign)
            group = person.campaigns.setdefault(campaign_key, {
                "label": campaign_labels[campaign_key], "message_logs": 0, "campaign_replies": 0,
            })
            group[source_name] += 1
    campaigns = tuple(
        _counts(label, (person for person in people.values() if key in person.campaigns), agent_replied_source, key)
        for key, label in sorted(campaign_labels.items(), key=lambda item: item[1].casefold())
    )
    return ViberReplyReport(tuple(people.values()), campaigns, agent_replied_source, match_mode,
                            message_logs, campaign_replies)


def people_rows(report: ViberReplyReport) -> list[dict]:
    """A person appears once globally, retaining every campaign and phone number."""
    rows = []
    for person in report.people:
        replied = [group["label"] for group in person.campaigns.values() if group[report.agent_replied_source]]
        pending = [group["label"] for group in person.campaigns.values() if not group[report.agent_replied_source]]
        rows.append({
            "Person": person.name,
            "Phones": "; ".join(sorted(person.phones)),
            "Campaigns": "; ".join(group["label"] for group in person.campaigns.values()),
            "Agent replied": "YES" if replied else "NO",
            "Replied campaigns": "; ".join(replied),
            "Awaiting reply campaigns": "; ".join(pending),
            "Message log rows": person.source_rows("message_logs"),
            "Campaign reply rows": person.source_rows("campaign_replies"),
        })
    return rows


def build_people_csv(report: ViberReplyReport) -> bytes:
    rows = people_rows(report)
    headers = ("Person", "Phones", "Campaigns", "Agent replied", "Replied campaigns",
               "Awaiting reply campaigns", "Message log rows", "Campaign reply rows")
    output = StringIO(newline="")
    writer = csv.DictWriter(output, fieldnames=headers, lineterminator="\r\n")
    writer.writeheader()
    writer.writerows({**row, "Phones": "\t" + row["Phones"] if row["Phones"] else ""} for row in rows)
    return output.getvalue().encode("utf-8-sig")
