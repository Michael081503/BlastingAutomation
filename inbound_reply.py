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


VIBER_REPLY_VERSION = 6
SBF_CAMPAIGN_EXCLUSION_VERSION = 1
CAMPAIGN_COLUMN_ALIASES_VERSION = 1
CAMPAIGN_HEADER_ALIASES = {"campaign", "campaigns", "bankcampaign", "bankcampaigns", "campaignname"}
VIBER_CAMPAIGN_REPLIED_VERSION = 1
VIBER_AGENT_REPLIED_SOURCE = "campaign_replies"
VIBER_PHONE_MATCHING_VERSION = 1
VIBER_MATCH_MODE = "phone"
VIBER_UNREPLIED_WORKBOOK_VERSION = 1
SOURCE_NAMES = ("message_logs", "campaign_replies")
MATCH_MODES = ("name", "phone")
HEADER_ALIASES = {
    "name": {"clientname", "contactname", "fullname", "name"},
    "phone": {"phone", "phonenumber", "mobile", "mobilenumber"},
    "campaign": CAMPAIGN_HEADER_ALIASES,
    "message_type": {"messagetype"},
}


class InboundReplyError(ValueError):
    """An input problem that the page can explain to the user."""


def normalized_name(value: str) -> str:
    text = unicodedata.normalize("NFKC", value).casefold()
    return " ".join("".join(char if char.isalnum() else " " for char in text).split())


def normalized_campaign(value: str) -> str:
    return " ".join(unicodedata.normalize("NFKC", value).casefold().split())


def excluded_inbound_campaign(value: str) -> bool:
    """Ignore SBF campaign rows, including names separated by underscores or hyphens."""
    return re.search(r"(?<![a-z0-9])sbf(?![a-z0-9])", normalized_campaign(value)) is not None


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
    source_values: tuple[object, ...] = ()


@dataclass(frozen=True)
class ReplySource:
    filename: str
    records: tuple[ReplyRecord, ...]
    excluded_rows: int = 0
    sheet: str = "CSV"
    headers: tuple[str, ...] = ()

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
    excluded_phone_rows: int = 0

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
    return positions if {"phone", "campaign"} <= positions.keys() else None


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
            f"{filename}: expected Phone and Bank Campaign headers. Client Name or Contact Name is optional."
        )
    headers = tuple(str(value) if value is not None else "" for value in row)
    records = []
    excluded = 0
    for row in rows:
        if not any(value is not None and str(value).strip() for value in row):
            continue
        values = {name: str(row[index]).strip() if index < len(row) and row[index] is not None else ""
                  for name, index in positions.items()}
        campaign = values["campaign"] or "Unspecified campaign"
        if excluded_inbound_campaign(campaign):
            continue
        if not normalized_name(values.get("name", "")) and not normalized_phone(values["phone"]):
            excluded += 1
            continue
        records.append(ReplyRecord(
            values.get("name", ""), values["phone"], campaign,
            values.get("message_type", ""), tuple(row),
        ))
    return ReplySource(filename, tuple(records), excluded, sheet, headers)


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
                    f"{filename}: expected one sheet with Phone and Bank Campaign columns. "
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
    """Deduplicate by the selected identity; phone mode never falls back to names."""
    if agent_replied_source not in SOURCE_NAMES or match_mode not in MATCH_MODES:
        raise InboundReplyError("Choose the agent-replied file and a person matching rule.")
    sources = dict(zip(SOURCE_NAMES, (message_logs, campaign_replies)))
    # Retain the legacy name matcher; phone matching requires an actual number.
    alternate_ids = defaultdict(set)
    if match_mode == "name":
        for source in sources.values():
            for row in source.records:
                name, phone = normalized_name(row.name), normalized_phone(row.phone)
                if name and phone:
                    alternate_ids[phone].add(name)
    people = {}
    campaign_labels = {}
    excluded_phone_rows = 0
    for source_name, source in sources.items():
        for row in source.records:
            name, phone = normalized_name(row.name), normalized_phone(row.phone)
            if match_mode == "phone" and not phone:
                excluded_phone_rows += 1
                continue
            primary, alternate = (name, phone) if match_mode == "name" else (phone, name)
            if not primary and len(alternate_ids.get(alternate, ())) == 1:
                primary = next(iter(alternate_ids[alternate]))
            identity = ("primary", primary) if primary else ("fallback", alternate)
            person = people.setdefault(identity, MatchedPerson(row.name or phone))
            if phone:
                person.phones.add(phone)
            campaign_key = normalized_campaign(row.campaign)
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
                            message_logs, campaign_replies, excluded_phone_rows)


def people_rows(report: ViberReplyReport) -> list[dict]:
    """Export one row per matched identity, retaining its name and campaigns."""
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


def unreplied_records(report: ViberReplyReport) -> tuple[ReplyRecord, ...]:
    """Return the first source row per unreplied phone and campaign."""
    if report.match_mode != "phone":
        raise InboundReplyError("The unreplied report requires phone-number matching.")
    pending = {
        (campaign_key, phone)
        for person in report.people
        for campaign_key, group in person.campaigns.items()
        if not group[report.agent_replied_source]
        for phone in person.phones
    }
    source = report.message_logs if report.agent_replied_source == "campaign_replies" else report.campaign_replies
    selected = {}
    for row in source.records:
        campaign_key = normalized_campaign(row.campaign)
        key = (campaign_key, normalized_phone(row.phone))
        if key in pending:
            selected.setdefault(key, row)
    return tuple(row for key, row in sorted(selected.items()))


def build_unreplied_workbook(report: ViberReplyReport) -> bytes:
    """Export campaign counts followed by the corresponding source account details."""
    records = unreplied_records(report)
    source = report.message_logs if report.agent_replied_source == "campaign_replies" else report.campaign_replies
    source_headers = source.headers or ("Name", "Phone", "Bank Campaign", "Message Type")
    details = []
    for record in records:
        values = record.source_values or (record.name, record.phone, record.campaign, record.message_type)
        values = tuple(values[:len(source_headers)]) + (None,) * max(0, len(source_headers) - len(values))
        details.append((*values, normalized_phone(record.phone), "Not replied"))
    return build_reply_workbook(report.campaigns, source_headers, details, matched_header="Matched Phone")


def build_reply_workbook(campaigns: tuple[CampaignReplyCounts, ...], source_headers: tuple[str, ...],
                         details: list[tuple], *, matched_header: str) -> bytes:
    """Use the same summary/detail layout for each channel's unreplied accounts."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    expected = sum(campaign.awaiting_reply for campaign in campaigns)
    if len(details) != expected:
        raise InboundReplyError("Unreplied account details do not match the campaign counts. Upload both files again.")
    workbook = Workbook()
    summary = workbook.active
    summary.title = "Summary"
    summary.append(["Campaign", "Not replied"])
    for campaign in campaigns:
        summary.append([campaign.campaign, campaign.awaiting_reply])
    summary.append(["TOTAL", expected])
    summary.column_dimensions["A"].width = 32
    summary.column_dimensions["B"].width = 20
    summary.auto_filter.ref = f"A1:B{len(campaigns) + 1}"

    sheet = workbook.create_sheet("Unreplied Details")
    sheet.append([*source_headers, matched_header, "Reply Status"])
    for row in details:
        sheet.append(row)
    sheet.auto_filter.ref = sheet.dimensions
    for column, header in enumerate([*source_headers, matched_header, "Reply Status"], 1):
        key = "".join(char for char in header.casefold() if char.isalnum())
        sheet.column_dimensions[get_column_letter(column)].width = (
            75 if key in {"messagecontent", "replycontent"} else
            32 if key in {"clientname", "contactname", "name", "createdby"} else
            26 if key in {"id", "messageid", "bankcampaign"} else 22
        )
        if key in {"phone", "phonenumber", "mobile", "mobilenumber", "matchedphone", "id", "messageid"}:
            for cells in sheet.iter_rows(min_row=2, min_col=column, max_col=column):
                cell = cells[0]
                if cell.value is not None:
                    cell.value = str(cell.value)
                cell.number_format = "@"
    for tab in workbook:
        tab.freeze_panes = "A2"
        for row in tab:
            for cell in row:
                if isinstance(cell.value, str):
                    cell.data_type = "s"  # Source messages and names are literal spreadsheet text.
                cell.alignment = Alignment(vertical="top", wrap_text=True)
        for cell in tab[1]:
            cell.font = Font(bold=True, color="242831")
            cell.fill = PatternFill("solid", fgColor="E2EAFB")
    for cell in summary[summary.max_row]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="E2EAFB")
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()
