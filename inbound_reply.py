"""Read Viber exports and match people across the two reply files."""

from __future__ import annotations

import csv
import re
import unicodedata
from collections import Counter, defaultdict
from copy import copy
from dataclasses import dataclass, field, replace
from datetime import date, datetime, timedelta, timezone
from io import BytesIO, StringIO
from pathlib import Path
from textwrap import TextWrapper
from typing import Iterable


VIBER_REPLY_VERSION = 10
SMS_CLIENT_TOUCH_VERSION = 3
CUSTOMER_REPLY_POPULATION_VERSION = 1
SBF_CAMPAIGN_EXCLUSION_VERSION = 1
CAMPAIGN_COLUMN_ALIASES_VERSION = 1
INBOUND_CSV_DELIMITER_VERSION = 1
CAMPAIGN_HEADER_ALIASES = {"campaign", "campaigns", "bankcampaign", "bankcampaigns", "campaignname"}
VIBER_CAMPAIGN_REPLIED_VERSION = 1
VIBER_AGENT_REPLIED_SOURCE = "campaign_replies"
VIBER_PHONE_MATCHING_VERSION = 1
VIBER_MATCH_MODE = "phone"
VIBER_UNREPLIED_WORKBOOK_VERSION = 3
INBOUND_MESSAGE_COLUMN_WIDTH = 96
EXCEL_MAX_ROW_HEIGHT = 409
INBOUND_ACCOUNT_ROW_HEIGHT = 63
EXCEL_CELL_TEXT_LIMIT = 32767
EXCEL_CELL_LINE_FEED_LIMIT = 253
SOURCE_NAMES = ("message_logs", "campaign_replies")
MATCH_MODES = ("name", "phone")
MANILA_TIMEZONE = timezone(timedelta(hours=8))
MINDANAO_CAMPAIGN = "SBC Mindanao"
PL_RECOVERY_CAMPAIGN = "PL RECOVERY"
SBLEGAL_CAMPAIGN = "sblegal"
MINDANAO_CAMPAIGN_ALIASES = frozenset({
    "mindanao", "sbcmindanao", "sbcgeomindanao",
    "region3", "r3", "sbcregion3", "sbcr3", "sbcgeoregion3", "sbcgeor3",
})
HEADER_ALIASES = {
    "name": {"clientname", "contactname", "fullname", "name"},
    "phone": {"phone", "phonenumber", "mobile", "mobilenumber"},
    "campaign": CAMPAIGN_HEADER_ALIASES,
    "message_type": {"messagetype"},
}


class InboundReplyError(ValueError):
    """An input problem that the page can explain to the user."""


def inbound_csv_delimiter(text: str) -> str:
    # Multiline message cells can prevent sniffing the full sample. The header
    # still identifies semicolon/tab exports without treating message text as CSV.
    for sample in (text[:16384], text.splitlines()[0] if text else ""):
        try:
            return csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
        except csv.Error:
            continue
    return ","


def normalized_name(value: str) -> str:
    text = unicodedata.normalize("NFKC", value).casefold()
    return " ".join("".join(char if char.isalnum() else " " for char in text).split())


def normalized_campaign(value: str) -> str:
    return " ".join(unicodedata.normalize("NFKC", value).casefold().split())


def inbound_campaign_label(value: str) -> str:
    """Normalize source labels for shared reporting campaigns."""
    name = normalized_name(value)
    if name in {SBLEGAL_CAMPAIGN, "unspecified campaign", "unmatched campaign"}:
        return SBLEGAL_CAMPAIGN
    if " pl recovery " in f" {name} ":
        return PL_RECOVERY_CAMPAIGN
    if name.replace(" ", "") in MINDANAO_CAMPAIGN_ALIASES:
        return MINDANAO_CAMPAIGN
    return value


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


def reply_day(value: object, *, day_first: bool = False) -> date | None:
    """Read the Manila calendar day of a message from an export."""
    if isinstance(value, datetime):
        return value.astimezone(MANILA_TIMEZONE).date() if value.tzinfo is not None else value.date()
    if isinstance(value, date):
        return value
    if not isinstance(value, str) or not value.strip():
        return None
    text = value.strip()
    try:
        return reply_day(datetime.fromisoformat(text))
    except ValueError:
        pass
    # Viber adds a comma before the time in its day-first timestamp.
    text = re.split(r"[\s,T]", text, maxsplit=1)[0]
    orders = ("%d{sep}%m{sep}%Y", "%m{sep}%d{sep}%Y")
    for separator in ("/", "-"):
        if separator in text:
            for pattern in ("%Y{sep}%m{sep}%d", *(orders if day_first else reversed(orders))):
                try:
                    return datetime.strptime(text, pattern.format(sep=separator)).date()
                except ValueError:
                    continue
    return None


def reply_date_column(keys: list[str], *preferred: str) -> int | None:
    for key in (*preferred, "timestamp", "datetime", "date"):
        if key in keys:
            return keys.index(key)
    return None


@dataclass(frozen=True)
class ReplyRecord:
    name: str
    phone: str
    campaign: str
    message_type: str = ""
    source_values: tuple[object, ...] = ()
    account: str = ""
    lookup_phones: tuple[str, ...] = ()
    placement_phone: str = ""
    placement_issue: str = ""
    reply_days: frozenset[date] = frozenset()


@dataclass(frozen=True)
class ReplySource:
    filename: str
    records: tuple[ReplyRecord, ...]
    excluded_rows: int = 0
    sheet: str = "CSV"
    headers: tuple[str, ...] = ()
    lookup_applied: bool = False
    unmatched_records: tuple[ReplyRecord, ...] = ()
    layout: str = ""
    duplicate_rows: int = 0
    has_reply_dates: bool = False

    @property
    def unmatched_addresses(self) -> int:
        return len(self.unmatched_records)

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

    def is_replied(self, source: str, campaign: str) -> bool:
        group = self.campaigns[campaign]
        return bool(group.get("agent_replied", group[source]))


@dataclass(frozen=True)
class CampaignReplyCounts:
    campaign: str
    unique_people: int
    agent_replied: int | None
    awaiting_reply: int | None
    matched_both_files: int | None

    @property
    def agent_percent(self) -> float | None:
        return self.agent_replied / self.unique_people if self.unique_people and self.agent_replied is not None else None

    @property
    def agent_worked_percent(self) -> float | None:
        return self.awaiting_reply / self.unique_people if self.unique_people and self.awaiting_reply is not None else None


@dataclass(frozen=True)
class ViberReplyReport:
    people: tuple[MatchedPerson, ...]
    campaigns: tuple[CampaignReplyCounts, ...]
    agent_replied_source: str
    match_mode: str
    message_logs: ReplySource
    campaign_replies: ReplySource
    excluded_phone_rows: int = 0
    client_touches: tuple[tuple[ReplyRecord, bool], ...] | None = None

    @property
    def totals(self) -> CampaignReplyCounts:
        return _counts("All campaigns", self.people, self.agent_replied_source)


def _header_positions(row: Iterable, *, require_campaign: bool = True) -> tuple[dict[str, int], str] | None:
    keys = ["".join(char for char in str(cell or "").casefold() if char.isalnum()) for cell in row]
    labels = set(keys)
    layout = ""
    aliases = HEADER_ALIASES
    if {"receiveddatetime", "accountno", "client", "phonenumber"} <= labels:
        layout = "client"
        aliases = {"name": {"debtorname"}, "phone": {"phonenumber"}, "campaign": {"client"}, "account": {"accountno"}}
    elif {"calltype", "account", "client", "contactnumber"} <= labels:
        layout = "agent_touch"
        aliases = {"name": {"customer"}, "phone": {"contactnumber"}, "campaign": {"client"}, "account": {"account"}}
    positions = {}
    for field_name, names in aliases.items():
        matches = [index for index, key in enumerate(keys) if key in names]
        if len(matches) > 1:
            return None
        if matches:
            positions[field_name] = matches[0]
    if not layout:
        for key in ("accountno", "accountnumber", "account"):
            if key in keys:
                positions["account"] = keys.index(key)
                break
    date_column = reply_date_column(keys, "receiveddatetime", "receivedat", "receiveddate",
                                    "sentdatetime", "sentat", "sentdate")
    if date_column is not None:
        positions["reply_date"] = date_column
    required = {"phone", "campaign"} if require_campaign else {"phone"}
    return (positions, layout) if required <= positions.keys() else None


def _read_rows(rows: Iterable, filename: str, sheet: str, *, require_campaign: bool = True) -> ReplySource:
    rows = iter(rows)
    header = None
    for _ in range(50):
        row = next(rows, None)
        if row is None:
            break
        header = _header_positions(row, require_campaign=require_campaign)
        if header is not None:
            break
    if header is None:
        expected = "Phone and Bank Campaign" if require_campaign else "Phone"
        raise InboundReplyError(f"{filename}: expected {expected} headers. Client Name or Contact Name is optional.")
    positions, layout = header
    headers = tuple(str(value) if value is not None else "" for value in row)
    records = []
    excluded = 0
    for row in rows:
        if not any(value is not None and str(value).strip() for value in row):
            continue
        values = {name: str(row[index]).strip() if index < len(row) and row[index] is not None else ""
                  for name, index in positions.items()}
        campaign = values.get("campaign") or SBLEGAL_CAMPAIGN
        if excluded_inbound_campaign(campaign):
            continue
        if not values.get("account") and not normalized_name(values.get("name", "")) and not normalized_phone(values["phone"]):
            excluded += 1
            continue
        date_index = positions.get("reply_date", len(row))
        day = reply_day(row[date_index], day_first=True) if date_index < len(row) else None
        records.append(ReplyRecord(
            values.get("name", ""), values["phone"], inbound_campaign_label(campaign),
            values.get("message_type", ""), tuple(row), values.get("account", ""),
            reply_days=frozenset((day,)) if day is not None else frozenset(),
        ))
    source = ReplySource(filename, tuple(records), excluded, sheet, headers, layout=layout,
                         has_reply_dates="reply_date" in positions)
    return deduplicate_client_accounts(source) if layout == "client" else source


def read_reply_source(filename: str, data: bytes, *, require_campaign: bool = True) -> ReplySource:
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
            return _read_rows(csv.reader(StringIO(text, newline=""), delimiter=delimiter, strict=True), filename, "CSV",
                              require_campaign=require_campaign)
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
                    header = _header_positions(row, require_campaign=require_campaign)
                    if header is not None:
                        positions, layout = header
                        candidates.append((sheet, layout, "campaign" in positions))
                        break
            candidates = ([item for item in candidates if item[1]]
                          or [item for item in candidates if item[2]] or candidates)
            if len(candidates) != 1:
                expected = "Phone and Bank Campaign" if require_campaign else "Phone"
                raise InboundReplyError(
                    f"{filename}: expected one sheet with {expected} columns. "
                    "Save the required worksheet as CSV and upload that file."
                )
            sheet = candidates[0][0]
            return _read_rows(sheet.iter_rows(values_only=True), filename, sheet.title,
                              require_campaign=require_campaign)
        finally:
            workbook.close()
    raise InboundReplyError(f"{filename}: upload a .csv or .xlsx file.")


def _counts(label: str, people: Iterable[MatchedPerson], replied_source: str,
            campaign: str | None = None) -> CampaignReplyCounts:
    people = tuple(people)
    replied = sum(any(person.is_replied(replied_source, key) for key in person.campaigns
                      if campaign is None or key == campaign) for person in people)
    matched = sum(all(person.source_rows(source, campaign) > 0 for source in SOURCE_NAMES) for person in people)
    return CampaignReplyCounts(label, len(people), replied, len(people) - replied, matched)


def client_account_key(row: ReplyRecord) -> tuple[str, str]:
    return ("account", row.account.casefold()) if row.account else ("phone", normalized_phone(row.phone))


def deduplicate_client_accounts(source: ReplySource) -> ReplySource:
    """Keep the first client row and retain each account's phones for placement lookup."""
    identified_phones = {normalized_phone(row.phone) for row in source.records if row.account}
    unique = {}
    phones = defaultdict(dict)
    skipped = duplicates = 0
    for row in source.records:
        key = client_account_key(row)
        if not key[1]:
            skipped += 1
        elif key[0] == "phone" and key[1] in identified_phones:
            duplicates += 1
        else:
            if key in unique:
                duplicates += 1
                first = unique[key]
                campaign = PL_RECOVERY_CAMPAIGN if row.campaign == PL_RECOVERY_CAMPAIGN else first.campaign
                unique[key] = replace(first, campaign=campaign, reply_days=first.reply_days | row.reply_days)
            else:
                unique[key] = row
            for phone in row.lookup_phones or (normalized_phone(row.phone),):
                if phone:
                    phones[key][phone] = None
    records = tuple(replace(row, lookup_phones=tuple(phones[key])) for key, row in unique.items())
    return replace(source, records=records, layout="client",
                   excluded_rows=source.excluded_rows + skipped, duplicate_rows=source.duplicate_rows + duplicates)


def _client_activity(source: ReplySource) -> dict[tuple[str, str], set[date]]:
    activity = defaultdict(set)
    for row in source.records:
        if row.account:
            activity["account", row.account.casefold()].update(row.reply_days)
        phone = normalized_phone(row.phone)
        if phone:
            activity["phone", phone].update(row.reply_days)
            if not row.account:
                activity["phone without account", phone].update(row.reply_days)
    return activity


def match_client_touches(clients: ReplySource, touches: ReplySource, replies: ReplySource) -> ViberReplyReport:
    """Match outgoing calls and SMS replies to the deduplicated Client list."""
    touch_activity = _client_activity(touches)
    reply_activity = _client_activity(replies)
    same_day = clients.has_reply_dates and replies.has_reply_dates
    client_phones = Counter(phone for row in clients.records
                            for phone in row.lookup_phones or (normalized_phone(row.phone),))
    people, outcomes, groups = [], [], defaultdict(list)
    for row in clients.records:
        phone = normalized_phone(row.phone)
        phone_key = "phone without account" if row.account else "phone"
        keys = [(phone_key, candidate) for candidate in row.lookup_phones or (phone,)
                if candidate and client_phones[candidate] == 1]
        if row.account:
            keys.append(("account", row.account.casefold()))
        touched = any(key in touch_activity for key in keys)
        matched = any(key in reply_activity for key in keys)
        replied = (any(row.reply_days.intersection(reply_activity.get(key, ())) for key in keys)
                   if same_day else matched)
        key = normalized_campaign(row.campaign)
        person = MatchedPerson(row.name or row.account or phone, {phone} if phone else set(), {
            key: {"label": row.campaign, "message_logs": 1, "campaign_replies": int(matched), "agent_replied": replied},
        })
        people.append(person)
        groups[key].append(person)
        outcomes.append((row, touched))
    campaigns = tuple(_counts(group[0].campaigns[key]["label"], group, "campaign_replies", key)
                      for key, group in sorted(groups.items()))
    return ViberReplyReport(tuple(people), campaigns, "campaign_replies", "phone",
                            clients, touches, client_touches=tuple(outcomes))


def match_viber_sources(message_logs: ReplySource, campaign_replies: ReplySource, *,
                        agent_replied_source: str, match_mode: str) -> ViberReplyReport:
    """Phone reports count customer replies, matching agent responses within each campaign."""
    if agent_replied_source not in SOURCE_NAMES or match_mode not in MATCH_MODES:
        raise InboundReplyError("Choose the agent-replied file and a person matching rule.")
    sources = dict(zip(SOURCE_NAMES, (message_logs, campaign_replies)))
    same_day = match_mode == "phone" and all(source.has_reply_dates for source in sources.values())
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
    days = {source: defaultdict(set) for source in SOURCE_NAMES}
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
            if same_day and row.reply_days:
                days[source_name][identity, campaign_key].update(row.reply_days)
    if match_mode == "phone":
        customer_source = next(source for source in SOURCE_NAMES if source != agent_replied_source)
        customers = {}
        for identity, person in people.items():
            person.campaigns = {key: group for key, group in person.campaigns.items() if group[customer_source]}
            if same_day:
                for key, group in person.campaigns.items():
                    match_key = (identity, key)
                    group["agent_replied"] = not days[customer_source].get(match_key, set()).isdisjoint(
                        days[agent_replied_source].get(match_key, ()))
            if person.campaigns:
                customers[identity] = person
        people = customers
        included = {key for person in people.values() for key in person.campaigns}
        campaign_labels = {key: label for key, label in campaign_labels.items() if key in included}
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
        replied, pending = [], []
        for key, group in person.campaigns.items():
            (replied if person.is_replied(report.agent_replied_source, key) else pending).append(group["label"])
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
    if report.client_touches is not None:
        return tuple(row for (row, _), person in zip(report.client_touches, report.people)
                     if not person.is_replied(report.agent_replied_source, normalized_campaign(row.campaign)))
    if report.match_mode != "phone":
        raise InboundReplyError("The unreplied report requires phone-number matching.")
    pending = {
        (campaign_key, phone)
        for person in report.people
        for campaign_key in person.campaigns
        if not person.is_replied(report.agent_replied_source, campaign_key)
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


def append_report_detail_rows(sheet, rows) -> dict[tuple[int, int], str]:
    """Write source rows and save text that exceeds Excel's per-cell limits."""
    full_text = {}
    for row_number, values in enumerate(rows, start=sheet.max_row + 1):
        display = []
        for column, value in enumerate(values, 1):
            if isinstance(value, str) and (
                value.count("\n") + value.count("\r") - value.count("\r\n") > EXCEL_CELL_LINE_FEED_LIMIT
                or (len(value) > EXCEL_CELL_TEXT_LIMIT // 2
                    and len(value.encode("utf-16-le")) // 2 > EXCEL_CELL_TEXT_LIMIT)
            ):
                full_text[row_number, column] = value
                value = value[:240] + " …"
            display.append(value)
        sheet.append(display)
    return full_text


def _report_text_parts(text: str) -> list[tuple[str, int]]:
    """Split text into short blocks that fit the wrapped continuation rows."""
    parts = []
    start = 0
    line_count = 0
    for segment in re.finditer(r"[^\r\n]{1,70}(?:\r\n|\r|\n)?|\r\n|\r|\n", text):
        line_count += 1
        if line_count == 12:
            parts.append((text[start:segment.end()], line_count))
            start = segment.end()
            line_count = 0
    if start < len(text):
        parts.append((text[start:], line_count))
    return parts


def _add_report_text_continuations(sheet, text_sheet, full_text: dict[tuple[int, int], str]) -> None:
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils.cell import quote_sheetname
    from openpyxl.worksheet.hyperlink import Hyperlink

    wrap = Alignment(vertical="top", wrap_text=True)
    highlight = PatternFill("solid", fgColor="E2EAFB")
    source_location = quote_sheetname(sheet.title)
    text_location = quote_sheetname(text_sheet.title)
    section_row = sheet.max_row + 2 if text_sheet is sheet else 1
    text_sheet.merge_cells(start_row=section_row, start_column=1, end_row=section_row, end_column=5)
    title = text_sheet.cell(section_row, 1, "Full text for long entries")
    title.font = Font(bold=True)
    title.fill = highlight
    title.alignment = wrap
    row = section_row + 1
    for (source_row, source_column), text in sorted(full_text.items()):
        source = sheet.cell(source_row, source_column)
        parts = _report_text_parts(text)
        source.fill = highlight
        source.hyperlink = Hyperlink(ref=source.coordinate, location=f"{text_location}!A{row}", tooltip="Read full text")
        for index, (part, line_count) in enumerate(parts, start=1):
            link = text_sheet.cell(row, 1, f"{source.coordinate} ({index}/{len(parts)})")
            link.style = "Hyperlink"
            link.alignment = wrap
            link.hyperlink = Hyperlink(ref=link.coordinate, location=f"{source_location}!{source.coordinate}",
                                       tooltip="Back to the account cell")
            text_sheet.merge_cells(start_row=row, start_column=2, end_row=row, end_column=5)
            message = text_sheet.cell(row, 2, part)
            message.data_type = "s"
            message.alignment = wrap
            text_sheet.row_dimensions[row].height = line_count * 30 + 18
            row += 1


def fit_report_text(workbook, detail_text: dict[str, dict[tuple[int, int], str]], *, separate_text: bool = False) -> None:
    """Wrap text and keep account rows compact, with links to longer entries."""
    from openpyxl.utils import column_index_from_string

    for sheet in workbook:
        full_text = detail_text.get(sheet.title)
        height_limit = INBOUND_ACCOUNT_ROW_HEIGHT if full_text is not None else EXCEL_MAX_ROW_HEIGHT
        widths = dict.fromkeys(range(1, sheet.max_column + 1), sheet.sheet_format.defaultColWidth or 13)
        for dimension in sheet.column_dimensions.values():
            first_column = dimension.min or column_index_from_string(dimension.index)
            for column in range(first_column, (dimension.max or first_column) + 1):
                widths[column] = dimension.width
        merged_widths = {
            (area.min_row, area.min_col): sum(widths[col] for col in range(area.min_col, area.max_col + 1))
            for area in sheet.merged_cells.ranges
        }
        wrappers = {}
        for row in sheet.iter_rows():
            needed_height = 0
            for cell in row:
                if cell.value is None:
                    continue
                if not cell.alignment.wrap_text:
                    alignment = copy(cell.alignment)
                    alignment.wrap_text = True
                    cell.alignment = alignment
                if cell.data_type != "s" or not cell.value:
                    continue
                width = merged_widths.get((cell.row, cell.column), widths[cell.column])
                available = max(1, int(width) - 2)
                if "\n" not in cell.value and "\r" not in cell.value and len(cell.value) <= available:
                    continue
                if available not in wrappers:
                    wrappers[available] = TextWrapper(width=available)
                paragraphs = cell.value.replace("\r\n", "\n").replace("\r", "\n").split("\n")
                lines = sum(max(1, len(wrappers[available].wrap(line))) for line in paragraphs)
                if lines > 1:
                    line_height = max(15, (cell.font.sz or 11) * 1.35)
                    text_height = lines * line_height + 3
                    needed_height = max(needed_height, text_height)
                    if full_text is not None and text_height > height_limit:
                        full_text.setdefault((cell.row, cell.column), cell.value)
            dimension = sheet.row_dimensions.get(row[0].row)
            current_height = (dimension.height if dimension else None) or sheet.sheet_format.defaultRowHeight or 15
            height = min(height_limit, max(current_height, needed_height))
            if height != current_height:
                sheet.row_dimensions[row[0].row].height = height
        if full_text:
            text_sheet = sheet
            if separate_text:
                text_sheet = workbook.create_sheet(f"{sheet.title} Text")
                for column in "ABCDE":
                    text_sheet.column_dimensions[column].width = 26
            _add_report_text_continuations(sheet, text_sheet, full_text)


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
    full_text = append_report_detail_rows(sheet, details)
    sheet.auto_filter.ref = sheet.dimensions
    for column, header in enumerate([*source_headers, matched_header, "Reply Status"], 1):
        key = "".join(char for char in header.casefold() if char.isalnum())
        sheet.column_dimensions[get_column_letter(column)].width = (
            INBOUND_MESSAGE_COLUMN_WIDTH if "body" in key or "content" in key else
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
    fit_report_text(workbook, {sheet.title: full_text})
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()
