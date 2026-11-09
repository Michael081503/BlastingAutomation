"""Group accounts and create separate Email/Viber CSV files."""

from __future__ import annotations

import re
import csv
from collections import Counter, defaultdict
from dataclasses import dataclass
from io import BytesIO, StringIO
from pathlib import Path
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl import Workbook

from automation import (
    AutomationError,
    IDENTIFIER_FIELDS,
    csv_template_headers,
    campaign_output_headers,
    template_column_positions,
    excel_template_to_csv,
    fill_csv_template,
    geo_doclass,
    has_contact_details,
    has_contact_value,
    normalize_header,
    normalize_product,
    ob_sort_key,
    source_column_positions,
    write_typed_template_sheet,
)
from export_config import (
    CHANNEL_SUFFIXES, GEO_AGENTS, GEO_OUTPUT_FILENAMES, GEO_ZIP_FILENAME,
    OUTPUT_FILENAMES, RECO_DOCLASS_BY_FILENAME, ZIP_FILENAME,
)
from email_validation import validated_email
from agent_cycle import assign_agents_to_accounts


GROUP_LABELS = {
    "qualified_cards_level_1": "QUALIFIED / CARDS / Level 1",
    "qualified_cards_level_6": "QUALIFIED / CARDS / Level 6",
    "qualified_bel": "QUALIFIED / BEL / All levels",
    "blank_cards_level_1": "BLANK / CARDS / Level 1",
    "blank_cards_level_6": "BLANK / CARDS / Level 6",
    "blank_bel": "BLANK / BEL / All levels",
    "blank_other_products": "BLANK / Other than CARDS and BEL / All levels",
    "qualified_review": "QUALIFIED / Other products or levels (review)",
    "blank_review": "BLANK / CARDS / Other levels (review)",
}
REVIEW_GROUP_KEYS = frozenset({"qualified_review", "blank_review"})
CHANNEL_LABELS = {"email": "Email", "viber": "Viber"}
GEO_FILENAME_CONFIG_VERSION = 3
GEO_LOANS_VERSION = 1
GEO_LOAN_PRODUCTS = frozenset({"csl", "pl", "sla"})
GROUPED_SHEET_EXPORT_VERSION = 1
EXCEL_EXPORT_VERSION = 2
TYPED_EXCEL_EXPORT_VERSION = 1
CAMPAIGN_OUTPUT_FORMAT_VERSION = 2
CSV_CAMPAIGN_DEFAULT_VERSION = 1
EXCEL_SAFE_CSV_VERSION = 3
CSV_PROTECTED_IDENTIFIER_FIELDS = IDENTIFIER_FIELDS - {"indexoneuid"}
ACTIVITY_CSV_IDENTIFIERS_VERSION = 1
REQUIRED_CSV_PROTECTION_VERSION = 1
AGENT_ASSIGNMENT_EXPORT_VERSION = 2
RECO_ONLY_AGENTS_VERSION = 1
GEO_FIXED_AGENTS_VERSION = 1
GEO_AGENT_AVAILABILITY_VERSION = 1
RECO_DOCLASS_VERSION = 1
SOURCE_MEMBERSHIP_VERSION = 1
SOURCE_ACCOUNT_FIELDS = ("primary_key", "cardnumber", "indexoneuid")
AREA_LABELS = {"mindanao": "Mindanao", "r3": "R3", "area_review": "Other or missing Area (review)"}
GEO_PRODUCT_GROUP_LABELS = {
    "qualified_cards": "QUALIFIED / CARDS",
    "qualified_bel": "QUALIFIED / BEL",
    "blank_cards": "BLANK / CARDS",
    "blank_bel": "BLANK / BEL",
    "blank_other_products": "LOANS / CSL, PL, SLA and other loan products",
    "qualified_review": "QUALIFIED / Other products (review)",
}
GEO_GROUP_LABELS = {
    f"{area}_{key}": f"{label} / {group_label}"
    for area, label in AREA_LABELS.items() for key, group_label in GEO_PRODUCT_GROUP_LABELS.items()
}
GEO_REVIEW_GROUP_KEYS = frozenset(
    f"{area}_{key}" for area in AREA_LABELS for key in GEO_PRODUCT_GROUP_LABELS
    if area == "area_review" or key == "qualified_review"
)


@dataclass
class ExportFile:
    key: str
    group_key: str
    channel: str
    label: str
    filename: str
    data: bytes
    count: int
    missing_emails: int
    headers: list[str]
    preview: list[dict[str, Any]]
    records: list[dict[str, Any]]
    doclass: str | None = None


@dataclass
class ExportBundle:
    files: list[ExportFile]
    counts: dict[str, int]
    zip_data: bytes
    zip_filename: str
    no_contact_rows_skipped: int


def source_account_counts(records: list[dict[str, Any]]) -> Counter[tuple[str, ...]]:
    """Retain complete account identities and duplicate counts from the source."""
    return Counter(
        tuple("" if record.get(field) is None else str(record[field]).strip() for field in SOURCE_ACCOUNT_FIELDS)
        for record in records
    )


def validate_source_exports(
    files: list[ExportFile], source_accounts: Counter[tuple[str, ...]], *, source_sheet: str,
) -> None:
    """Block CSV and Excel downloads outside the selected sheet's account counts."""
    if not source_sheet.strip():
        raise AutomationError("Downloads blocked: select a source worksheet first.")
    channels = defaultdict(Counter)
    for file in files:
        encoding = "utf-8-sig" if file.data.startswith(b"\xef\xbb\xbf") else "cp1252"
        rows = csv.reader(StringIO(file.data.decode(encoding), newline=""))
        positions = source_column_positions(next(rows, []))
        csv_accounts = source_account_counts([
            {field: row[positions[field]] if field in positions and positions[field] < len(row) else None
             for field in SOURCE_ACCOUNT_FIELDS}
            for row in rows
        ])
        if csv_accounts != source_account_counts(file.records) or sum(csv_accounts.values()) != file.count:
            raise AutomationError(
                f"Downloads blocked: {file.filename} has inconsistent account data. Clear all files and upload the source again."
            )
        channels[file.channel].update(csv_accounts)
    if any(accounts - source_accounts for accounts in channels.values()):
        raise AutomationError(
            f'Downloads blocked: an output account is not verified against the selected worksheet "{source_sheet}". '
            "Clear all files and upload the current source workbook again."
        )


def build_export_excel(file: ExportFile) -> bytes:
    """Create the complete export with native cell types and text identifiers."""
    workbook = Workbook(write_only=True)
    write_typed_template_sheet(workbook, file.records, file.headers, "Accounts", doclass=file.doclass)
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def build_excel_zip(files: list[ExportFile]) -> bytes:
    """Keep campaign filenames, native cell types and exact account identifiers."""
    output = BytesIO()
    with ZipFile(output, "w", ZIP_DEFLATED) as archive:
        for file in files:
            archive.writestr(Path(file.filename).with_suffix(".xlsx").name, build_export_excel(file))
    return output.getvalue()


def build_export_csv(file: ExportFile, *, protect_identifiers: bool = True) -> bytes:
    """Keep CSV identifiers as text in Excel, including subsequent CSV saves.

    A leading tab survives Excel's CSV save/reopen cycle, unlike a text formula.
    This is part of the CSV field value; plain upload bytes remain available.
    Account numbers and phones are prefixed. indexoneuid stays unchanged.
    """
    if not protect_identifiers:
        return file.data
    encoding = "utf-8-sig" if file.data.startswith(b"\xef\xbb\xbf") else "cp1252"
    rows = csv.reader(StringIO(file.data.decode(encoding), newline=""))
    output = StringIO(newline="")
    writer = csv.writer(output, lineterminator="\r\n")
    headers = next(rows)
    writer.writerow(headers)
    fields = {normalize_header(field) for field in CSV_PROTECTED_IDENTIFIER_FIELDS}
    columns = {index for index, header in enumerate(headers) if normalize_header(header) in fields}
    columns.update(index for target, index in source_column_positions(headers).items() if target in CSV_PROTECTED_IDENTIFIER_FIELDS)
    for row in rows:
        for column in columns:
            if row[column] and not row[column].startswith("\t"):
                row[column] = "\t" + row[column]
        writer.writerow(row)
    return output.getvalue().encode(encoding)


def build_csv_zip(files: list[ExportFile], *, protect_identifiers: bool = True) -> bytes:
    output = BytesIO()
    with ZipFile(output, "w", ZIP_DEFLATED) as archive:
        for file in files:
            archive.writestr(file.filename, build_export_csv(file, protect_identifiers=protect_identifiers))
    return output.getvalue()


def campaign_output_files(files: list[ExportFile], *, excel: bool, protect_csv_identifiers: bool = True) -> list[tuple[str, bytes]]:
    """Use the chosen format consistently for folder saves and downloads."""
    return [(Path(file.filename).with_suffix(".xlsx").name, build_export_excel(file)) if excel
            else (file.filename, build_export_csv(file, protect_identifiers=protect_csv_identifiers)) for file in files]


def _level(value: Any) -> str:
    text = str(value or "").strip().casefold()
    match = re.fullmatch(r"(?:level\s*)?(\d+)(?:\.0+)?", text)
    return f"level{int(match.group(1))}" if match else normalize_header(value)


def separate_negative(records: list[dict[str, Any]]) -> dict[str, list[dict[str, Any]]]:
    """Keep source order and group every contactable row exactly once.

    Both statuses split CARDS by Level 1/6 and include BEL at every level.
    Other blank-status products get one file together. Unmatched levels/products
    get a review file. Accounts without a valid email or phone are excluded.
    """
    groups = {key: [] for key in GROUP_LABELS}
    for record in records:
        record = {**record, "product": normalize_product(record.get("product"))}
        if not has_contact_details(record):
            continue
        qualified = normalize_header(record.get("ChristmasAmnesty")) == "qualified"
        product = normalize_header(record.get("product"))
        level = _level(record.get("level"))
        prefix = "qualified" if qualified else "blank"
        if product == "cards" and level == "level1":
            key = f"{prefix}_cards_level_1"
        elif product == "cards" and level == "level6":
            key = f"{prefix}_cards_level_6"
        elif product == "bel":
            key = f"{prefix}_bel"
        elif not qualified and product != "cards":
            key = "blank_other_products"
        else:
            key = f"{prefix}_review"
        # Also enforce the template value for callers that supply records directly.
        groups[key].append({**record, "ChristmasAmnesty": "QUALIFIED" if qualified else None})
    return groups


def separate_geo(records: list[dict[str, Any]]) -> dict[str, list[dict[str, Any]]]:
    """Group each Area by Christmas Amnesty and product, combining all levels."""
    groups = {key: [] for key in GEO_GROUP_LABELS}
    for record in records:
        record = {**record, "product": normalize_product(record.get("product"))}
        if not has_contact_details(record):
            continue
        area = normalize_header(record.get("AREA"))
        area_key = "mindanao" if area == "mindanao" else "r3" if area in {"r3", "region3"} else "area_review"
        qualified = normalize_header(record.get("ChristmasAmnesty")) == "qualified"
        product = normalize_header(record.get("product"))
        # These loan products always use the existing area's LOANS file,
        # including qualified accounts. Preserve each row's original status.
        if product in GEO_LOAN_PRODUCTS:
            key = "blank_other_products"
        elif product in {"cards", "bel"}:
            key = f"{'qualified' if qualified else 'blank'}_{product}"
        else:
            key = "qualified_review" if qualified else "blank_other_products"
        groups[f"{area_key}_{key}"].append({
            **record, "ChristmasAmnesty": "QUALIFIED" if qualified else None,
        })
    return groups


def export_filename(name: str, extension: str) -> str:
    """Allow editable basenames while keeping downloads valid on Windows."""
    if not isinstance(name, str) or not name.strip():
        raise AutomationError("Enter a filename for every output.")
    name = name.strip()
    if re.search(r'[<>:"/\\|?*\x00-\x1f]', name) or name.endswith("."):
        raise AutomationError(f'Invalid filename "{name}". Use a plain filename without folder paths or reserved characters.')
    if name.lower().endswith((".xlsx", ".csv", ".zip")):
        name = name.rsplit(".", 1)[0]
    if not name or re.fullmatch(r"(?:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\..*)?", name, re.I):
        raise AutomationError(f'Invalid filename "{name}". Choose another name.')
    return name + extension


def separate_channels(records: list[dict[str, Any]]) -> dict[str, list[dict[str, Any]]]:
    """Route both-contact accounts twice and blank the other contact column.

    The source records are not changed. Source order and every other mapped
    value are retained in each channel. Only valid, normalized email addresses
    enter Email files; an invalid address does not prevent Viber routing.
    """
    channels = {key: [] for key in CHANNEL_LABELS}
    for record in records:
        record = {**record, "product": normalize_product(record.get("product")), "doclass": geo_doclass(record.get("AREA"))}
        email = validated_email(record.get("email"))
        if email is not None:
            channels["email"].append({**record, "email": email, "phone": None})
        if has_contact_value(record.get("phone")):
            channels["viber"].append({**record, "email": None})
    return channels


def channel_filename(base_name: str, channel: str, extension: str) -> str:
    if channel not in CHANNEL_SUFFIXES:
        raise AutomationError(f'Missing suffix for "{channel}" in export_config.py.')
    base = export_filename(base_name, extension)
    suffix = CHANNEL_SUFFIXES[channel]
    if not isinstance(suffix, str) or not suffix.strip():
        raise AutomationError(f'The "{channel}" suffix is empty. Check export_config.py.')
    return export_filename(base[:-len(extension)] + suffix, extension)


def reco_doclass(filename: str) -> str:
    """Match whole filename labels in L1, L6, PL priority order."""
    labels = set(re.findall(r"[a-z0-9]+", Path(filename).stem.casefold()))
    return next((value for label, value in RECO_DOCLASS_BY_FILENAME.items() if label.casefold() in labels), "")


def _build_channel_files(
    template: bytes,
    records: list[dict[str, Any]],
    *,
    base_name: str,
    group_key: str,
    group_label: str,
    encoding: str,
    agents: tuple[tuple[str, str], ...] | None = None,
    reco: bool = False,
) -> list[ExportFile]:
    if not records:
        return []
    records = sorted(records, key=lambda record: ob_sort_key(record.get("OB")), reverse=True)
    if agents is not None:
        records = assign_agents_to_accounts(records, agents)
    headers = csv_template_headers(template)
    required = [("Monthly", "Monthly"), ("Terms", "CHCode")] if agents is not None else []
    if reco:
        required.append(("doclass", "doclass"))
    positions = template_column_positions(headers)
    missing = [header for target, header in required if target not in positions]
    if missing:
        output = StringIO(newline="")
        csv.writer(output).writerow([*headers, *missing])
        template = output.getvalue().encode("utf-8-sig")
    headers = csv_template_headers(template)
    files = []
    for channel, rows in separate_channels(records).items():
        filename = channel_filename(base_name, channel, ".csv")
        doclass = reco_doclass(filename) if reco else None
        if reco:
            rows = [{**row, "doclass": doclass} for row in rows]
        data = fill_csv_template(template, rows, encoding=encoding, doclass=doclass)
        files.append(ExportFile(
            key=f"{group_key}_{channel}",
            group_key=group_key,
            channel=channel,
            label=f"{group_label} / {CHANNEL_LABELS[channel]}",
            filename=filename,
            data=data,
            count=len(rows),
            missing_emails=sum(not has_contact_value(row.get("email")) for row in rows),
            headers=headers,
            preview=rows[:100],
            records=rows,
            doclass=doclass,
        ))
    return files


def _bundle(
    files: list[ExportFile], counts: dict[str, int], zip_name: str, excluded: int,
) -> ExportBundle:
    names = [file.filename.casefold() for file in files]
    if len(set(names)) != len(names):
        raise AutomationError("Output filenames must be unique. Give each file a different name.")
    output = BytesIO()
    with ZipFile(output, "w", compression=ZIP_DEFLATED) as archive:
        for file in files:
            archive.writestr(file.filename, file.data)
    return ExportBundle(files, counts, output.getvalue(), zip_name, excluded)


def _csv_export_template(
    template: bytes, sheet_name: str, header_row: int, template_extension: str,
) -> bytes:
    if template_extension == ".xlsx":
        return excel_template_to_csv(template, sheet_name, header_row)
    if template_extension == ".csv":
        output = StringIO(newline="")
        csv.writer(output).writerow(campaign_output_headers(csv_template_headers(template)))
        return output.getvalue().encode("utf-8-sig")
    raise AutomationError("Upload an .xlsx or .csv template. Outputs are always CSV.")


def build_sheet_exports(
    template: bytes,
    records: list[dict[str, Any]],
    sheet_name: str,
    header_row: int,
    *,
    base_name: str,
    label: str,
    template_extension: str = ".xlsx",
    encoding: str = "cp1252",
    agents: tuple[tuple[str, str], ...] | None = None,
    reco: bool = False,
) -> ExportBundle:
    """Export CSV channels using an Excel or CSV template's headers."""
    template = _csv_export_template(template, sheet_name, header_row, template_extension)
    contactable = [record for record in records if has_contact_details(record)]
    files = _build_channel_files(
        template, contactable, base_name=base_name,
        group_key="all", group_label=label, encoding=encoding, agents=agents, reco=reco,
    )
    return _bundle(
        files, {"all": len(contactable)}, export_filename(base_name, ".zip"),
        len(records) - len(contactable),
    )


def build_negative_exports(
    template: bytes,
    records: list[dict[str, Any]],
    sheet_name: str,
    header_row: int,
    template_extension: str = ".xlsx",
    *,
    encoding: str = "cp1252",
    source_sheet: str = "NEGATIVE",
    agents: tuple[tuple[str, str], ...] | None = None,
) -> ExportBundle:
    """Apply the NEGATIVE grouping rules to any source sheet."""
    return _build_grouped_exports(
        template, records, sheet_name, header_row, template_extension,
        groups=separate_negative(records), labels=GROUP_LABELS, names=OUTPUT_FILENAMES,
        zip_name=ZIP_FILENAME, encoding=encoding, source_sheet=source_sheet, agents=agents,
        reco=True,
    )


def build_geo_exports(
    template: bytes,
    records: list[dict[str, Any]],
    sheet_name: str,
    header_row: int,
    template_extension: str = ".xlsx",
    *,
    encoding: str = "cp1252",
    source_sheet: str = "NEGATIVE",
    agents: tuple[tuple[str, str], ...] | None = None,
    available_agent_codes: tuple[str, ...] | None = None,
) -> ExportBundle:
    """Distribute GEO accounts across its fixed agent order.

    Availability filters GEO's roster while retaining its fixed order.
    The legacy agents keyword is ignored so RECO's cycle cannot assign GEO.
    """
    selected_agents = GEO_AGENTS
    if available_agent_codes is not None:
        known_codes = {code for _, code in GEO_AGENTS}
        if (not available_agent_codes or len(set(available_agent_codes)) != len(available_agent_codes)
                or not set(available_agent_codes) <= known_codes):
            raise AutomationError("Choose at least one available GEO agent from the configured roster.")
        selected_agents = tuple(agent for agent in GEO_AGENTS if agent[1] in available_agent_codes)
    template = _csv_export_template(template, sheet_name, header_row, template_extension)
    if not any(normalize_header(header) == "area" for header in csv_template_headers(template)):
        raise AutomationError("GEO exports need an AREA header in the template. Use the built-in 16-column template.")
    return _build_grouped_exports(
        template, records, "", 1, ".csv", groups=separate_geo(records),
        labels=GEO_GROUP_LABELS, names=GEO_OUTPUT_FILENAMES, zip_name=GEO_ZIP_FILENAME, encoding=encoding,
        source_sheet=source_sheet, agents=selected_agents,
    )


def _build_grouped_exports(
    template: bytes, records: list[dict[str, Any]], sheet_name: str, header_row: int,
    template_extension: str, *, groups: dict[str, list[dict[str, Any]]],
    labels: dict[str, str], names: dict[str, str], zip_name: str, encoding: str,
    source_sheet: str = "NEGATIVE",
    agents: tuple[tuple[str, str], ...] | None = None,
    reco: bool = False,
) -> ExportBundle:
    template = _csv_export_template(template, sheet_name, header_row, template_extension)
    counts = {key: len(rows) for key, rows in groups.items()}
    safe_sheet = re.sub(r'[^\w .-]+', "_", source_sheet).strip(" .") or "sheet"
    prefix = "" if normalize_header(source_sheet) == "negative" else f"{safe_sheet} - "
    filenames = {}
    for key in groups:
        if key not in names:
            raise AutomationError(f'Missing filename for "{key}" in export_config.py.')
        filenames[key] = export_filename(prefix + export_filename(names[key], ".csv"), ".csv")
    if len({name.casefold() for name in filenames.values()}) != len(filenames):
        raise AutomationError("Export filenames must be unique. Check export_config.py.")
    zip_name = export_filename(prefix + export_filename(zip_name, ".zip"), ".zip")
    files = []
    for key, rows in groups.items():
        if not rows:
            continue
        files.extend(_build_channel_files(
            template, rows, base_name=filenames[key],
            group_key=key, group_label=labels[key], encoding=encoding, agents=agents, reco=reco,
        ))
    return _bundle(
        files, counts, zip_name,
        excluded=sum(not has_contact_details(record) for record in records),
    )
