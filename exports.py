"""Group accounts and fill separate Email/Viber copies of the same template."""

from __future__ import annotations

import re
from dataclasses import dataclass
from io import BytesIO
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile

from automation import (
    AutomationError,
    fill_csv_template,
    fill_excel_template,
    has_contact_details,
    has_contact_value,
    normalize_header,
)
from export_config import CHANNEL_SUFFIXES, OUTPUT_FILENAMES, ZIP_FILENAME


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
    preview: list[dict[str, Any]]


@dataclass
class ExportBundle:
    files: list[ExportFile]
    counts: dict[str, int]
    zip_data: bytes
    zip_filename: str
    no_contact_rows_skipped: int


def _level(value: Any) -> str:
    text = str(value or "").strip().casefold()
    match = re.fullmatch(r"(?:level\s*)?(\d+)(?:\.0+)?", text)
    return f"level{int(match.group(1))}" if match else normalize_header(value)


def separate_negative(records: list[dict[str, Any]]) -> dict[str, list[dict[str, Any]]]:
    """Keep source order and group every contactable row exactly once.

    Both statuses split CARDS by Level 1/6 and include BEL at every level.
    Other blank-status products get one file together. Unmatched levels/products
    get a review file. Accounts missing both email and phone are excluded.
    """
    groups = {key: [] for key in GROUP_LABELS}
    for record in records:
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


def export_filename(name: str, extension: str) -> str:
    """Allow editable basenames while keeping downloads valid on Windows."""
    if not isinstance(name, str) or not name.strip():
        raise AutomationError("An export filename is empty. Check export_config.py.")
    name = name.strip()
    if re.search(r'[<>:"/\\|?*\x00-\x1f]', name) or name.endswith("."):
        raise AutomationError(f'Invalid export filename "{name}". Edit export_config.py.')
    if name.lower().endswith((".xlsx", ".csv", ".zip")):
        name = name.rsplit(".", 1)[0]
    if not name or re.fullmatch(r"(?:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\..*)?", name, re.I):
        raise AutomationError(f'Invalid export filename "{name}". Edit export_config.py.')
    return name + extension


def separate_channels(records: list[dict[str, Any]]) -> dict[str, list[dict[str, Any]]]:
    """Route both-contact accounts twice and blank the other contact column.

    The source records are not changed. Source order and every other mapped
    value are retained in each channel; neither-contact records go nowhere.
    """
    channels = {key: [] for key in CHANNEL_LABELS}
    for record in records:
        if has_contact_value(record.get("email")):
            channels["email"].append({**record, "phone": None})
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


def _build_channel_files(
    template: bytes,
    records: list[dict[str, Any]],
    sheet_name: str,
    header_row: int,
    *,
    base_name: str,
    group_key: str,
    group_label: str,
    extension: str,
) -> list[ExportFile]:
    if not records:
        return []
    files = []
    for channel, rows in separate_channels(records).items():
        filename = channel_filename(base_name, channel, extension)
        data = (
            fill_csv_template(template, rows) if extension == ".csv"
            else fill_excel_template(template, rows, sheet_name, header_row)
        )
        files.append(ExportFile(
            key=f"{group_key}_{channel}",
            group_key=group_key,
            channel=channel,
            label=f"{group_label} / {CHANNEL_LABELS[channel]}",
            filename=filename,
            data=data,
            count=len(rows),
            missing_emails=sum(not has_contact_value(row.get("email")) for row in rows),
            preview=rows[:100],
        ))
    return files


def _bundle(
    files: list[ExportFile], counts: dict[str, int], zip_name: str, excluded: int,
) -> ExportBundle:
    names = [file.filename.casefold() for file in files]
    if len(set(names)) != len(names):
        raise AutomationError("Export filenames must be unique. Check export_config.py and CHANNEL_SUFFIXES.")
    output = BytesIO()
    with ZipFile(output, "w", compression=ZIP_DEFLATED) as archive:
        for file in files:
            archive.writestr(file.filename, file.data)
    return ExportBundle(files, counts, output.getvalue(), zip_name, excluded)


def build_sheet_exports(
    template: bytes,
    records: list[dict[str, Any]],
    sheet_name: str,
    header_row: int,
    *,
    base_name: str,
    label: str,
    extension: str = ".xlsx",
) -> ExportBundle:
    """Apply the same channel split when a sheet is exported as one group."""
    contactable = [record for record in records if has_contact_details(record)]
    files = _build_channel_files(
        template, contactable, sheet_name, header_row, base_name=base_name,
        group_key="all", group_label=label, extension=extension,
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
    extension: str = ".xlsx",
) -> ExportBundle:
    groups = separate_negative(records)
    counts = {key: len(rows) for key, rows in groups.items()}
    filenames = {}
    for key in groups:
        if key not in OUTPUT_FILENAMES:
            raise AutomationError(f'Missing filename for "{key}" in export_config.py.')
        filenames[key] = export_filename(OUTPUT_FILENAMES[key], extension)
    if len({name.casefold() for name in filenames.values()}) != len(filenames):
        raise AutomationError("Export filenames must be unique. Check export_config.py.")
    zip_name = export_filename(ZIP_FILENAME, ".zip")
    files = []
    for key, rows in groups.items():
        if not rows:
            continue
        files.extend(_build_channel_files(
            template, rows, sheet_name, header_row, base_name=filenames[key],
            group_key=key, group_label=GROUP_LABELS[key], extension=extension,
        ))
    return _bundle(
        files, counts, zip_name,
        excluded=sum(not has_contact_details(record) for record in records),
    )
