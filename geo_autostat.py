"""Prepare AUTOSTAT directly from the worksheet selected on GEO & RECO."""

from __future__ import annotations

import hashlib
from datetime import datetime
from functools import partial

import streamlit as st

from automation import AutomationError, OB_SORT_VERSION, XLSX_MIME
from autostat import (
    AUTOSTAT_VERSION, ContactSheet, SourceSelection, autostat_source_excluded,
    build_autostat_channel_excel, build_autostat_excel, build_autostat_excel_zip, build_autostat_exports,
)
from email_validation import EMAIL_VALIDATION_VERSION
from folder_output import MANILA_TIMEZONE


GEO_AUTOSTAT_VERSION = 4
GEO_FINANCING_ACCOUNT_VERSION = 1
GEO_AUTOSTAT_CHANNEL_DOWNLOADS_VERSION = 1


def mapped_contact_source(
    name: str, data: bytes, sheet_name: str, header_row: int,
    header_count: int, positions: dict[str, int | None],
) -> SourceSelection:
    """Use the page's column choices while reading the original workbook cells."""
    positions = {target: column for target, column in positions.items() if column is not None}
    if any(column < 0 or column >= header_count for column in positions.values()):
        raise AutomationError("The selected columns do not match the source header row.")
    # Only the selected identity and OB columns may supply deduplication and order.
    # Unselected columns can have old or ambiguous header names.
    headers = [f"Unmapped column {index + 1}" for index in range(header_count)]
    for target in ("indexoneuid", "OB"):
        if target in positions:
            headers[positions[target]] = target
    sheet = ContactSheet(
        sheet_name, header_row, tuple(headers),
        positions.get("cardnumber", positions.get("primary_key")),
        (positions["email"],) if "email" in positions else (),
        (positions["phone"],) if "phone" in positions else (),
    )
    return SourceSelection(name, data, sheet)


def render_geo_autostat(
    name: str, data: bytes, sheet_name: str, header_row: int,
    header_count: int, positions: dict[str, int | None], *, encoding: str,
) -> None:
    """Make the selected sheet's activity workbook available immediately on upload."""
    source = mapped_contact_source(name, data, sheet_name, header_row, header_count, positions)
    if not source.sheet.ready:
        st.session_state.pop("geo_autostat_result", None)
        return
    st.subheader("AUTOSTAT — Email and Viber")
    if autostat_source_excluded(name) or autostat_source_excluded(sheet_name):
        st.session_state.pop("geo_autostat_result", None)
        st.caption("AUTOSTAT excludes SBC INSURANCE, B2 and B4 files and worksheets.")
        return
    with st.expander("AutoStat log settings"):
        remark_by = st.text_input("Remark By", value="MBURGOS", key="geo_autostat_remark_by")
        blank_time = st.checkbox("Leave date and time blank", key="geo_autostat_blank_time")
    signature = (
        GEO_AUTOSTAT_VERSION, AUTOSTAT_VERSION, EMAIL_VALIDATION_VERSION, OB_SORT_VERSION,
        name, hashlib.sha256(data).hexdigest(), sheet_name, header_row,
        tuple(sorted(positions.items())), remark_by, blank_time, encoding,
    )
    result = st.session_state.get("geo_autostat_result")
    if result is None or result["signature"] != signature:
        result = {"signature": signature}
        try:
            with st.spinner("Preparing AUTOSTAT from the selected worksheet..."):
                timestamp = None if blank_time else datetime.now(MANILA_TIMEZONE)
                result.update(
                    bundle=build_autostat_exports([source], timestamp=timestamp, remark_by=remark_by, encoding=encoding),
                    timestamp=timestamp,
                )
        except AutomationError as exc:
            result["error"] = str(exc)
        st.session_state["geo_autostat_result"] = result
    if "error" in result:
        st.warning("AUTOSTAT could not be prepared: " + result["error"])
        return
    bundle = result["bundle"]
    email_count, phone_count = (file.count for file in bundle.files)
    if not (email_count or phone_count):
        st.caption("No valid email or mobile contacts with account numbers were found for AUTOSTAT.")
        return
    st.caption(
        f"Ready from {sheet_name}: {email_count:,} Email accounts and {phone_count:,} Viber accounts. "
        "No second upload is needed. Financing/Card No. (cardnumber) supplies Account No. when available. "
        "Account numbers stay as text; each log is sorted by OB, highest first."
    )
    if any(row["Account precision issues"] for row in bundle.summary):
        st.warning(
            "Some source account cells are numeric and may already be rounded. "
            "Use the original full account numbers stored as Text; missing digits cannot be recovered."
        )
    if not source.sheet.phone_columns:
        st.warning(
            "No phone column is selected for Viber. Open Source columns, enable manual mapping, "
            "and select the mobile-number column for phone."
        )
    elif not phone_count:
        st.warning(
            "No valid mobile numbers with account numbers were found for Viber. Check the selected phone column. "
            "Mobile numbers must normalize to 10 digits starting with 9."
        )
    for column, channel, file in zip(st.columns(2), ("Email", "Viber"), bundle.files):
        with column:
            st.caption(f"{channel}: {file.count:,} valid accounts")
            st.download_button(
                f"Download AUTOSTAT_{channel}.xlsx", partial(build_autostat_channel_excel, file), f"AUTOSTAT_{channel}.xlsx",
                mime=XLSX_MIME, on_click="ignore", key=f"geo_autostat_{channel.lower()}_excel", type="primary",
            )
    st.download_button(
        "Download both AUTOSTAT files (ZIP)", partial(build_autostat_excel_zip, bundle), "AUTOSTAT_FILES.zip",
        mime="application/zip", on_click="ignore", key="geo_autostat_zip",
    )
    st.download_button(
        "Download AUTOSTAT.xlsx (Email + Viber sheets)", partial(build_autostat_excel, bundle), "AUTOSTAT.xlsx",
        mime=XLSX_MIME, on_click="ignore", key="geo_autostat_excel",
    )
