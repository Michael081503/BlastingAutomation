"""Streamlit interface for preparing AUTOSTAT activity logs."""

from __future__ import annotations

import hashlib
from dataclasses import replace
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl.utils import get_column_letter

from automation import AutomationError, CSV_OUTPUT_ENCODINGS
from autostat import (
    AUTOSTAT_VERSION, COMPILED_OUTPUT_VERSION, EMAIL_HEADERS, VIBER_HEADERS, SourceSelection,
    autostat_source_excluded, build_autostat_exports, inspect_contact_source,
)
from email_validation import EMAIL_VALIDATION_VERSION
from export_config import OUTPUT_FOLDER
from folder_output import MANILA_TIMEZONE, save_files_to_folder


AUTOSTAT_PAGE_VERSION = 5
COMPACT_SUMMARY_VERSION = 1
GROUPED_SOURCE_SETTINGS_VERSION = 1


def render_autostat() -> None:
    st.title("AUTOSTAT")
    with st.sidebar:
        output_format = st.selectbox(
            "CSV file type", list(CSV_OUTPUT_ENCODINGS), key="autostat_csv_file_type",
        )
    encoding = CSV_OUTPUT_ENCODINGS[output_format]
    uploads = st.file_uploader(
        "Source files", type=["xlsx", "csv"], accept_multiple_files=True, key="autostat_uploads",
    )
    fixed_timestamp = None
    with st.expander("Log settings"):
        remark_by = st.text_input("Remark By", value="MBURGOS", key="autostat_remark_by")
        blank_timestamp = st.checkbox("Leave date and time blank", key="autostat_blank_time")
        current_time = st.checkbox(
            "Use current Manila date and time", value=True, key="autostat_current_time", disabled=blank_timestamp,
        )
        if not blank_timestamp and not current_time:
            default = st.session_state.setdefault("autostat_default_time", datetime.now(MANILA_TIMEZONE))
            left, right = st.columns(2)
            log_date = left.date_input("Log date (Asia/Manila)", default.date(), key="autostat_date")
            log_time = right.time_input("Log time (Asia/Manila)", default.time().replace(microsecond=0), key="autostat_time")
            fixed_timestamp = datetime.combine(log_date, log_time, tzinfo=MANILA_TIMEZONE)
    if not uploads:
        st.session_state.pop("autostat_result", None)
        return

    sources, configurations, problems, excluded_names = [], [], [], []
    cache = st.session_state.setdefault("autostat_source_info", {})
    active_cache = {}
    source_columns_panel = None

    def inspect(data, extension, **kwargs):
        cache_key = (hashlib.sha256(data).hexdigest(), extension, tuple(kwargs.items()), AUTOSTAT_VERSION)
        if cache_key not in cache:
            cache[cache_key] = inspect_contact_source(data, extension, **kwargs)
        active_cache[cache_key] = cache[cache_key]
        return cache[cache_key]

    for file_index, upload in enumerate(uploads):
        if autostat_source_excluded(upload.name):
            excluded_names.append(upload.name)
            configurations.append((upload.name, "excluded"))
            continue
        data = upload.getvalue()
        extension = Path(upload.name).suffix.lower()
        digest = hashlib.sha256(data).hexdigest()
        prefix = f"autostat_{AUTOSTAT_VERSION}_{file_index}_{digest}"
        configurations.append((upload.name, digest))
        try:
            sheets = inspect(data, extension)
            if source_columns_panel is None:
                source_columns_panel = st.expander("Edit source columns", expanded=False)
            with source_columns_panel:
                st.write(f"**{file_index + 1}. {upload.name}**")
                selected = st.multiselect(
                    "Sheets to process", [sheet.name for sheet in sheets],
                    default=[sheet.name for sheet in sheets if sheet.ready], key=f"{prefix}_sheets",
                )
                configurations.append(tuple(selected))
                for sheet in sheets:
                    if sheet.name not in selected:
                        continue
                    st.write(f"**{sheet.name}**")
                    sheet_key = f"{prefix}_{sheet.name}"
                    header_row = int(st.number_input(
                        "Header row", min_value=1, max_value=1_048_575, value=sheet.header_row,
                        key=f"{sheet_key}_header",
                    ))
                    if header_row != sheet.header_row:
                        sheet = inspect(data, extension, sheet_name=sheet.name, header_row=header_row)[0]
                    column_key = f"{sheet_key}_{header_row}"
                    columns = list(range(len(sheet.headers)))

                    def label(index, headers=sheet.headers):
                        return "Choose a column" if index is None else f"{get_column_letter(index + 1)}: {headers[index] or '(blank header)'}"

                    account_options = [None, *columns]
                    account = st.selectbox(
                        "Account number column", account_options,
                        index=account_options.index(sheet.account_column), format_func=label,
                        key=f"{column_key}_account",
                    )
                    email_columns = st.multiselect(
                        "Email columns", columns, default=list(sheet.email_columns),
                        format_func=label, key=f"{column_key}_emails",
                    )
                    phone_columns = st.multiselect(
                        "Phone columns", columns, default=list(sheet.phone_columns),
                        format_func=label, key=f"{column_key}_phones",
                    )
                    configurations.append((sheet.name, header_row, account, tuple(email_columns), tuple(phone_columns)))
                    if account is None or not (email_columns or phone_columns):
                        problems.append(f"{upload.name} / {sheet.name}: select an account column and at least one contact column.")
                    elif account in (*email_columns, *phone_columns):
                        problems.append(f"{upload.name} / {sheet.name}: the account column must be separate from contacts.")
                    else:
                        sources.append(SourceSelection(upload.name, data, replace(
                            sheet, account_column=account, email_columns=tuple(email_columns), phone_columns=tuple(phone_columns),
                        )))
                unselected = [sheet.name for sheet in sheets if sheet.name not in selected]
                if unselected:
                    st.caption("Not processed: " + ", ".join(unselected))
                if not selected:
                    st.info("Select a sheet and its columns to include this file.")
        except (AutomationError, UnicodeError, ValueError) as exc:
            problems.append(f"{upload.name}: {exc}")
    st.session_state["autostat_source_info"] = active_cache
    if excluded_names:
        with st.expander(f"Excluded files ({len(excluded_names)})"):
            for name in excluded_names:
                st.write(name)
    for problem in problems:
        st.error(problem)
    if not sources:
        st.info("Select at least one sheet with an account number column and an email or phone column.")
    signature = (
        AUTOSTAT_VERSION, COMPILED_OUTPUT_VERSION, AUTOSTAT_PAGE_VERSION, EMAIL_VALIDATION_VERSION, tuple(configurations), remark_by,
        blank_timestamp, current_time, fixed_timestamp, encoding, EMAIL_HEADERS, VIBER_HEADERS,
    )
    previous = st.session_state.get("autostat_result")
    if previous is not None and (previous["signature"] != signature or problems):
        st.session_state.pop("autostat_result", None)

    if st.button("Preview and prepare AUTOSTAT files", type="primary", key="autostat_prepare", disabled=bool(problems) or not sources):
        st.session_state.pop("autostat_result", None)
        try:
            with st.spinner("Checking contacts and preparing AUTOSTAT logs..."):
                timestamp = None if blank_timestamp else fixed_timestamp or datetime.now(MANILA_TIMEZONE)
                bundle = build_autostat_exports(sources, timestamp=timestamp, remark_by=remark_by, encoding=encoding)
                st.session_state["autostat_result"] = {"signature": signature, "bundle": bundle, "timestamp": timestamp}
        except AutomationError as exc:
            st.error(str(exc))
        except Exception:
            st.error("The source files could not be processed. Check their selected sheets and columns, then try a fresh Excel or CSV copy.")
    result = st.session_state.get("autostat_result")
    if result is None:
        return
    bundle = result["bundle"]
    email_count = bundle.files[0].count
    phone_count = bundle.files[1].count
    email_column, phone_column = st.columns(2)
    email_column.metric("Valid email accounts", f"{email_count:,}")
    phone_column.metric("Valid phone accounts", f"{phone_count:,}")
    if not (email_count or phone_count):
        st.warning("No valid contacts with account numbers were found. The output files contain template headers only.")
    st.download_button(
        "Download all AUTOSTAT files (ZIP)", bundle.zip_data, "AUTOSTAT_FILES.zip",
        mime="application/zip", on_click="ignore", key="autostat_zip",
    )
    for index, file in enumerate(bundle.files):
        st.download_button(
            f"Download {file.filename}", file.data, file.filename, mime="text/csv",
            key=f"autostat_download_{index}", on_click="ignore",
        )
    with st.expander("Save to a folder"):
        output_folder = st.text_input("Output folder", value=OUTPUT_FOLDER, key="autostat_output_folder")
        if st.button("Save AUTOSTAT files to folder", key="autostat_save"):
            try:
                folder = save_files_to_folder(
                    [(file.filename, file.data) for file in bundle.files], output_folder, batch_name="AUTOSTAT",
                )
                result["saved_folder"] = str(folder)
                st.session_state["autostat_result"] = result
            except AutomationError as exc:
                st.error(str(exc))
        if result.get("saved_folder"):
            st.success(f"Saved {len(bundle.files)} files:")
            st.code(result["saved_folder"], language=None)
    with st.expander("Preview"):
        filename = st.selectbox(
            "File to preview", [file.filename for file in bundle.files],
            key=f"autostat_preview_{hash(signature)}",
        )
        file = next(file for file in bundle.files if file.filename == filename)
        st.dataframe(pd.DataFrame(file.preview, columns=file.headers), hide_index=True, width="stretch")
        st.caption(f"Showing the first {min(file.count, 100):,} of {file.count:,} rows.")
