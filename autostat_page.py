"""Streamlit interface for preparing AUTOSTAT activity logs."""

from __future__ import annotations

import hashlib
from dataclasses import replace
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl.utils import get_column_letter

from automation import AutomationError, CSV_OUTPUT_ENCODINGS, OB_SORT_VERSION, normalize_header
from autostat import (
    AUTOSTAT_VERSION, AUTOSTAT_REMARK_BY, COMPILED_OUTPUT_VERSION, EMAIL_HEADERS, VIBER_HEADERS, SourceSelection,
    autostat_source_excluded, build_autostat_exports, inspect_contact_source, build_autostat_channel_excel,
)
from autostat_downloads import render_autostat_downloads
from b2_autostat import (
    B2_AUTOSTAT_VERSION, B2_COLLECTOR_TERMS_VERSION, B2_EMAIL_HEADERS, B2_PHONE_HEADERS, B2_PRIMARY_KEY_VERSION, B2ActivitySource,
    b2_account_sheet, b2_metadata_columns, build_b2_autostat_exports, build_b4_autostat_exports, B4_AUTOSTAT_VERSION,
)
from email_validation import EMAIL_VALIDATION_VERSION
from export_config import OUTPUT_FOLDER
from folder_output import MANILA_TIMEZONE, save_files_to_folder
from upload_state import upload_widget_key


AUTOSTAT_PAGE_VERSION = 12
AUTOSTAT_AUTOMATIC_UPLOADS_VERSION = 1
B4_ACTIVITY_PANEL_VERSION = 1
B2_UPLOAD_VERSION = 1
B2_ACTIVITY_PANEL_VERSION = 1
COMPACT_SUMMARY_VERSION = 1
GROUPED_SOURCE_SETTINGS_VERSION = 1


def render_autostat() -> None:
    st.title("AUTOSTAT")
    st.write(
        "Upload the automation's output files to prepare activity logs automatically. "
        "The original workbook or endorsement database is not needed."
    )
    upload_type = st.radio("Upload type", ("Standard", "B2", "B4"), horizontal=True, key="autostat_upload_type")
    if upload_type == "Standard":
        st.caption(
            "Standard uses the same Email and Viber templates as GEO & RECO. "
            "Upload the generated Email and Viber files together to include both channels. "
            "Use B2 or B4 for those bucket outputs."
        )
    _render_autostat_uploads(b2=upload_type != "Standard", bucket=4 if upload_type == "B4" else 2,
                             compact=upload_type == "B4")


def render_b2_autostat_uploads(*, encoding: str) -> None:
    """Offer independent B2 uploads and activity downloads on NEGATIVE B2."""
    uploads = st.file_uploader(
        "B2 source files", type=["xlsx", "csv"], accept_multiple_files=True, key=upload_widget_key("geo_b2_autostat_uploads"),
    )
    if not uploads:
        st.session_state.pop("geo_b2_autostat_result", None)
        return
    _render_autostat_uploads(b2=True, encoding=encoding, state_prefix="geo_b2_autostat_",
                             uploads=uploads)


def render_b4_autostat_uploads(*, encoding: str) -> None:
    """Automatically prepare B4 logs from the CARDS/BEL files in their own panel."""
    uploads = st.file_uploader(
        "B4 source files", type=["xlsx", "csv"], accept_multiple_files=True,
        key=upload_widget_key("b4_autostat_uploads"),
        help="Upload the B4 CARDS/BEL output files. The endorsement database is not needed for these activity logs.",
    )
    if not uploads:
        st.session_state.pop("b4_autostat_result", None)
        return
    _render_autostat_uploads(b2=True, bucket=4, encoding=encoding, state_prefix="b4_autostat_",
                             uploads=uploads, compact=True)


def _render_autostat_uploads(
    *, b2: bool, encoding: str | None = None, state_prefix: str | None = None,
    uploads=None, bucket: int = 2, compact: bool = False,
) -> None:
    state_prefix = state_prefix or (f"autostat_b{bucket}_" if b2 else "autostat_")
    result_key = state_prefix + "result"
    bucket_label = f"B{bucket}"
    workbook_name = f"AUTOSTAT_{bucket_label}" if b2 else "AUTOSTAT"
    options = st.expander(f"{bucket_label} AutoStat options") if compact else None
    if b2 and not compact:
        st.subheader(f"{bucket_label} activity logs")
        st.caption(
            f"Upload {bucket_label} account files to fill the Email, SMS and Viber templates. "
            "Names are copied from the source; Terms supplies the collector agent code when available."
        )
    if encoding is None:
        with st.sidebar:
            output_format = st.selectbox(
                "Output files", ("Excel (.xlsx)", "CSV (.csv)"), key=f"{state_prefix}output_format",
            )
            csv_file_type = st.selectbox(
                "CSV file type", list(CSV_OUTPUT_ENCODINGS), key=f"{state_prefix}csv_file_type",
            )
        encoding = CSV_OUTPUT_ENCODINGS[csv_file_type]
    else:
        with options if compact else st.container():
            output_format = st.selectbox(
                f"{bucket_label} output files", ("Excel (.xlsx)", "CSV (.csv)"), key=f"{state_prefix}output_format",
            )
    excel_output = output_format == "Excel (.xlsx)"
    if uploads is None:
        uploads = st.file_uploader(
            f"{bucket_label} output files" if b2 else "GEO & RECO output files", type=["xlsx", "csv"],
            accept_multiple_files=True, key=upload_widget_key(f"{state_prefix}uploads"),
            help="Upload the CSV or Excel files generated by the automation. All selected files are compiled into activity logs.",
        )
    fixed_timestamp = None
    default_collector = ""
    with options if compact else st.expander("Log settings"):
        st.session_state[f"{state_prefix}remark_by"] = AUTOSTAT_REMARK_BY
        st.text_input("Remarks By" if b2 else "Remark By", key=f"{state_prefix}remark_by", disabled=True)
        remark_by = AUTOSTAT_REMARK_BY
        if b2:
            default_collector = st.text_input(
                "Default collector (when missing in the source)", key=f"{state_prefix}collector",
            )
        blank_timestamp = st.checkbox("Leave date and time blank", key=f"{state_prefix}blank_time")
        current_time = st.checkbox(
            "Use current Manila date and time", value=True, key=f"{state_prefix}current_time", disabled=blank_timestamp,
        )
        if not blank_timestamp and not current_time:
            default = st.session_state.setdefault(f"{state_prefix}default_time", datetime.now(MANILA_TIMEZONE))
            left, right = st.columns(2)
            log_date = left.date_input("Log date (Asia/Manila)", default.date(), key=f"{state_prefix}date")
            log_time = right.time_input("Log time (Asia/Manila)", default.time().replace(microsecond=0), key=f"{state_prefix}time")
            fixed_timestamp = datetime.combine(log_date, log_time, tzinfo=MANILA_TIMEZONE)
    if not uploads:
        st.session_state.pop(result_key, None)
        return

    sources, configurations, problems, excluded_names = [], [], [], []
    cache = st.session_state.setdefault(f"{state_prefix}source_info", {})
    active_cache = {}
    source_columns_panel = None

    def inspect(data, extension, **kwargs):
        cache_key = (hashlib.sha256(data).hexdigest(), extension, tuple(kwargs.items()), AUTOSTAT_VERSION)
        if cache_key not in cache:
            cache[cache_key] = inspect_contact_source(data, extension, **kwargs)
        active_cache[cache_key] = cache[cache_key]
        return [b2_account_sheet(sheet) for sheet in cache[cache_key]] if b2 else cache[cache_key]

    for file_index, upload in enumerate(uploads):
        if not b2 and autostat_source_excluded(upload.name):
            excluded_names.append(upload.name)
            configurations.append((upload.name, "excluded"))
            continue
        data = upload.getvalue()
        extension = Path(upload.name).suffix.lower()
        digest = hashlib.sha256(data).hexdigest()
        prefix = f"{state_prefix}{AUTOSTAT_VERSION}_{file_index}_{digest}"
        if b2:
            prefix += f"_{B2_PRIMARY_KEY_VERSION}"
        configurations.append((upload.name, digest))
        try:
            sheets = inspect(data, extension)
            if not b2:
                excluded_names.extend(f"{upload.name} / {sheet.name}" for sheet in sheets
                                      if autostat_source_excluded(sheet.name))
                sheets = [sheet for sheet in sheets if not autostat_source_excluded(sheet.name)]
                if not sheets:
                    continue
            if source_columns_panel is None:
                source_columns_panel = options.expander("Edit source columns") if compact else st.expander("Edit source columns")
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
                    primary_key_account = (
                        b2 and sheet.account_column is not None
                        and normalize_header(sheet.headers[sheet.account_column]) == "primarykey"
                    )
                    account = st.selectbox(
                        "Account number column", account_options,
                        index=account_options.index(sheet.account_column), format_func=label,
                        key=f"{column_key}_account", disabled=primary_key_account,
                    )
                    if primary_key_account:
                        st.caption(f"primary_key supplies the account number in all three {bucket_label} logs.")
                    email_columns = st.multiselect(
                        "Email columns", columns, default=list(sheet.email_columns),
                        format_func=label, key=f"{column_key}_emails",
                    )
                    phone_columns = st.multiselect(
                        "Phone columns", columns, default=list(sheet.phone_columns),
                        format_func=label, key=f"{column_key}_phones",
                    )
                    configurations.append((sheet.name, header_row, account, tuple(email_columns), tuple(phone_columns)))
                    name_column, collector_column = None, None
                    if b2:
                        name_default, collector_default = b2_metadata_columns(sheet.headers)
                        name_column = st.selectbox(
                            "Name column", account_options, index=account_options.index(name_default),
                            format_func=label, key=f"{column_key}_name",
                        )
                        collector_column = st.selectbox(
                            "Collector column", account_options, index=account_options.index(collector_default),
                            format_func=label, key=f"{column_key}_collector_{B2_COLLECTOR_TERMS_VERSION}",
                        )
                        configurations.append((name_column, collector_column))
                    if account is None or not (email_columns or phone_columns):
                        problems.append(f"{upload.name} / {sheet.name}: select an account column and at least one contact column.")
                    elif account in (*email_columns, *phone_columns):
                        problems.append(f"{upload.name} / {sheet.name}: the account column must be separate from contacts.")
                    elif b2 and any(column is not None and column in (account, *email_columns, *phone_columns)
                                    for column in (name_column, collector_column)):
                        problems.append(f"{upload.name} / {sheet.name}: name and collector must be separate from the account and contacts.")
                    else:
                        source = SourceSelection(upload.name, data, replace(
                            sheet, account_column=account, email_columns=tuple(email_columns), phone_columns=tuple(phone_columns),
                        ))
                        sources.append(B2ActivitySource(source, name_column, collector_column) if b2 else source)
                unselected = [sheet.name for sheet in sheets if sheet.name not in selected]
                if unselected:
                    st.caption("Not processed: " + ", ".join(unselected))
                if not selected:
                    st.info("Select a sheet and its columns to include this file.")
        except (AutomationError, UnicodeError, ValueError) as exc:
            problems.append(f"{upload.name}: {exc}")
    st.session_state[f"{state_prefix}source_info"] = active_cache
    if excluded_names:
        with st.expander(f"Excluded files ({len(excluded_names)})"):
            for name in excluded_names:
                st.write(name)
            st.caption("Standard excludes SBC INSURANCE, B2 and B4 files and sheets. Use the matching bucket mode for B2/B4 outputs.")
    for problem in problems:
        st.error(problem)
    if not sources:
        st.info("Select at least one sheet with an account number column and an email or phone column.")
    signature = (
        AUTOSTAT_VERSION, COMPILED_OUTPUT_VERSION, AUTOSTAT_PAGE_VERSION, EMAIL_VALIDATION_VERSION, OB_SORT_VERSION,
        tuple(configurations), remark_by,
        blank_timestamp, current_time, fixed_timestamp, encoding, output_format, EMAIL_HEADERS, VIBER_HEADERS,
        b2, B2_AUTOSTAT_VERSION, B2_EMAIL_HEADERS, B2_PHONE_HEADERS, default_collector,
        bucket, B4_AUTOSTAT_VERSION, compact,
    )
    previous = st.session_state.get(result_key)
    if previous is not None and (previous["signature"] != signature or problems or not sources):
        st.session_state.pop(result_key, None)

    if sources and not problems and st.session_state.get(result_key) is None:
        st.session_state.pop(result_key, None)
        try:
            with st.spinner("Checking contacts and preparing AUTOSTAT logs..."):
                timestamp = None if blank_timestamp else fixed_timestamp or datetime.now(MANILA_TIMEZONE)
                if b2:
                    builder = build_b4_autostat_exports if bucket == 4 else build_b2_autostat_exports
                    bundle = builder(sources, timestamp=timestamp, remark_by=remark_by,
                                     default_collector=default_collector, encoding=encoding)
                else:
                    bundle = build_autostat_exports(sources, timestamp=timestamp, remark_by=remark_by, encoding=encoding)
                st.session_state[result_key] = {"signature": signature, "bundle": bundle, "timestamp": timestamp}
        except AutomationError as exc:
            st.error(str(exc))
        except Exception:
            st.error("The source files could not be processed. Check their selected sheets and columns, then try a fresh Excel or CSV copy.")
    result = st.session_state.get(result_key)
    if result is None:
        return
    bundle = result["bundle"]
    email_count = bundle.files[0].count
    phone_count = bundle.files[1].count
    if b2 and not compact:
        for column, file, channel in zip(st.columns(3), bundle.files, ("Email", "SMS", "Viber")):
            column.metric(f"{channel} accounts", f"{file.count:,}")
    elif not compact:
        email_column, phone_column = st.columns(2)
        email_column.metric("Valid email accounts", f"{email_count:,}")
        phone_column.metric("Valid phone accounts", f"{phone_count:,}")
    skipped_rows = sum(row["Source error rows skipped"] for row in bundle.summary)
    if skipped_rows:
        st.warning(
            f"Skipped {skipped_rows:,} database row{'s' if skipped_rows != 1 else ''} with scientific-notation account numbers, Excel cell errors, "
            "or formulas without saved values. Only valid rows are included in AUTOSTAT."
        )
    if any(row["Account precision issues"] for row in bundle.summary):
        st.warning(
            "Some source account numbers are numeric and may already be rounded. "
            "Upload the original source with full card numbers stored as text; missing digits cannot be recovered."
        )
    if not (email_count or phone_count):
        st.info("No valid email or mobile contacts with account numbers were found for AUTOSTAT. Check the selected sheets and columns.")
        return
    if not compact:
        selected_sheets = [source.source.sheet if b2 else source.sheet for source in sources]
        if not any(sheet.phone_columns for sheet in selected_sheets):
            st.warning(
                "No phone column is selected for Viber. Open Edit source columns and select the mobile-number column."
            )
        elif not phone_count:
            st.warning(
                "No valid mobile numbers with account numbers were found for Viber. Check the selected phone columns. "
                "Mobile numbers must normalize to 10 digits starting with 9."
            )
    render_autostat_downloads(
        bundle, state_prefix=state_prefix, workbook_name=workbook_name, excel_output=excel_output,
        show_counts=False, compact=compact,
    )
    if not compact:
        st.caption(
            f"Open {workbook_name}.xlsx in Excel to keep account numbers as text. Whole-number accounts have no .0 suffix. "
            "When copying elsewhere, format the destination account column as Text before pasting."
        )
    if not excel_output and not compact:
        st.caption("Opening CSV files directly in Excel can round account digits to zeros. Use the Excel download for viewing and editing.")
    with options.expander("Save to a folder") if compact else st.expander("Save to a folder"):
        output_folder = st.text_input("Output folder", value=OUTPUT_FOLDER, key=f"{state_prefix}output_folder")
        if st.button("Save AUTOSTAT files to folder", key=f"{state_prefix}save"):
            try:
                folder = save_files_to_folder(
                    [(Path(file.filename).with_suffix(".xlsx").name, build_autostat_channel_excel(file))
                     if excel_output else (file.filename, file.data) for file in bundle.files],
                    output_folder, batch_name=workbook_name,
                )
                result["saved_folder"] = str(folder)
                st.session_state[result_key] = result
            except AutomationError as exc:
                st.error(str(exc))
        if result.get("saved_folder"):
            st.success(f"Saved {len(bundle.files)} files:")
            st.code(result["saved_folder"], language=None)
    with options.expander("Preview") if compact else st.expander("Preview"):
        filename = st.selectbox(
            "File to preview", [file.filename for file in bundle.files],
            key=f"{state_prefix}preview_{hash(signature)}",
        )
        file = next(file for file in bundle.files if file.filename == filename)
        st.dataframe(pd.DataFrame(file.preview, columns=file.headers), hide_index=True, width="stretch")
        st.caption(f"Showing the first {min(file.count, 100):,} of {file.count:,} rows.")
