"""AUTOSTAT interface for matching the DataGrid B2 account list to SBC contacts."""

from __future__ import annotations

import hashlib
from datetime import datetime

import pandas as pd
import streamlit as st

from automation import AutomationError, CSV_OUTPUT_ENCODINGS, MAPPING_VERSION, WORKBOOK_READER_VERSION
from autostat import AUTOSTAT_VERSION, SourceSelection, inspect_contact_source
from b2_matching import B2_MATCHING_VERSION, build_b2_exports, matching_sheet_score
from export_config import AUTOSTAT_B2_FILENAMES, OUTPUT_FOLDER
from folder_output import MANILA_TIMEZONE, save_files_to_folder


def render_b2_page(*, datagrid_upload=None, encoding: str | None = None) -> None:
    st.subheader("NEGATIVE B2 — match accounts and contacts")
    st.write("Use DataGrid as the account list and the full SBC database as the source of contact details.")
    st.caption(
        "Upload the latest SBC database each day. Replacing either workbook clears previous downloads, "
        "even when the filename is unchanged."
    )
    st.caption(
        "Account No. matches Account Key; Financing/Card No. matches Contract Number. "
        "When both identifiers are present, both must match the same database record. "
        "Only verified B2 accounts are included. Account No. in the logs uses Contract Number."
    )
    if encoding is None:
        with st.sidebar:
            output_format = st.selectbox("CSV file type", list(CSV_OUTPUT_ENCODINGS), key="b2_csv_type")
        encoding = CSV_OUTPUT_ENCODINGS[output_format]
    left, right = st.columns(2)
    with left:
        database_upload = st.file_uploader("1. Latest full SBC database", type=["xlsx"], key="b2_database_upload")
    with right:
        if datagrid_upload is None:
            grid_upload = st.file_uploader("2. DataGrid account list", type=["xlsx"], key="b2_grid_upload")
        else:
            grid_upload = datagrid_upload
            st.write(f"DataGrid account list: {grid_upload.name}")
    if database_upload is None or grid_upload is None:
        st.session_state.pop("b2_result", None)
        st.info(
            "Upload the latest full SBC database to prepare NEGATIVE B2 - CARDS.csv and NEGATIVE B2 - BEL.csv."
            if grid_upload is not None else
            "Upload both workbooks to prepare NEGATIVE B2 - CARDS.csv and NEGATIVE B2 - BEL.csv."
        )
        return
    sources, signatures = [], []
    cache = st.session_state.setdefault("b2_source_info", {})
    for upload, database, label in ((database_upload, True, "SBC database"), (grid_upload, False, "DataGrid accounts")):
        data = upload.getvalue()
        digest = hashlib.sha256(data).hexdigest()
        try:
            key = (digest, AUTOSTAT_VERSION, MAPPING_VERSION, WORKBOOK_READER_VERSION)
            if key not in cache:
                cache[key] = inspect_contact_source(data, ".xlsx")
            sheets = cache[key]
            best = max(sheets, key=lambda sheet: matching_sheet_score(sheet, database=database))
            names = [sheet.name for sheet in sheets]
            selected = st.selectbox(
                f"{label} sheet", names, index=names.index(best.name), key=f"b2_sheet_{database}_{digest}",
            )
            sheet = next(sheet for sheet in sheets if sheet.name == selected)
            header = int(st.number_input(
                f"{label} header row", min_value=1, max_value=1_048_575, value=sheet.header_row,
                key=f"b2_header_{database}_{digest}_{selected}",
            ))
            if header != sheet.header_row:
                sheet = inspect_contact_source(data, ".xlsx", sheet_name=selected, header_row=header)[0]
            st.caption(f"{label}: {upload.name} — {len(sheet.headers)} source columns found.")
            sources.append(SourceSelection(upload.name, data, sheet))
            signatures.append((upload.name, digest, selected, header))
        except AutomationError as exc:
            st.session_state.pop("b2_result", None)
            st.error(str(exc))
            return
    remark_by = st.text_input("Remark By", value="MBURGOS", key="b2_remark_by")
    current_time = st.checkbox("Use current Manila date and time", value=True, key="b2_current_time")
    fixed_time = None
    if not current_time:
        default = st.session_state.setdefault("b2_default_time", datetime.now(MANILA_TIMEZONE))
        left, right = st.columns(2)
        day = left.date_input("Log date (Asia/Manila)", default.date(), key="b2_date")
        time = right.time_input("Log time (Asia/Manila)", default.time().replace(microsecond=0), key="b2_time")
        fixed_time = datetime.combine(day, time, tzinfo=MANILA_TIMEZONE)
    st.caption(
        "Each output contains Email and Viber activity rows using the existing template columns, statuses and remarks. "
        "The first valid email and first valid phone are kept per account. Phones use the rightmost 10 digits starting with 9."
    )
    signature = (tuple(signatures), remark_by, fixed_time, encoding, B2_MATCHING_VERSION,
                 AUTOSTAT_VERSION, MAPPING_VERSION, tuple(AUTOSTAT_B2_FILENAMES.items()))
    previous = st.session_state.get("b2_result")
    if previous is not None and previous["signature"] != signature:
        st.session_state.pop("b2_result", None)
    if st.button("Match and prepare B2 files", type="primary", key="b2_prepare"):
        st.session_state.pop("b2_result", None)
        try:
            with st.spinner("Matching DataGrid accounts and checking SBC contacts..."):
                timestamp = fixed_time or datetime.now(MANILA_TIMEZONE)
                result = build_b2_exports(sources[0], sources[1], timestamp=timestamp, remark_by=remark_by, encoding=encoding)
                st.session_state["b2_result"] = {"signature": signature, "matched": result, "timestamp": timestamp}
        except AutomationError as exc:
            st.error(str(exc))
        except Exception:
            st.error("The matching files could not be processed. Check the selected database and DataGrid sheets and header rows.")
    prepared = st.session_state.get("b2_result")
    if prepared is None:
        return
    result = prepared["matched"]
    bundle = result.bundle
    email_count = sum(row["Email rows"] for row in bundle.summary)
    phone_count = sum(row["Viber rows"] for row in bundle.summary)
    if email_count or phone_count:
        st.success(f"Ready: {email_count:,} Email entries and {phone_count:,} Viber entries in two B2 files.")
    else:
        st.warning("No verified B2 accounts with valid contacts were found. Both files contain template headers only.")
    st.caption(f"Log time (Asia/Manila): {prepared['timestamp'].strftime('%Y-%m-%d %I:%M:%S %p')}")
    st.dataframe(pd.DataFrame(bundle.summary), hide_index=True, width="stretch")
    with st.expander("Matching report", expanded=bool(result.audit)):
        st.dataframe(pd.DataFrame(result.stats.items(), columns=["Check", "Count"]), hide_index=True, width="stretch")
        if result.audit:
            st.warning(f"{len(result.audit):,} DataGrid row(s) were excluded. Review their account identifiers and reasons below.")
            st.dataframe(pd.DataFrame(result.audit), hide_index=True, width="stretch")
    st.subheader("Save to a folder")
    output_folder = st.text_input("Output folder", value=OUTPUT_FOLDER, key="b2_output_folder")
    st.caption("Saves on the computer running this app. Use downloads to save files on your PC when the app is hosted online.")
    if st.button("Save B2 files to folder", key="b2_save"):
        try:
            folder = save_files_to_folder(
                [(file.filename, file.data) for file in bundle.files], output_folder, batch_name="NEGATIVE B2",
            )
            prepared["saved_folder"] = str(folder)
            st.session_state["b2_result"] = prepared
        except AutomationError as exc:
            st.error(str(exc))
    if prepared.get("saved_folder"):
        st.success("Saved both B2 files:")
        st.code(prepared["saved_folder"], language=None)
    st.download_button("Download B2 files (ZIP)", bundle.zip_data, "NEGATIVE B2 AUTOSTAT.zip",
                       mime="application/zip", key="b2_zip", on_click="ignore")
    for index, file in enumerate(bundle.files):
        st.download_button(f"Download {file.filename}", file.data, file.filename, mime="text/csv",
                           key=f"b2_download_{index}", on_click="ignore")
    selected = st.selectbox("B2 file to preview", [file.filename for file in bundle.files], key=f"b2_preview_{hash(signature)}")
    file = next(file for file in bundle.files if file.filename == selected)
    st.dataframe(pd.DataFrame(file.preview, columns=file.headers), hide_index=True, width="stretch")
    st.caption(f"Showing the first {min(file.count, 100):,} of {file.count:,} entries. Preparing logs does not send messages.")
