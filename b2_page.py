"""Prepare B2 blasting templates from DataGrid accounts and current SBC contacts."""

from __future__ import annotations

import hashlib

import pandas as pd
import streamlit as st

from automation import AutomationError, CSV_OUTPUT_ENCODINGS, MAPPING_VERSION, WORKBOOK_READER_VERSION
from autostat import AUTOSTAT_VERSION, SourceSelection, inspect_contact_source
from b2_matching import B2_MATCHING_VERSION, B2_TEMPLATE_HEADERS, B2_TEMPLATE_VERSION, build_b2_exports, matching_sheet_score
from export_config import AUTOSTAT_B2_FILENAMES, OUTPUT_FOLDER
from folder_output import save_files_to_folder


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
        "Only verified B2 accounts are included. primary_key uses Account Key; cardnumber uses Contract Number."
    )
    with st.expander("B2 blasting template"):
        st.write("Using the 16 columns from 1EMAIL_Blasting - Template.xlsx.")
        st.dataframe(pd.DataFrame(columns=B2_TEMPLATE_HEADERS), hide_index=True, width="stretch")
        st.caption(
            "One row per account, with email and phone in separate columns. Monthly and Terms stay blank. "
            "Old I.C. supplies indexoneuid; the database supplies available name, outstanding balance, assignment date and Area. "
            "Product is CARDS or BEL; doclass is SBC CARDS B2 or SBC BEL B2. Other missing values stay blank."
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
    st.caption(
        "Each output uses the 16-column blasting template, with one row per account. "
        "The first valid email and first valid phone are kept per account. Phones use the rightmost 10 digits starting with 9."
    )
    signature = (tuple(signatures), encoding, B2_MATCHING_VERSION, B2_TEMPLATE_VERSION, B2_TEMPLATE_HEADERS,
                 AUTOSTAT_VERSION, MAPPING_VERSION, WORKBOOK_READER_VERSION, tuple(AUTOSTAT_B2_FILENAMES.items()))
    previous = st.session_state.get("b2_result")
    if previous is not None and previous["signature"] != signature:
        st.session_state.pop("b2_result", None)
    if st.button("Match and prepare B2 files", type="primary", key="b2_prepare"):
        st.session_state.pop("b2_result", None)
        try:
            with st.spinner("Matching DataGrid accounts and checking SBC contacts..."):
                result = build_b2_exports(sources[0], sources[1], encoding=encoding)
                st.session_state["b2_result"] = {"signature": signature, "matched": result}
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
    account_count = sum(file.count for file in bundle.files)
    if email_count or phone_count:
        st.success(f"Ready: {account_count:,} accounts in two B2 blasting files, with {email_count:,} emails and {phone_count:,} phones.")
    else:
        st.warning("No verified B2 accounts with valid contacts were found. Both files contain template headers only.")
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
    st.download_button("Download B2 files (ZIP)", bundle.zip_data, "NEGATIVE B2.zip",
                       mime="application/zip", key="b2_zip", on_click="ignore")
    for index, file in enumerate(bundle.files):
        st.download_button(f"Download {file.filename}", file.data, file.filename, mime="text/csv",
                           key=f"b2_download_{index}", on_click="ignore")
    selected = st.selectbox("B2 file to preview", [file.filename for file in bundle.files], key=f"b2_preview_{hash(signature)}")
    file = next(file for file in bundle.files if file.filename == selected)
    st.dataframe(pd.DataFrame(file.preview, columns=file.headers), hide_index=True, width="stretch")
    st.caption(f"Showing the first {min(file.count, 100):,} of {file.count:,} accounts. Preparing files does not send messages.")
