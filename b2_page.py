"""Prepare B2 and B4 blasting templates from account lists and SBC contacts."""

from __future__ import annotations

import hashlib

import pandas as pd
import streamlit as st

from automation import AutomationError, CSV_OUTPUT_ENCODINGS, MAPPING_VERSION, WORKBOOK_READER_VERSION
from autostat import AUTOSTAT_VERSION, SourceSelection, inspect_contact_source
from b2_matching import (
    B2_MATCHING_VERSION, B2_TEMPLATE_HEADERS, B2_TEMPLATE_VERSION, B2_B4_AGENTS_VERSION,
    B4_MATCHING_VERSION, build_b2_exports, build_b4_exports,
    is_datagrid_sheet, is_sbc_database_sheet, matching_sheet_score, validate_b2_source,
)
from export_config import AUTOSTAT_B2_FILENAMES, AUTOSTAT_B4_FILENAMES, OUTPUT_FOLDER
from folder_output import save_files_to_folder
from upload_state import upload_widget_key


B4_PAGE_VERSION = 1


def render_b2_page(*, datagrid_upload=None, encoding: str | None = None, bucket: int = 2) -> None:
    bucket_label = f"B{bucket}"
    prefix = f"b{bucket}_"
    result_key = prefix + "result"
    filenames = AUTOSTAT_B2_FILENAMES if bucket == 2 else AUTOSTAT_B4_FILENAMES
    st.subheader(f"NEGATIVE {bucket_label} — match accounts and contacts")
    st.write("Use DataGrid as the account list and the full SBC database as the source of contact details.")
    st.caption(
        "Upload the latest SBC database each day. Replacing either workbook clears previous downloads, "
        "even when the filename is unchanged."
    )
    st.caption(
        "Account No. matches Account Key; Financing/Card No. matches Contract Number. "
        "When both identifiers are present, both must match the same database record. "
        f"Only verified {bucket_label} accounts are included. primary_key uses Account Key; cardnumber uses Contract Number."
    )
    with st.expander(f"{bucket_label} blasting template"):
        st.write("Using the 16 columns from 1EMAIL_Blasting - Template.xlsx.")
        st.dataframe(pd.DataFrame(columns=B2_TEMPLATE_HEADERS), hide_index=True, width="stretch")
        st.caption(
            "One row per account, with email and phone in separate columns. Monthly contains the agent name; Terms contains the agent code. "
            "Old I.C. supplies indexoneuid; the database supplies available name, outstanding balance, assignment date and Area. "
            f"Product is CARDS or BEL; doclass is SBC CARDS {bucket_label} or SBC BEL {bucket_label}. Other missing values stay blank."
        )
    if encoding is None:
        with st.sidebar:
            output_format = st.selectbox("CSV file type", list(CSV_OUTPUT_ENCODINGS), key=f"{prefix}csv_type")
        encoding = CSV_OUTPUT_ENCODINGS[output_format]
    left, right = st.columns(2)
    with left:
        database_upload = st.file_uploader(
            "1. Latest full SBC database" if bucket == 2 else "1. B4 endorsement database", type=["xlsx"],
            key=upload_widget_key(f"{prefix}database_upload"),
            help="Use the database with Account Key, Contract Number, Bucket Code, EMAIL (01) and Mobile 1."
            if bucket == 2 else "Use the B4 endorsement workbook with Account Key, Contract Number, Bucket Code, Email (01) and Phone (01).",
        )
    with right:
        if datagrid_upload is None:
            grid_upload = st.file_uploader(
                "2. DataGrid account list" if bucket == 2 else "2. B4 account list", type=["xlsx"],
                key=upload_widget_key(f"{prefix}grid_upload"),
                help="Use your BP or CARDS AND BEL account list, with Account No. and Financing/Card No."
                if bucket == 2 else "Use NEGATIVE B4 DB.xlsx, with Account No. and Financing/Card No.",
            )
        else:
            grid_upload = datagrid_upload
            st.write(f"DataGrid account list: {grid_upload.name}")
    cache = st.session_state.setdefault(f"{prefix}source_info", {})

    def inspect_upload(upload):
        if upload is None:
            return None
        data = upload.getvalue()
        digest = hashlib.sha256(data).hexdigest()
        key = (digest, AUTOSTAT_VERSION, MAPPING_VERSION, WORKBOOK_READER_VERSION)
        if key not in cache:
            cache[key] = inspect_contact_source(data, ".xlsx")
        return data, digest, cache[key]

    try:
        database_info = inspect_upload(database_upload)
        grid_info = inspect_upload(grid_upload)
    except AutomationError as exc:
        st.session_state.pop(result_key, None)
        st.error(str(exc))
        return
    if (database_info is not None and grid_info is not None
            and not any(is_sbc_database_sheet(sheet.headers) for sheet in database_info[2])
            and any(is_datagrid_sheet(sheet.headers) for sheet in database_info[2])
            and any(is_sbc_database_sheet(sheet.headers) for sheet in grid_info[2])):
        database_upload, grid_upload = grid_upload, database_upload
        database_info, grid_info = grid_info, database_info
        st.info(f"Detected the SBC database in {database_upload.name} and the account list in {grid_upload.name}.")
    if database_upload is None or grid_upload is None:
        st.session_state.pop(result_key, None)
        available = database_info or grid_info
        if available is not None and any(is_sbc_database_sheet(sheet.headers) for sheet in available[2]):
            st.info("The uploaded file is the SBC database. Upload the DataGrid account list in the empty box.")
            return
        if available is not None and any(is_datagrid_sheet(sheet.headers) for sheet in available[2]):
            st.info("The uploaded file is the DataGrid account list. Upload the latest full SBC database in the empty box.")
            return
        st.info(
            f"Upload the latest full SBC database to prepare NEGATIVE {bucket_label} - CARDS.csv and NEGATIVE {bucket_label} - BEL.csv."
            if grid_upload is not None else
            f"Upload both workbooks to prepare NEGATIVE {bucket_label} - CARDS.csv and NEGATIVE {bucket_label} - BEL.csv."
        )
        return
    if all(not any(is_sbc_database_sheet(sheet.headers) for sheet in info[2])
           and any(is_datagrid_sheet(sheet.headers) for sheet in info[2]) for info in (database_info, grid_info)):
        st.session_state.pop(result_key, None)
        st.error(
            f"{database_upload.name} and {grid_upload.name} are both DataGrid account lists. "
            "Keep the account list you want to process and replace the other upload with the latest full SBC database, sheet SBC."
        )
        return
    sources, signatures = [], []
    for upload, info, database, label in (
        (database_upload, database_info, True, "SBC database" if bucket == 2 else "B4 endorsement"),
        (grid_upload, grid_info, False, "DataGrid accounts" if bucket == 2 else "B4 accounts"),
    ):
        data, digest, sheets = info
        try:
            best = max(sheets, key=lambda sheet: matching_sheet_score(sheet, database=database))
            names = [sheet.name for sheet in sheets]
            selected = st.selectbox(
                f"{label} sheet", names, index=names.index(best.name), key=f"{prefix}sheet_{database}_{digest}",
            )
            sheet = next(sheet for sheet in sheets if sheet.name == selected)
            header = int(st.number_input(
                f"{label} header row", min_value=1, max_value=1_048_575, value=sheet.header_row,
                key=f"{prefix}header_{database}_{digest}_{selected}",
            ))
            if header != sheet.header_row:
                sheet = inspect_contact_source(data, ".xlsx", sheet_name=selected, header_row=header)[0]
            st.caption(f"{label}: {upload.name} — {len(sheet.headers)} source columns found.")
            source = SourceSelection(upload.name, data, sheet)
            validate_b2_source(source, database=database, bucket=bucket)
            sources.append(source)
            signatures.append((upload.name, digest, selected, header))
        except AutomationError as exc:
            st.session_state.pop(result_key, None)
            st.error(str(exc))
            return
    st.caption("Each output uses the 16-column blasting template, with one row per account. " + (
        "EMAIL (01) and Mobile 1 are checked first. Other contact fields are fallbacks when needed. " if bucket == 2 else
        "Email (01) supplies email; Phone (01) supplies phone. Other contact fields are not substituted. "
    ) + "Phones use the rightmost 10 digits starting with 9.")
    st.caption(
        "Each CARDS/BEL file is split evenly by alternating accounts: "
        "Almera Ranceli / AMRANCELI, then Kyarill Palima / KPALIMA. "
        "An odd extra account goes to Almera. Monthly and Terms stay paired on every row."
    )
    signature = (tuple(signatures), encoding, B2_MATCHING_VERSION, B2_TEMPLATE_VERSION, B2_TEMPLATE_HEADERS,
                 AUTOSTAT_VERSION, MAPPING_VERSION, WORKBOOK_READER_VERSION, bucket, B4_MATCHING_VERSION,
                 B2_B4_AGENTS_VERSION, tuple(filenames.items()))
    previous = st.session_state.get(result_key)
    if previous is not None and previous["signature"] != signature:
        st.session_state.pop(result_key, None)
    if st.button(f"Match and prepare {bucket_label} files", type="primary", key=f"{prefix}prepare"):
        st.session_state.pop(result_key, None)
        try:
            with st.spinner("Matching DataGrid accounts and checking SBC contacts..."):
                builder = build_b2_exports if bucket == 2 else build_b4_exports
                result = builder(sources[0], sources[1], encoding=encoding)
                st.session_state[result_key] = {"signature": signature, "matched": result}
        except AutomationError as exc:
            st.error(str(exc))
        except Exception:
            st.error("The matching files could not be processed. Check the selected database and DataGrid sheets and header rows.")
    prepared = st.session_state.get(result_key)
    if prepared is None:
        return
    result = prepared["matched"]
    bundle = result.bundle
    email_count = sum(row["Email rows"] for row in bundle.summary)
    phone_count = sum(row["Viber rows"] for row in bundle.summary)
    account_count = sum(file.count for file in bundle.files)
    if email_count or phone_count:
        st.success(f"Ready: {account_count:,} accounts in two {bucket_label} blasting files, with {email_count:,} emails and {phone_count:,} phones.")
    else:
        st.warning(f"No verified {bucket_label} accounts with valid contacts were found. Both files contain template headers only.")
    st.dataframe(pd.DataFrame(bundle.summary), hide_index=True, width="stretch")
    with st.expander("Matching report", expanded=bool(result.audit)):
        st.dataframe(pd.DataFrame(result.stats.items(), columns=["Check", "Count"]), hide_index=True, width="stretch")
        if result.audit:
            st.warning(f"{len(result.audit):,} DataGrid row(s) were excluded. Review their account identifiers and reasons below.")
            st.dataframe(pd.DataFrame(result.audit), hide_index=True, width="stretch")
    st.subheader("Save to a folder")
    output_folder = st.text_input("Output folder", value=OUTPUT_FOLDER, key=f"{prefix}output_folder")
    st.caption("Saves on the computer running this app. Use downloads to save files on your PC when the app is hosted online.")
    if st.button(f"Save {bucket_label} files to folder", key=f"{prefix}save"):
        try:
            folder = save_files_to_folder(
                [(file.filename, file.data) for file in bundle.files], output_folder, batch_name=f"NEGATIVE {bucket_label}",
            )
            prepared["saved_folder"] = str(folder)
            st.session_state[result_key] = prepared
        except AutomationError as exc:
            st.error(str(exc))
    if prepared.get("saved_folder"):
        st.success(f"Saved both {bucket_label} files:")
        st.code(prepared["saved_folder"], language=None)
    st.download_button(f"Download {bucket_label} files (ZIP)", bundle.zip_data, f"NEGATIVE {bucket_label}.zip",
                       mime="application/zip", key=f"{prefix}zip", on_click="ignore")
    for index, file in enumerate(bundle.files):
        st.download_button(f"Download {file.filename}", file.data, file.filename, mime="text/csv",
                           key=f"{prefix}download_{index}", on_click="ignore")
    selected = st.selectbox(f"{bucket_label} file to preview", [file.filename for file in bundle.files], key=f"{prefix}preview_{hash(signature)}")
    file = next(file for file in bundle.files if file.filename == selected)
    st.dataframe(pd.DataFrame(file.preview, columns=file.headers), hide_index=True, width="stretch")
    st.caption(f"Showing the first {min(file.count, 100):,} of {file.count:,} accounts. Preparing files does not send messages.")
