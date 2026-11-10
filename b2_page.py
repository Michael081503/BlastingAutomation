"""Upload B2/B4 workbooks and automatically prepare their downloads."""

from __future__ import annotations

import hashlib
from datetime import datetime
from functools import partial

import pandas as pd
import streamlit as st

from automation import AutomationError, CSV_OUTPUT_ENCODINGS, MAPPING_VERSION, OB_SORT_VERSION, WORKBOOK_READER_VERSION
from autostat import AUTOSTAT_VERSION, SourceSelection, inspect_contact_source
from agent_availability_page import render_agent_availability
from b2_matching import (
    B2_MATCHING_VERSION, B2_TEMPLATE_HEADERS, B2_TEMPLATE_VERSION, B2_B4_AGENTS_VERSION, B2_B4_OB_SORT_VERSION,
    B4_MATCHING_VERSION, B2_B4_CONTACT_HEADERS_VERSION, B2_B4_AVAILABILITY_VERSION, build_b2_exports, build_b4_exports,
    SPM_ENDORSEMENT_VERSION,
    is_datagrid_sheet, is_sbc_database_sheet, matching_sheet_score, validate_b2_source,
)
from export_config import AUTOSTAT_B2_FILENAMES, AUTOSTAT_B4_FILENAMES, COLLECTION_AGENTS
from exports import EXCEL_SAFE_CSV_VERSION, build_csv_zip, build_export_csv
from folder_output import MANILA_TIMEZONE
from upload_state import upload_widget_key


B4_PAGE_VERSION = 2
B2_COMPACT_DOWNLOADS_VERSION = 1
ZIP_FORMAT_VERSION = 1
BROWSER_DOWNLOADS_ONLY_VERSION = 1
B2_B4_CSV_PROTECTION_VERSION = 1


def no_matching_rows_message(result, *, bucket: int, database_only: bool = False) -> str:
    other = 4 if bucket == 2 else 2
    stats = result.stats
    if not stats.get(f"Database B{bucket} rows", 0):
        other_count = stats.get(f"Database B{other} rows", 0)
        detail = f"contains {other_count:,} B{other} accounts and no B{bucket} accounts" if other_count else f"has no B{bucket} accounts"
        return f"The uploaded database {detail}. Upload the matching B{bucket} database."
    if not stats["Matched rows"]:
        if database_only:
            return f"No unambiguous B{bucket} accounts with both Account Key and Contract Number were found in this database. Check Details."
        return f"No account numbers from this B{bucket} list match the uploaded database. Upload the corresponding B{bucket} database."
    if stats["Other or missing bucket"] == stats["Matched rows"]:
        return f"The matched accounts are outside B{bucket} in this database. Upload the current B{bucket} database."
    if stats["Unsupported product"]:
        return f"No supported B{bucket} CARDS or BEL accounts with valid contacts were found. Check Details."
    contacts = "primary email or Mobile 1 contacts" if bucket == 4 else "email or mobile contacts"
    return f"No matched B{bucket} accounts have valid {contacts}. Check the contact database."


def render_b2_page(*, datagrid_upload=None, encoding: str | None = None, bucket: int = 2) -> None:
    bucket_label = f"B{bucket}"
    prefix = f"b{bucket}_"
    result_key = prefix + "result"
    filenames = AUTOSTAT_B2_FILENAMES if bucket == 2 else AUTOSTAT_B4_FILENAMES
    st.subheader(f"NEGATIVE {bucket_label}")
    if encoding is None:
        with st.sidebar:
            output_format = st.selectbox("CSV file type", list(CSV_OUTPUT_ENCODINGS), key=f"{prefix}csv_type")
        encoding = CSV_OUTPUT_ENCODINGS[output_format]
    assignment_date = datetime.now(MANILA_TIMEZONE).date()
    with st.expander(f"{bucket_label} agent availability"):
        agents = render_agent_availability("b2_b4", assignment_date, COLLECTION_AGENTS, key_prefix=prefix)
    database_only = bucket == 4 and st.checkbox(
        "Use all B4 accounts from this database", key="b4_use_database_accounts",
        help="Use the SPM endorsement alone. Include bucket 4 accounts with a valid primary email or Mobile 1.",
    )
    left, right = st.columns(2)
    with left:
        database_upload = st.file_uploader(
            "1. Latest full SBC database" if bucket == 2 else "1. SPM endorsement / B4 database", type=["xlsx"],
            key=upload_widget_key(f"{prefix}database_upload"),
            help=((f"Use NEGATIVE B{bucket} DATABASE.xlsx or the full contact database for this bucket. " if bucket == 2
                   else "Use SPM ENDORSEMENT B4 or NEGATIVE B4 DATABASE.xlsx. ")
                  + "Contacts come from Email (1) / Email (01) and Mobile 1."),
        )
    with right:
        grid_upload = datagrid_upload if datagrid_upload is not None else st.file_uploader(
            "2. DataGrid account list" if bucket == 2 else "2. B4 account list", type=["xlsx"],
            key=upload_widget_key(f"{prefix}grid_upload"),
            disabled=database_only,
            help="Use the account list with Account No. and Financing/Card No.; for B4 this is NEGATIVE B4 DB.xlsx.",
        )
    if agents is None:
        st.session_state.pop(result_key, None)
        return
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
        grid_info = None if database_only else inspect_upload(grid_upload)
    except AutomationError as exc:
        st.session_state.pop(result_key, None)
        st.error(str(exc))
        return
    role_message = None
    if (database_info is not None and grid_info is not None
            and not any(is_sbc_database_sheet(sheet.headers) for sheet in database_info[2])
            and any(is_datagrid_sheet(sheet.headers) for sheet in database_info[2])
            and any(is_sbc_database_sheet(sheet.headers) for sheet in grid_info[2])):
        database_upload, grid_upload = grid_upload, database_upload
        database_info, grid_info = grid_info, database_info
        role_message = f"Detected the SBC database in {database_upload.name} and the account list in {grid_upload.name}."
    if database_upload is None or (grid_upload is None and not database_only):
        st.session_state.pop(result_key, None)
        available = database_info or grid_info
        if available is not None:
            if any(is_sbc_database_sheet(sheet.headers) for sheet in available[2]):
                st.info("The uploaded file is the SBC database. Upload the DataGrid account list in the empty box." if bucket == 2
                        else "Upload a B4 account list, or select Use all B4 accounts from this database.")
            elif any(is_datagrid_sheet(sheet.headers) for sheet in available[2]):
                st.info("The uploaded file is the DataGrid account list. Upload the latest full SBC database in the empty box.")
            else:
                st.info("Upload both the contact database and account list.")
        return
    if not database_only and all(not any(is_sbc_database_sheet(sheet.headers) for sheet in info[2])
           and any(is_datagrid_sheet(sheet.headers) for sheet in info[2]) for info in (database_info, grid_info)):
        st.session_state.pop(result_key, None)
        st.error(
            f"{database_upload.name} and {grid_upload.name} are both DataGrid account lists. "
            "Replace one with the latest full SBC database, sheet SBC."
        )
        return

    sources, signatures = [], []
    selection_error = None
    with st.expander(f"{bucket_label} options"):
        if role_message:
            st.info(role_message)
        selections = [
            (database_upload, database_info, True, "SBC database" if bucket == 2 else "B4 endorsement"),
        ]
        if not database_only:
            selections.append((grid_upload, grid_info, False, "DataGrid accounts" if bucket == 2 else "B4 accounts"))
        for upload, info, database, label in selections:
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
                source = SourceSelection(upload.name, data, sheet)
                validate_b2_source(source, database=database, bucket=bucket)
                sources.append(source)
                signatures.append((upload.name, digest, selected, header))
            except AutomationError as exc:
                selection_error = str(exc)
                break
    if selection_error:
        st.session_state.pop(result_key, None)
        st.error(selection_error)
        return

    signature = (tuple(signatures), encoding, B2_MATCHING_VERSION, B2_TEMPLATE_VERSION, B2_TEMPLATE_HEADERS,
                 AUTOSTAT_VERSION, MAPPING_VERSION, WORKBOOK_READER_VERSION, bucket, B4_MATCHING_VERSION,
                 B2_B4_AGENTS_VERSION, B2_B4_OB_SORT_VERSION, OB_SORT_VERSION,
                 tuple(filenames.items()), B2_COMPACT_DOWNLOADS_VERSION, B2_B4_CONTACT_HEADERS_VERSION,
                 assignment_date.isoformat(), agents, B2_B4_AVAILABILITY_VERSION,
                 B2_B4_CSV_PROTECTION_VERSION, EXCEL_SAFE_CSV_VERSION, SPM_ENDORSEMENT_VERSION, database_only)
    prepared = st.session_state.get(result_key)
    if prepared is None or prepared["signature"] != signature:
        st.session_state.pop(result_key, None)
        prepared = {"signature": signature}
        try:
            with st.spinner(f"Preparing {bucket_label} downloads..."):
                builder = build_b2_exports if bucket == 2 else build_b4_exports
                prepared["matched"] = builder(sources[0], None if database_only else sources[1], encoding=encoding, agents=agents)
        except AutomationError as exc:
            prepared["error"] = str(exc)
        except Exception:
            prepared["error"] = "The workbooks could not be processed. Check the selected sheets and header rows in Options."
        st.session_state[result_key] = prepared
    if "error" in prepared:
        st.error(prepared["error"])
        return
    result = prepared["matched"]
    bundle = result.bundle
    nonempty = [file for file in bundle.files if file.count]
    if not nonempty:
        st.error(no_matching_rows_message(result, bucket=bucket, database_only=database_only))
    else:
        if len(nonempty) == len(bundle.files):
            st.download_button(f"Download {bucket_label} CSV ZIP (.csv files)", partial(build_csv_zip, bundle.files), f"NEGATIVE {bucket_label}_CSV.zip",
                               mime="application/zip", key=f"{prefix}zip", on_click="ignore")
        for index, file in enumerate(bundle.files):
            if file.count:
                st.download_button(f"Download {file.filename}", partial(build_export_csv, file), file.filename, mime="text/csv",
                                   key=f"{prefix}download_{index}", on_click="ignore")
    with st.container(key=f"{prefix}details_section"):
        st.subheader(f"{bucket_label} details")
        st.dataframe(pd.DataFrame(bundle.summary), hide_index=True, width="stretch")
        st.dataframe(pd.DataFrame(result.stats.items(), columns=["Check", "Count"]), hide_index=True, width="stretch")
        if result.audit:
            st.dataframe(pd.DataFrame(result.audit), hide_index=True, width="stretch")
        selected = st.selectbox(f"{bucket_label} file to preview", [file.filename for file in bundle.files], key=f"{prefix}preview_{hash(signature)}")
        file = next(file for file in bundle.files if file.filename == selected)
        st.dataframe(pd.DataFrame(file.preview, columns=file.headers), hide_index=True, width="stretch")
