"""SBC INSURANCE status selection, automatic cleaning and campaign downloads."""

from __future__ import annotations

import hashlib
from functools import partial
from pathlib import Path

import pandas as pd
import streamlit as st

from automation import (
    AutomationError, CSV_OUTPUT_ENCODINGS, MAPPING_VERSION, WORKBOOK_READER_VERSION,
    XLSX_MIME, inspect_workbook, template_column_positions,
)
from exports import EXCEL_SAFE_CSV_VERSION, build_export_csv, build_export_excel
from sbc_insurance import (
    SBC_INSURANCE_VERSION, SBC_INSURANCE_MASTERLIST_VERSION, build_filtered_datagrid,
    inspect_insurance_masterlist, prepare_insurance, read_insurance_masterlist,
    read_insurance_source, status_label,
)
from upload_state import upload_widget_key


SBC_INSURANCE_PAGE_VERSION = 2


def render_sbc_insurance() -> None:
    st.title("SBC INSURANCE")
    st.caption("Filter Status Code first, then remove duplicate Financing/Card No. values. The first matching row is kept.")
    with st.sidebar:
        output_format = st.selectbox("Output files", ("CSV (.csv)", "Excel (.xlsx)"), key="sbc_insurance_output_format")
        csv_types = list(CSV_OUTPUT_ENCODINGS)
        csv_type = st.selectbox("CSV file type", csv_types, index=csv_types.index("CSV UTF-8 (Comma delimited)"),
                                key="sbc_insurance_csv_type")
    encoding = "utf-8-sig" if output_format == "Excel (.xlsx)" else CSV_OUTPUT_ENCODINGS[csv_type]
    source_panel, master_panel = st.columns(2)
    with source_panel:
        upload = st.file_uploader("Unfiltered SBC INSURANCE DataGrid", type=["xlsx"], key=upload_widget_key("sbc_insurance_upload"))
    with master_panel:
        master_upload = st.file_uploader(
            "SBC INSURANCE masterlist (optional)", type=["xlsx"], key=upload_widget_key("sbc_insurance_masterlist_upload"),
            help="Matches IS Ref No# to Account No., with PN_Number as a Financing/Card No. fallback. Supplies EMAIL, MOBILE and ADDRESS1–ADDRESS3.",
        )
    if upload is None:
        st.session_state.pop("sbc_insurance_result", None)
        return
    data = upload.getvalue()
    digest = hashlib.sha256(data).hexdigest()
    cache = st.session_state.setdefault("sbc_insurance_source_info", {})
    try:
        info_key = (digest, WORKBOOK_READER_VERSION, MAPPING_VERSION)
        if info_key not in cache:
            cache[info_key] = inspect_workbook(data)
        sheets = cache[info_key]
        names = [sheet.name for sheet in sheets]
        best = max(sheets, key=lambda sheet: sheet.matched_fields)
        selected = st.selectbox("Source sheet", names, index=names.index(best.name), key=f"sbc_insurance_sheet_{digest}")
        info = next(sheet for sheet in sheets if sheet.name == selected)
        with st.expander("Source options"):
            header = int(st.number_input("Source header row", min_value=1, max_value=1_048_575,
                                         value=info.header_row, key=f"sbc_insurance_header_{digest}_{selected}"))
        source_key = (digest, selected, header, SBC_INSURANCE_VERSION, WORKBOOK_READER_VERSION, MAPPING_VERSION)
        if source_key not in cache:
            with st.spinner("Reading insurance accounts and statuses..."):
                cache[source_key] = read_insurance_source(data, selected, header)
        source = cache[source_key]
        masterlist, master_key = None, None
        if master_upload is not None:
            master_data = master_upload.getvalue()
            master_digest = hashlib.sha256(master_data).hexdigest()
            master_info_key = ("masterlist_sheets", master_digest, SBC_INSURANCE_MASTERLIST_VERSION)
            if master_info_key not in cache:
                cache[master_info_key] = inspect_insurance_masterlist(master_data)
            master_sheets = cache[master_info_key]
            master_names = [sheet.name for sheet in master_sheets]
            default_master = next((name for name in master_names if name.strip().casefold() == "list"), master_names[0])
            with st.expander("Masterlist options"):
                master_name = st.selectbox("Masterlist sheet", master_names, index=master_names.index(default_master),
                                           key=f"sbc_insurance_master_sheet_{master_digest}")
                master_info = next(sheet for sheet in master_sheets if sheet.name == master_name)
                master_header = int(st.number_input("Masterlist header row", min_value=1, max_value=1_048_575,
                                                    value=master_info.header_row,
                                                    key=f"sbc_insurance_master_header_{master_digest}_{master_name}"))
            master_key = ("masterlist", master_digest, master_name, master_header, SBC_INSURANCE_MASTERLIST_VERSION)
            if master_key not in cache:
                with st.spinner("Reading masterlist email, mobile and postal addresses..."):
                    cache[master_key] = read_insurance_masterlist(master_data, master_name, master_header)
            masterlist = cache[master_key]
            st.caption(f"Masterlist: {master_name} · {len(masterlist.references):,} account references · {len(masterlist.accounts):,} distinct PN_Number accounts.")
        counts = source.status_counts
        options = sorted(counts, key=lambda value: (bool(value), value))
        status_widget = "sbc_insurance_statuses"
        if status_widget in st.session_state:
            previous = st.session_state[status_widget]
            valid = [status for status in previous if status in counts]
            if previous != valid:
                st.session_state[status_widget] = valid
        st.subheader("Status filter")
        included = st.multiselect("Status Code values to include", options, format_func=status_label, key=status_widget,
                                  help="Only rows with these statuses enter deduplication. Select (Blank status) to include blank statuses.")
        st.dataframe(pd.DataFrame([
            {"Status Code": status_label(status), "Source rows": counts[status], "Included": status in included}
            for status in options
        ]), hide_index=True, width="stretch", height=280)
        if not included:
            st.session_state.pop("sbc_insurance_result", None)
            st.info("Select the Status Code values to include. Files will be prepared automatically.")
            return
        signature = (source_key, tuple(sorted(included)), encoding, EXCEL_SAFE_CSV_VERSION, master_key)
        prepared = st.session_state.get("sbc_insurance_result")
        if prepared is None or prepared["signature"] != signature:
            st.session_state.pop("sbc_insurance_result", None)
            prepared = {"signature": signature, "result": prepare_insurance(source, included, encoding=encoding, masterlist=masterlist)}
            st.session_state["sbc_insurance_result"] = prepared
        result = prepared["result"]
    except AutomationError as exc:
        st.session_state.pop("sbc_insurance_result", None)
        st.error(str(exc))
        return
    st.subheader("Results")
    unique, duplicates, removed = st.columns(3)
    unique.metric("Unique accounts", f"{len(result.rows):,}")
    duplicates.metric("Duplicates removed", f"{result.stats['Duplicate financing/card accounts removed']:,}")
    removed.metric("Removed by status", f"{result.stats['Removed by status']:,}")
    st.dataframe(pd.DataFrame(result.stats.items(), columns=["Check", "Count"]), hide_index=True, width="stretch")
    if not result.rows:
        st.info("No account rows remain for the selected statuses.")
        return
    if result.contacts is None:
        st.caption("Uses the same 16-column campaign template as GEO, RECO, B2 and B4. Upload a masterlist to fill email and phone.")
    else:
        st.caption("EMAIL fills email and MOBILE fills phone. ADDRESS1–ADDRESS3 are included in the filtered DataGrid download; the campaign keeps its 16 columns.")
        st.caption("Account No. matches IS Ref No# first, then Financing/Card No. matches PN_Number when no reference is found. Repeated matches use the first valid contacts; differences are counted in Results.")
    if result.stats["Selected rows with blank Financing/Card No."]:
        st.info("Blank Financing/Card No. values form one duplicate group. The first row matching the selected statuses is kept, as in the sample output.")
    excel = output_format == "Excel (.xlsx)"
    file = result.file
    filename = Path(file.filename).with_suffix(".xlsx").name if excel else file.filename
    st.download_button(f"Download {filename}", partial(build_export_excel if excel else build_export_csv, file),
                       filename, mime=XLSX_MIME if excel else "text/csv", type="primary",
                       on_click="ignore", key="sbc_insurance_download")
    st.download_button("Download filtered DataGrid.xlsx", partial(build_filtered_datagrid, result),
                       "SBC INSURANCE Filtered DataGrid.xlsx", mime=XLSX_MIME,
                       on_click="ignore", key="sbc_insurance_filtered_download")
    st.subheader("Campaign preview")
    positions = template_column_positions(file.headers)
    targets = {index: target for target, index in positions.items()}
    st.dataframe(pd.DataFrame([
        [file.doclass if targets.get(index) == "doclass" else record.get(targets.get(index)) for index in range(len(file.headers))]
        for record in file.preview
    ], columns=file.headers).map(lambda value: "" if value is None else str(value)), hide_index=True, width="stretch")
    st.caption(f"{len(file.preview):,} / {file.count:,} accounts")
    if result.contacts is not None:
        st.subheader("Masterlist contact details")
        st.dataframe(pd.DataFrame([
            {"Financing/Card No.": row.financing, "EMAIL": contact.email if contact else None,
             "MOBILE": contact.phone if contact else None,
             **dict(zip(("ADDRESS1", "ADDRESS2", "ADDRESS3"), contact.address if contact else (None, None, None)))}
            for row, contact in list(zip(result.rows, result.contacts))[:100]
        ]).fillna(""), hide_index=True, width="stretch")
