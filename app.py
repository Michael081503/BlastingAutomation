"""Run with: python -m streamlit run app.py"""

from __future__ import annotations

import hashlib
from pathlib import Path

import pandas as pd
import streamlit as st

from export_config import CHANNEL_SUFFIXES, OUTPUT_FILENAMES, OUTPUT_FOLDER, ZIP_FILENAME
from exports import GROUP_LABELS, REVIEW_GROUP_KEYS, build_negative_exports, build_sheet_exports
from folder_output import save_files_to_folder

from automation import (
    FIELD_MAPPING,
    TARGET_HEADERS,
    XLSX_MIME,
    AutomationError,
    create_default_template,
    csv_template_headers,
    default_source_sheet,
    inspect_workbook,
    normalize_header,
    output_filename,
    read_source,
)


def workbook_info(data: bytes, *, template: bool = False):
    """Cache only workbook metadata, scoped to this browser session."""
    slot = "template_info" if template else "source_info"
    digest = hashlib.sha256(data).hexdigest()
    previous = st.session_state.get(slot)
    if previous is None or len(previous) != 3 or previous[:2] != (digest, TARGET_HEADERS):
        previous = (digest, TARGET_HEADERS, inspect_workbook(data, template=template))
        st.session_state[slot] = previous
    return previous[2]


def show_preview(records: list[dict], count: int) -> None:
    preview = pd.DataFrame(records, columns=TARGET_HEADERS, dtype=object)
    st.dataframe(
        preview.map(lambda value: "" if value is None else str(value)),
        hide_index=True,
        width="stretch",
    )
    st.caption(f"Showing the first {min(count, 100):,} rows. The file contains all {count:,} rows.")


def show_folder_save(result: dict, source_name: str, sheet_name: str) -> None:
    st.subheader("Save to a folder")
    output_folder = st.text_input("Output folder", value=OUTPUT_FOLDER, key="output_folder")
    st.caption(
        "Saves directly to the computer running this app. "
        "Each export creates a new folder with its date and time inside the location above. "
        "When hosted online, the folder is on the server; use the downloads to save files on your PC."
    )
    if st.button("Save files to folder", type="primary", key="save_to_folder"):
        files = [(file.filename, file.data) for file in result["files"]]
        try:
            folder = save_files_to_folder(
                files, output_folder, batch_name=f"{Path(source_name).stem}_{sheet_name}",
            )
            result["saved_folder"] = str(folder)
            result["saved_file_count"] = len(files)
            st.session_state["result"] = result
        except AutomationError as exc:
            st.error(str(exc))
    if result.get("saved_folder"):
        st.success(f"Saved {result['saved_file_count']} file(s) to this folder:")
        st.code(result["saved_folder"], language=None)


def main() -> None:
    st.set_page_config(page_title="Blasting Automation", page_icon="📋", layout="wide")
    st.title("Workbook to template")
    st.write("Upload your source file, choose a sheet, and fill your template with the mapped values.")
    if st.session_state.get("blank_template_headers") != TARGET_HEADERS:
        st.session_state["blank_template"] = create_default_template()
        st.session_state["blank_template_headers"] = TARGET_HEADERS
    blank_template = st.session_state["blank_template"]

    with st.sidebar:
        st.subheader("Your field mapping")
        st.dataframe(
            pd.DataFrame(FIELD_MAPPING.items(), columns=["Source header", "Template header"]),
            hide_index=True,
            width="stretch",
        )
        st.caption(
            "Headers ignore case, spacing, and punctuation. "
            "Product Type is accepted for product; OB is accepted for OB (Masterlist)."
        )
        st.divider()
        st.download_button(
            "Download blank template",
            data=blank_template,
            file_name="blank_template.xlsx",
            mime=XLSX_MIME,
            on_click="ignore",
        )
        st.caption("Save outputs to a folder or download them. Your original files are never overwritten.")

    upload_left, upload_right = st.columns(2)
    with upload_left:
        source_file = st.file_uploader("1. Source workbook", type=["xlsx"], key="source_upload")
    with upload_right:
        template_file = st.file_uploader(
            "2. Template (optional)",
            type=["xlsx", "csv"],
            key="template_upload",
            help="Upload your template, or use the blank template with all mapped target headers.",
        )

    if source_file is None:
        st.session_state.pop("result", None)
        st.info("Upload a source workbook to begin. NEGATIVE will be selected automatically if it exists.")
        st.caption("NEGATIVE exports are separated by Christmas Amnesty, product, and level.")
        return

    source_bytes = source_file.getvalue()
    template_bytes = template_file.getvalue() if template_file is not None else blank_template
    template_extension = Path(template_file.name).suffix.lower() if template_file is not None else ".xlsx"
    source_digest = hashlib.sha256(source_bytes).hexdigest()
    template_digest = hashlib.sha256(template_bytes).hexdigest()
    try:
        source_sheets = workbook_info(source_bytes)
        default_sheet = default_source_sheet(source_sheets)
        template_sheets = (
            workbook_info(template_bytes, template=True) if template_extension == ".xlsx" else []
        )
        if template_extension == ".csv":
            csv_template_headers(template_bytes)
    except AutomationError as exc:
        st.error(str(exc))
        return

    st.subheader("Choose the sheet to process")
    sheet_left, sheet_right = st.columns(2)
    source_names = [sheet.name for sheet in source_sheets]
    with sheet_left:
        source_sheet = st.selectbox(
            "Source sheet",
            source_names,
            index=source_names.index(default_sheet),
            key=f"source_sheet_{source_digest}",
        )
        source_info = next(sheet for sheet in source_sheets if sheet.name == source_sheet)
        st.caption("Start with NEGATIVE, then choose another sheet when ready.")
    with sheet_right:
        if template_sheets:
            best_template = max(template_sheets, key=lambda sheet: sheet.matched_fields)
            template_names = [sheet.name for sheet in template_sheets]
            template_sheet = st.selectbox(
                "Template sheet",
                template_names,
                index=template_names.index(best_template.name),
                key=f"template_sheet_{template_digest}",
            )
            template_info = next(sheet for sheet in template_sheets if sheet.name == template_sheet)
        else:
            template_sheet = ""
            template_info = None
            st.caption("The output will use your CSV template's column order and delimiter.")

    if not any(normalize_header(sheet.name) == "negative" for sheet in source_sheets):
        st.info("No NEGATIVE sheet was found. Choose the sheet you want to process.")
    if template_file is None:
        st.caption("Using the built-in template with all mapped target headers. Upload your template to use its layout.")
    else:
        st.caption(
            "The selected template sheet's data below the header will be replaced. "
            "Additional template columns will be blank; other sheets are kept."
        )

    grouped = False
    st.caption("Accounts are excluded only when both debtor_email and phone_1 are empty.")
    st.caption(
        "Every group produces Email and Viber files using the same headers. "
        "Accounts with both contact details appear in both files."
    )
    if normalize_header(source_sheet) == "negative":
        grouped = st.radio(
            "Output files",
            ["Separate Christmas Amnesty groups", "All accounts in one group"],
            key=f"output_mode_{source_digest}",
        ) == "Separate Christmas Amnesty groups"
        if grouped:
            st.caption(
                "QUALIFIED: CARDS Level 1, CARDS Level 6, and BEL at all levels. "
                "BLANK: the same three groups, plus other products. Empty groups produce no file."
            )
            with st.expander("Filename settings (export_config.py)"):
                st.caption(
                    "Edit the quoted filename values in export_config.py, save, and rerun the app. "
                    "The app adds _Email or _Viber, followed by .xlsx or .csv to match your template."
                )
                st.code(Path(__file__).with_name("export_config.py").read_text(encoding="utf-8"), language="python")

    with st.expander("Header rows", expanded=False):
        st.caption("Headers are detected within the first 50 rows. Change a row number if needed.")
        header_left, header_right = st.columns(2)
        with header_left:
            source_header = int(st.number_input(
                "Source header row",
                min_value=1,
                max_value=1_048_575,
                value=source_info.header_row,
                step=1,
                key=f"source_header_{source_digest}_{source_sheet}",
            ))
        with header_right:
            template_header = int(st.number_input(
                "Template header row",
                min_value=1,
                max_value=1_048_575,
                value=template_info.header_row if template_info is not None else 1,
                step=1,
                disabled=template_extension == ".csv",
                key=f"template_header_{template_digest}_{template_sheet}",
            ))

    signature = (
        source_digest, template_digest, source_sheet, source_header,
        template_sheet, template_header, template_extension,
        TARGET_HEADERS, grouped, tuple(OUTPUT_FILENAMES.items()), ZIP_FILENAME,
        tuple(CHANNEL_SUFFIXES.items()),
    )
    previous_result = st.session_state.get("result")
    if previous_result is not None and previous_result["signature"] != signature:
        st.session_state.pop("result", None)

    if st.button("Preview and prepare files", type="primary"):
        st.session_state.pop("result", None)
        try:
            with st.spinner(f"Reading {source_sheet} and filling the template…"):
                mapped = read_source(source_bytes, source_sheet, source_header)
                export_result = {}
                if mapped.records:
                    if grouped:
                        bundle = build_negative_exports(
                            template_bytes, mapped.records, template_sheet,
                            template_header, template_extension,
                        )
                    else:
                        bundle = build_sheet_exports(
                            template_bytes, mapped.records, template_sheet, template_header,
                            base_name=output_filename(source_file.name, source_sheet, template_extension),
                            label=source_sheet, extension=template_extension,
                        )
                    export_result = {
                        "files": bundle.files,
                        "group_counts": bundle.counts,
                        "zip_data": bundle.zip_data,
                        "zip_filename": bundle.zip_filename,
                    }
                st.session_state["result"] = {
                    "signature": signature,
                    "grouped": grouped,
                    "count": len(mapped.records),
                    "blank_rows": mapped.blank_rows_skipped,
                    "no_contact_rows": mapped.no_contact_rows_skipped,
                    "missing_emails": sum(not record["email"] for record in mapped.records),
                    **export_result,
                }
        except AutomationError as exc:
            st.error(str(exc))
        except Exception:
            st.error(
                "The workbook could not be processed. Check for unsupported Excel features "
                "or damaged cells, and try saving a fresh .xlsx copy in Excel."
            )

    result = st.session_state.get("result")
    if result is None:
        return
    if result["count"]:
        st.success(f"Ready: {result['count']:,} accounts from {source_sheet}, in {len(result['files'])} files.")
    else:
        st.warning("No accounts remain to export. Check the source rows and excluded-account count below.")
    count_col, blank_col, excluded_col, email_col = st.columns(4)
    count_col.metric("Accounts included", f"{result['count']:,}")
    blank_col.metric("Empty rows skipped", f"{result['blank_rows']:,}")
    excluded_col.metric("Removed: no email and phone", f"{result['no_contact_rows']:,}")
    email_col.metric("Phone-only accounts", f"{result['missing_emails']:,}")
    if not result["count"]:
        return
    show_folder_save(result, source_file.name, source_sheet)
    mime = "text/csv" if template_extension == ".csv" else XLSX_MIME
    files = result["files"]
    counts = result["group_counts"]
    labels = GROUP_LABELS if result["grouped"] else {"all": source_sheet}
    by_channel = {(file.group_key, file.channel): file for file in files}
    st.subheader("Separated files")
    st.dataframe(pd.DataFrame([
        {
            "Group": label,
            "Accounts": counts[key],
            "Email rows": by_channel[(key, "email")].count if (key, "email") in by_channel else 0,
            "Viber rows": by_channel[(key, "viber")].count if (key, "viber") in by_channel else 0,
        }
        for key, label in labels.items()
        if key not in REVIEW_GROUP_KEYS or counts[key]
    ]), hide_index=True, width="stretch")
    st.caption(
        "Accounts with both contact details are counted once under Accounts and appear in both files. "
        "Email files leave phone blank; Viber files leave email blank. All headers are retained."
    )
    if result["grouped"]:
        for key in GROUP_LABELS:
            if key in REVIEW_GROUP_KEYS and counts[key]:
                st.warning(
                    f"{counts[key]:,} rows need review ({GROUP_LABELS[key]}). "
                    "They are included in a separate review file."
                )
    st.download_button(
        "Download all files (ZIP)", result["zip_data"], result["zip_filename"],
        mime="application/zip", type="primary", on_click="ignore",
    )
    for file in files:
        st.download_button(
            f"Download {file.filename}", file.data, file.filename,
            mime=mime, key=f"download_{file.key}", on_click="ignore",
        )
    st.subheader("Preview an export")
    selected_file = st.selectbox(
        "File to preview", [file.filename for file in files],
        key=f"preview_file_{hash(signature)}",
    )
    file = next(file for file in files if file.filename == selected_file)
    show_preview(file.preview, file.count)


if __name__ == "__main__":
    main()
