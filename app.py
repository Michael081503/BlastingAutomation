"""Run with: python -m streamlit run app.py"""

from __future__ import annotations

import hashlib
import importlib
from dataclasses import replace
from functools import partial
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl.utils import get_column_letter

import app_startup

importlib.reload(app_startup).refresh_application_modules()

from branding import BANK_ICON, render_bank_branding
from export_config import (
    CHANNEL_SUFFIXES, GEO_OUTPUT_FILENAMES, GEO_ZIP_FILENAME,
    OUTPUT_FILENAMES, OUTPUT_FOLDER, ZIP_FILENAME,
)
from exports import (
    GEO_FILENAME_CONFIG_VERSION, GEO_GROUP_LABELS, GEO_REVIEW_GROUP_KEYS, GROUP_LABELS, REVIEW_GROUP_KEYS,
    GROUPED_SHEET_EXPORT_VERSION,
    EXCEL_EXPORT_VERSION, build_excel_zip, build_export_excel,
    build_geo_exports, build_negative_exports, build_sheet_exports,
)
from folder_output import save_files_to_folder
from email_validation import EMAIL_VALIDATION_VERSION

from automation import (
    CSV_OUTPUT_ENCODINGS,
    DEFAULT_TEMPLATE_HEADERS,
    DERIVED_TEMPLATE_HEADERS,
    FIELD_MAPPING,
    MAPPING_VERSION,
    OB_SORT_VERSION,
    WORKBOOK_READER_VERSION,
    OPTIONAL_FIELD_MAPPING,
    TARGET_HEADERS,
    XLSX_MIME,
    AutomationError,
    create_default_csv_template,
    csv_template_headers,
    default_source_sheet,
    inspect_workbook,
    inspect_source_sheet,
    preview_source_rows,
    source_column_positions,
    normalize_header,
    output_filename,
    read_source,
)


def workbook_info(data: bytes, *, template: bool = False):
    """Cache only workbook metadata, scoped to this browser session."""
    slot = "template_info" if template else "source_info"
    digest = hashlib.sha256(data).hexdigest()
    signature = (digest, MAPPING_VERSION, WORKBOOK_READER_VERSION)
    previous = st.session_state.get(slot)
    if previous is None or len(previous) != 4 or previous[:3] != signature:
        previous = (*signature, inspect_workbook(data, template=template))
        st.session_state[slot] = previous
    return previous[3]


def show_preview(records: list[dict], count: int, headers: list[str]) -> None:
    targets = {
        normalize_header(target): target
        for target in (*TARGET_HEADERS, *OPTIONAL_FIELD_MAPPING.values(), *DERIVED_TEMPLATE_HEADERS)
    }
    columns = [targets.get(normalize_header(header)) for header in headers]
    preview = pd.DataFrame(
        [[record.get(target) for target in columns] for record in records],
        columns=headers,
        dtype=object,
    )
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
        output_files = [(file.filename, file.data) for file in result["files"]]
        try:
            folder = save_files_to_folder(
                output_files, output_folder, batch_name=f"{Path(source_name).stem}_{sheet_name}",
            )
            result["saved_folder"] = str(folder)
            result["saved_file_count"] = len(output_files)
            st.session_state["result"] = result
        except AutomationError as exc:
            st.error(str(exc))
    if result.get("saved_folder"):
        st.success(f"Saved {result['saved_file_count']} file(s) to this folder:")
        st.code(result["saved_folder"], language=None)


def render_workbook_page() -> None:
    render_bank_branding()
    st.title("GEO & RECO")
    st.write("Upload your GEO or RECO workbook, choose a sheet, and export the mapped values as CSV files.")
    st.caption("GEO separates Mindanao and R3. RECO uses the existing product and level groups. B2 matching has its own NEGATIVE B2 page.")
    with st.sidebar:
        output_format = st.selectbox(
            "CSV file type", list(CSV_OUTPUT_ENCODINGS), key="csv_file_type",
            help="Use CSV UTF-8 for names or other text that the standard CSV format cannot represent.",
        )
    encoding = CSV_OUTPUT_ENCODINGS[output_format]
    if st.session_state.get("blank_template_headers") != (DEFAULT_TEMPLATE_HEADERS, encoding):
        st.session_state["blank_template"] = create_default_csv_template(encoding=encoding)
        st.session_state["blank_template_headers"] = (DEFAULT_TEMPLATE_HEADERS, encoding)
    blank_template = st.session_state["blank_template"]

    with st.sidebar:
        st.subheader("Your field mapping")
        st.dataframe(
            pd.DataFrame({**FIELD_MAPPING, **OPTIONAL_FIELD_MAPPING}.items(), columns=["Source header", "Template header"]),
            hide_index=True,
            width="stretch",
        )
        st.caption(
            "Headers ignore case, spacing, and punctuation. "
            "GEO aliases are accepted, including email, account, card_no, assign_date, and placement. "
            "Area is copied when present; it is optional for other workbooks."
        )
        st.divider()
        st.download_button(
            "Download blank template",
            data=blank_template,
            file_name="blank_template.csv",
            mime="text/csv",
            on_click="ignore",
        )
        st.caption("Save outputs to a folder or download them. Your original files are never overwritten.")

    upload_left, upload_right = st.columns(2)
    with upload_left:
        source_file = st.file_uploader("1. Source workbook", type=["xlsx"], key="source_upload")
    if source_file is not None:
        source_bytes = source_file.getvalue()
        try:
            source_sheets = workbook_info(source_bytes)
        except AutomationError as exc:
            st.session_state.pop("result", None)
            st.error(str(exc))
            return
    with upload_right:
        template_file = st.file_uploader(
            "2. Template (optional)",
            type=["xlsx", "csv"],
            key="template_upload",
            help="Upload your template, or use the built-in template with all 16 campaign headers.",
        )

    if source_file is None:
        st.session_state.pop("result", None)
        st.info("Upload a source workbook to begin. NEGATIVE will be selected automatically if it exists.")
        st.caption("Every worksheet uses the same separation rules. GEO worksheets separate Mindanao and R3 first.")
        return

    template_bytes = template_file.getvalue() if template_file is not None else blank_template
    template_extension = Path(template_file.name).suffix.lower() if template_file is not None else ".csv"
    source_digest = hashlib.sha256(source_bytes).hexdigest()
    template_digest = hashlib.sha256(template_bytes).hexdigest()
    try:
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
        st.caption("Start with NEGATIVE, then choose another sheet. The same grouping rules apply to every sheet.")
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
            st.caption("CSV output keeps the template's column order and always separates columns with commas.")

    if not any(normalize_header(sheet.name) == "negative" for sheet in source_sheets):
        st.info("No NEGATIVE sheet was found. Choose the sheet you want to process.")
    if template_file is None:
        st.caption("Using the built-in CSV template with all 16 campaign headers. Upload your template to use its column layout.")
    else:
        st.caption(
            "Outputs are always CSV, using only the selected template headers and column order. "
            "Additional template columns will be blank."
        )

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
    if source_header != source_info.header_row:
        try:
            source_info = inspect_source_sheet(source_bytes, source_sheet, source_header)
        except AutomationError as exc:
            st.error(str(exc))
            return

    defaults = source_column_positions(source_info.headers)
    missing_contact = not {"email", "phone"}.intersection(defaults)
    missing_account = not {"primary_key", "cardnumber"}.intersection(defaults)
    source_columns = None
    with st.expander("Source columns", expanded=missing_contact or missing_account):
        st.caption("Check the actual column names and sample values. Change Header rows above if the labels are on a different row.")
        preview_key = (source_digest, source_sheet, source_header, MAPPING_VERSION)
        previews = st.session_state.setdefault("source_previews", {})
        if preview_key not in previews:
            previews[preview_key] = preview_source_rows(source_bytes, source_sheet, source_header)
        labels = [f"{get_column_letter(index + 1)}: {header or '(blank header)'}"
                  for index, header in enumerate(source_info.headers)]
        sample_rows = [list(row[:len(labels)]) + [None] * max(0, len(labels) - len(row)) for row in previews[preview_key]]
        sample = pd.DataFrame(sample_rows, columns=labels, dtype=object)
        st.dataframe(sample.map(lambda value: "" if value is None else str(value)), hide_index=True, width="stretch")
        manual = st.checkbox(
            "Map source columns manually", value=missing_contact or missing_account,
            key=f"manual_mapping_{source_digest}_{source_sheet}_{source_header}_{MAPPING_VERSION}",
        )
        if manual:
            st.caption("Choose which source column supplies each template field. Leave fields without a source as Not available.")
            source_columns = {}
            options = [None, *range(len(source_info.headers))]
            fields = st.columns(3)

            def column_label(index, choices=labels):
                return "Not available" if index is None else choices[index]

            for number, target in enumerate((*TARGET_HEADERS, *OPTIONAL_FIELD_MAPPING.values())):
                with fields[number % 3]:
                    source_columns[target] = st.selectbox(
                        f"Source column for {target}", options, index=options.index(defaults.get(target)),
                        format_func=column_label,
                        key=f"source_mapping_{source_digest}_{source_sheet}_{source_header}_{MAPPING_VERSION}_{target}",
                    )
            available = tuple(target for target, column in source_columns.items() if column is not None)
            source_info = replace(source_info, available_fields=available, has_area="AREA" in available)

    can_prepare = bool({"email", "phone"}.intersection(source_info.available_fields)
                       and {"primary_key", "cardnumber"}.intersection(source_info.available_fields))
    if not can_prepare:
        st.info("Select an email or phone column and an account-number column under Source columns to prepare the template.")
        st.session_state.pop("result", None)
    missing_fields = [target for target in TARGET_HEADERS if target not in source_info.available_fields]
    if missing_fields:
        st.info("Available values will be copied. These fields are missing from the source and will stay blank: " + ", ".join(missing_fields))
    geo = source_info.has_area
    if geo:
        st.caption("doclass is filled automatically: Mindanao = SBC GEO Mindanao; Region 3 / R3 = SBC GEO R3.")
    st.caption("Accounts are excluded when they have neither a valid email address nor a phone number.")
    st.caption(
        "Every group produces Email and Viber CSV files using the same headers. "
        "Accounts with both contact details appear in both files."
    )
    st.caption(
        "Invalid email formats are removed from Email files. Spaces around valid addresses are trimmed. "
        "This checks the address format; it does not confirm that a mailbox exists or can receive mail."
    )
    grouped_label = "Separate Area and Christmas Amnesty groups" if geo else "Separate Christmas Amnesty groups"
    grouping_fields = {"ChristmasAmnesty": "Christmas Amnesty", "product": "product"}
    if not geo:
        grouping_fields["level"] = "level"
    missing_grouping = [label for field, label in grouping_fields.items() if field not in source_info.available_fields]
    output_modes = ["All accounts in one group"] if missing_grouping else [grouped_label, "All accounts in one group"]
    if missing_grouping:
        st.caption("Grouping columns are missing: " + ", ".join(missing_grouping) + ". All available accounts will be prepared as Email and Viber files.")
    saved_modes = st.session_state.setdefault("sheet_output_modes", {})
    mode_key = (source_digest, source_sheet, source_header)
    previous_mode = saved_modes.get(mode_key, output_modes[0])
    output_mode = st.radio(
        "Output files",
        output_modes,
        index=output_modes.index(previous_mode) if previous_mode in output_modes else 0,
        key=f"output_mode_{source_digest}_{source_sheet}_{source_header}_{MAPPING_VERSION}",
    )
    saved_modes[mode_key] = output_mode
    grouped = output_mode == grouped_label
    if grouped:
        if geo:
            st.caption(
                "Mindanao and R3 (Region 3) are separated first. "
                "Other or missing Area values go to separate review files. AREA retains its source value."
            )
            st.caption(
                "QUALIFIED: CARDS and BEL. BLANK: CARDS, BEL, and other products. "
                "Each product group combines all levels. Empty groups produce no file. "
                "Qualified other products go to a separate review group."
            )
        else:
            st.caption(
                "QUALIFIED: CARDS Level 1, CARDS Level 6, and BEL at all levels. "
                "BLANK: the same three groups, plus other products. Empty groups produce no file."
            )
        if normalize_header(source_sheet) != "negative":
            st.caption("Grouped filenames include the selected worksheet name before the configured group name.")
    st.caption("Output filenames are set in export_config.py. Edit the quoted names there, then prepare the files again.")

    signature = (
        source_digest, template_digest, source_sheet, source_header,
        template_sheet, template_header, template_extension,
        TARGET_HEADERS, DEFAULT_TEMPLATE_HEADERS, grouped, tuple(OUTPUT_FILENAMES.items()), ZIP_FILENAME,
        tuple(CHANNEL_SUFFIXES.items()), EMAIL_VALIDATION_VERSION, ".csv", encoding, MAPPING_VERSION, geo,
        tuple(GEO_OUTPUT_FILENAMES.items()), GEO_ZIP_FILENAME,
        "config-filenames-and-email-summary-v1", GEO_FILENAME_CONFIG_VERSION, GROUPED_SHEET_EXPORT_VERSION, OB_SORT_VERSION,
        EXCEL_EXPORT_VERSION,
        tuple(source_columns.items()) if source_columns is not None else None,
    )
    previous_result = st.session_state.get("result")
    if previous_result is not None and previous_result["signature"] != signature:
        st.session_state.pop("result", None)

    if st.button("Preview and prepare files", type="primary", disabled=not can_prepare):
        st.session_state.pop("result", None)
        try:
            with st.spinner(f"Reading {source_sheet} and filling the template…"):
                mapped = read_source(
                    source_bytes, source_sheet, source_header, allow_missing=True, source_columns=source_columns,
                )
                export_result = {}
                if mapped.records:
                    if grouped:
                        builder = build_geo_exports if geo else build_negative_exports
                        bundle = builder(
                            template_bytes, mapped.records, template_sheet,
                            template_header, template_extension, encoding=encoding, source_sheet=source_sheet,
                        )
                    else:
                        bundle = build_sheet_exports(
                            template_bytes, mapped.records, template_sheet, template_header,
                            base_name=output_filename(source_file.name, source_sheet),
                            label=source_sheet, template_extension=template_extension, encoding=encoding,
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
                    "geo": geo,
                    "count": len(mapped.records),
                    "blank_rows": mapped.blank_rows_skipped,
                    "no_contact_rows": mapped.no_contact_rows_skipped,
                    "invalid_emails": mapped.invalid_email_rows,
                    "valid_emails": mapped.valid_email_rows,
                    "missing_email_rows": mapped.missing_email_rows,
                    "emails_checked": mapped.valid_email_rows + mapped.invalid_email_rows + mapped.missing_email_rows,
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
    excluded_col.metric("Removed: no valid email or phone", f"{result['no_contact_rows']:,}")
    email_col.metric("Phone-only accounts", f"{result['missing_emails']:,}")
    checked_col, valid_col, invalid_col, missing_col = st.columns(4)
    checked_col.metric("Accounts checked for email", f"{result['emails_checked']:,}")
    valid_col.metric("Valid email formats", f"{result['valid_emails']:,}")
    invalid_col.metric("Invalid emails removed", f"{result['invalid_emails']:,}")
    missing_col.metric("Missing emails", f"{result['missing_email_rows']:,}")
    if result["invalid_emails"]:
        st.info(
            f"Removed {result['invalid_emails']:,} invalid email address(es) from Email exports. "
            "Accounts with a phone number are still included in Viber exports."
        )
    if not result["count"]:
        return
    files = result["files"]
    show_folder_save(result, source_file.name, source_sheet)
    counts = result["group_counts"]
    labels = (GEO_GROUP_LABELS if result["geo"] else GROUP_LABELS) if result["grouped"] else {"all": source_sheet}
    review_keys = GEO_REVIEW_GROUP_KEYS if result["geo"] else REVIEW_GROUP_KEYS
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
        if key not in review_keys or counts[key]
    ]), hide_index=True, width="stretch")
    st.caption(
        "Accounts with both contact details are counted once under Accounts and appear in both files. "
        "Email files leave phone blank; Viber files leave email blank. All headers are retained."
    )
    if result["grouped"]:
        for key in labels:
            if key in review_keys and counts[key]:
                st.warning(
                    f"{counts[key]:,} rows need review ({labels[key]}). "
                    "They are included in a separate review file."
                )
    st.caption(
        "Use Excel files to open, edit, or upload accounts to AUTOSTAT. "
        "Account numbers are stored as text to preserve every supplied digit. "
        "Extract the Excel ZIP before opening its workbooks."
    )
    st.download_button(
        "Download Excel files (ZIP, keeps account numbers as text)", partial(build_excel_zip, files),
        Path(result["zip_filename"]).stem + "_Excel.zip",
        mime="application/zip", type="primary", on_click="ignore", key="excel_zip",
    )
    with st.expander("CSV files for system upload"):
        st.caption("Opening CSV files directly in Excel can round account digits to zeros. Use Excel files for viewing and editing.")
        st.download_button(
            "Download all files (ZIP)", result["zip_data"], result["zip_filename"],
            mime="application/zip", on_click="ignore",
        )
        for file in files:
            st.download_button(
                f"Download {file.filename}", file.data, file.filename,
                mime="text/csv", key=f"download_{file.key}", on_click="ignore",
            )
    st.subheader("Preview an export")
    selected_file = st.selectbox(
        "File to preview", [file.filename for file in files],
        key=f"preview_file_{hash(signature)}",
    )
    file = next(file for file in files if file.filename == selected_file)
    st.download_button(
        "Download Excel for this export", partial(build_export_excel, file),
        Path(file.filename).with_suffix(".xlsx").name, mime=XLSX_MIME,
        key=f"excel_download_{file.key}", on_click="ignore",
    )
    show_preview(file.preview, file.count, file.headers)


def main() -> None:
    st.set_page_config(page_title="Security Bank | Blasting Automation", page_icon=str(BANK_ICON), layout="wide")
    page = st.navigation([
        st.Page(render_workbook_page, title="GEO & RECO", icon=":material/folder_copy:", default=True),
        st.Page("pages/NEGATIVE_B2.py", title="NEGATIVE B2", icon=":material/compare_arrows:", url_path="NEGATIVE_B2"),
        st.Page("pages/AUTOSTAT.py", title="AUTOSTAT", icon=":material/mark_email_read:", url_path="AUTOSTAT"),
    ])
    page.run()


if __name__ == "__main__":
    main()
