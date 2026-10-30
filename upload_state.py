"""Clear uploaded files and their prepared outputs across all application pages."""

from __future__ import annotations

import re

import streamlit as st


CLEAR_UPLOADS_VERSION = 1
B4_UPLOADS_VERSION = 1
UPLOAD_GENERATION_KEY = "_upload_generation"
UPLOAD_KEYS = (
    "source_upload", "template_upload", "geo_b2_autostat_uploads",
    "autostat_uploads", "autostat_b2_uploads", "b2_database_upload", "b2_grid_upload",
    "b4_database_upload", "b4_grid_upload",
)
FILE_STATE_KEYS = {
    "result", "source_info", "template_info", "source_previews", "sheet_output_modes",
    "geo_autostat_result", "geo_b2_autostat_result", "geo_b2_autostat_source_info",
    "autostat_result", "autostat_source_info", "autostat_b2_result", "autostat_b2_source_info",
    "b2_result", "b2_source_info", "b4_result", "b4_source_info",
}
FILE_STATE_PREFIXES = (
    "source_sheet_", "template_sheet_", "source_header_", "template_header_",
    "manual_mapping_", "source_mapping_", "output_mode_", "preview_file_",
    "download_", "excel_download_", "b2_sheet_", "b2_header_", "b2_preview_",
    "autostat_preview_", "autostat_b2_preview_", "geo_b2_autostat_preview_",
    "b4_sheet_", "b4_header_", "b4_preview_",
)
SETTINGS_KEYS = {"csv_file_type", "output_folder", "autostat_upload_type", "b2_csv_type", "b2_output_folder", "b4_csv_type", "b4_output_folder"} | {
    prefix + suffix
    for prefix in ("autostat_", "autostat_b2_", "geo_b2_autostat_", "geo_autostat_")
    for suffix in ("output_format", "csv_file_type", "remark_by", "collector", "blank_time",
                   "current_time", "date", "time", "default_time", "output_folder")
}


def upload_widget_key(name: str) -> str:
    """Replace uploader widgets after clearing so old browser selections cannot return."""
    generation = st.session_state.get(UPLOAD_GENERATION_KEY, 0)
    return f"{name}_clear_{generation}" if generation else name


def clear_all_files() -> None:
    """Remove only file-related session state; keep the user's page and settings."""
    generation = st.session_state.get(UPLOAD_GENERATION_KEY, 0) + 1
    settings = dict(st.session_state.get("_retained_upload_settings", {}))
    settings.update({key: st.session_state[key] for key in SETTINGS_KEYS if key in st.session_state})
    st.session_state["_retained_upload_settings"] = settings
    for key in list(st.session_state):
        if (
            key in FILE_STATE_KEYS or key in UPLOAD_KEYS
            or key.startswith(FILE_STATE_PREFIXES)
            or any(key.startswith(name + "_clear_") for name in UPLOAD_KEYS)
            or re.match(r"(?:autostat|autostat_b2|geo_b2_autostat)_\d+_", key)
        ):
            del st.session_state[key]
    st.session_state[UPLOAD_GENERATION_KEY] = generation


def render_clear_all_files() -> None:
    # Conditional log controls disappear with their uploads. Keep those settings
    # independently of Streamlit's cleanup of widgets that are no longer shown.
    settings = st.session_state.get("_retained_upload_settings", {})
    for key in settings:
        if key in st.session_state:
            settings[key] = st.session_state[key]
        else:
            st.session_state[key] = settings[key]
    st.sidebar.button(
        "Clear all files", key="clear_all_files", on_click=clear_all_files,
        help="Clears uploaded files and prepared downloads on every page. Your settings are kept.",
        icon=":material/clear_all:",
    )
