"""Refresh cached application modules after a Streamlit code update."""

from __future__ import annotations

import importlib
from pathlib import Path
import sys
from threading import RLock
from time import sleep


# Retain the same lock if this helper is explicitly reloaded during development.
if "_REFRESH_LOCK" not in globals():
    _REFRESH_LOCK = RLock()
STARTUP_RECOVERY_VERSION = 1
_MAX_REFRESH_ATTEMPTS = 3


def _evicted_application_module(error: KeyError | ImportError) -> str | None:
    """Recognize import machinery losing an application module, not application errors."""
    traceback = error.__traceback__
    while traceback is not None and traceback.tb_next is not None:
        traceback = traceback.tb_next
    if traceback is None:
        return None
    frame = traceback.tb_frame
    origin = frame.f_globals.get("__name__")
    if isinstance(error, KeyError):
        if (origin not in {"importlib._bootstrap", "_frozen_importlib"}
                or frame.f_code.co_name not in {"_load_unlocked", "_find_and_load", "_find_and_load_unlocked", "_exec"}
                or len(error.args) != 1):
            return None
        name = error.args[0]
    else:
        if origin not in {"importlib", "importlib._bootstrap", "_frozen_importlib"}:
            return None
        name = error.name
    if (not isinstance(name, str) or not name.isidentifier() or name in sys.modules
            or not (Path(__file__).resolve().parent / f"{name}.py").is_file()):
        return None
    return name


def _has_required_name(module: object, name: str) -> bool:
    value = module
    try:
        for part in name.split("."):
            value = value[part] if isinstance(value, dict) else getattr(value, part)
    except (AttributeError, KeyError):
        return False
    return True


def refresh_application_modules() -> set[str]:
    """Serialize refreshes across sessions and recover a bounded import-cache race."""
    with _REFRESH_LOCK:
        refreshed = set()
        for attempt in range(_MAX_REFRESH_ATTEMPTS):
            try:
                return _refresh_application_modules(refreshed)
            except (KeyError, ImportError) as error:
                evicted = _evicted_application_module(error)
                if evicted is None or attempt == _MAX_REFRESH_ATTEMPTS - 1:
                    raise
                # Replay the dependency pass so cached consumers receive the new module.
                refreshed.add(evicted)
                importlib.invalidate_caches()
                sleep(0.05)
    raise RuntimeError("Application startup did not complete.")


def _refresh_application_modules(refreshed: set[str]) -> set[str]:
    """Reload missing interfaces and their dependents in dependency order.

    A running cloud process can retain modules from before a deployment. Only
    refresh modules when required names are missing, so routine reruns keep
    exception and dataclass identities stable.
    """
    modules = (
        ("upload_state", ("CLEAR_UPLOADS_VERSION", "WORKSPACE_ACTIONS_VERSION", "B4_UPLOADS_VERSION", "B4_AUTOSTAT_UPLOADS_VERSION", "CAMPAIGN_FORMAT_SETTINGS_VERSION", "EXCEL_SAFE_CSV_SETTINGS_VERSION", "upload_widget_key", "clear_all_files", "render_clear_all_files"), ()),
        ("branding", (
            "render_bank_branding", "BANK_ICON", "HALLOWEEN_THEME_VERSION", "HALLOWEEN_GHOST_VERSION", "HALLOWEEN_SPIDER_VERSION",
            "render_clear_all_files",
        ), ("upload_state",)),
        ("email_validation", ("EMAIL_VALIDATION_VERSION", "validated_email"), ()),
        ("export_config", (
            "CHANNEL_SUFFIXES", "OUTPUT_FILENAMES", "OUTPUT_FOLDER", "ZIP_FILENAME",
            "GEO_OUTPUT_FILENAMES", "GEO_ZIP_FILENAME", "AUTOSTAT_BEL_FILENAMES",
            "GEO_LOANS_FILENAMES_VERSION",
            "AUTOSTAT_B2_FILENAMES", "AUTOSTAT_B4_FILENAMES", "COLLECTION_AGENTS",
            "GEO_RECO_AGENTS", "GEO_RECO_CYCLE_START",
            "GEO_AGENTS", "GEO_AGENT_ROSTER_VERSION",
            "RECO_DOCLASS_BY_FILENAME",
            "GEO_OUTPUT_FILENAMES.mindanao_qualified_cards", "GEO_OUTPUT_FILENAMES.r3_qualified_cards",
        ), ()),
        ("automation", (
            "MAPPING_VERSION", "OPTIONAL_FIELD_MAPPING", "CSV_OUTPUT_ENCODINGS",
            "PRODUCT_NORMALIZATION_VERSION", "normalize_product",
            "create_default_csv_template",
            "DERIVED_TEMPLATE_HEADERS", "geo_doclass",
            "ASSIGNMENT_TEMPLATE_HEADERS",
            "campaign_output_headers", "template_column_positions",
            "SheetInfo.available_fields", "inspect_source_sheet",
            "SheetInfo.headers", "source_column_positions", "preview_source_rows",
            "WORKBOOK_READER_VERSION",
            "SOURCE_SHEET_SELECTION_VERSION", "default_source_sheet",
            "OB_SORT_VERSION", "ob_sort_key",
            "RECO_DOCLASS_WRITER_VERSION",
            "AUTOSTAT_COLUMN_FORMAT_VERSION",
            "write_text_csv_sheet",
            "write_typed_template_sheet",
            "MappedData.valid_email_rows", "MappedData.missing_email_rows",
        ), ("email_validation",)),
        ("agent_cycle", (
            "AGENT_CYCLE_VERSION", "AgentCycle", "load_cycle", "save_cycle", "cycle_agents", "assign_agents_to_accounts", "reorder_cycle",
        ), ("automation", "export_config")),
        ("agent_availability", (
            "AGENT_AVAILABILITY_VERSION", "agent_roster", "available_agents", "save_available_agents",
        ), ("automation", "export_config")),
        ("agent_availability_page", (
            "AGENT_AVAILABILITY_PAGE_VERSION", "AGENT_ABSENCE_CHECKBOX_VERSION", "render_agent_availability",
        ), ("agent_availability", "automation")),
        ("exports", (
            "GEO_GROUP_LABELS", "GEO_PRODUCT_GROUP_LABELS", "GEO_REVIEW_GROUP_KEYS", "build_geo_exports",
            "GEO_FILENAME_CONFIG_VERSION",
            "GEO_LOANS_VERSION", "GEO_LOAN_PRODUCTS",
            "GROUPED_SHEET_EXPORT_VERSION",
            "EXCEL_EXPORT_VERSION", "build_export_excel", "build_excel_zip",
            "TYPED_EXCEL_EXPORT_VERSION", "ExportFile.__dataclass_fields__.records",
            "CAMPAIGN_OUTPUT_FORMAT_VERSION", "CSV_CAMPAIGN_DEFAULT_VERSION", "campaign_output_files",
            "EXCEL_SAFE_CSV_VERSION", "build_export_csv", "build_csv_zip",
            "CSV_PROTECTED_IDENTIFIER_FIELDS",
            "ACTIVITY_CSV_IDENTIFIERS_VERSION",
            "REQUIRED_CSV_PROTECTION_VERSION",
            "AGENT_ASSIGNMENT_EXPORT_VERSION",
            "RECO_ONLY_AGENTS_VERSION",
            "GEO_FIXED_AGENTS_VERSION",
            "GEO_AGENT_AVAILABILITY_VERSION",
            "RECO_DOCLASS_VERSION", "reco_doclass", "ExportFile.__dataclass_fields__.doclass",
            "SOURCE_MEMBERSHIP_VERSION", "source_account_counts", "validate_source_exports",
        ), ("email_validation", "export_config", "automation", "agent_cycle")),
        ("folder_output", ("save_files_to_folder", "FOLDER_OUTPUT_VERSION"), ("automation", "exports")),
        ("agent_order", ("AGENT_ORDER_VERSION", "AGENT_ORDER_CALLBACK_VERSION", "AGENT_NATIVE_ROWS_VERSION", "render_agent_order"), ("agent_availability_page",)),
        ("agent_cycle_page", ("AGENT_CYCLE_PAGE_VERSION", "AGENT_DRAG_ORDER_VERSION", "AGENT_STABLE_ORDER_VERSION", "RECO_AGENT_AVAILABILITY_VERSION", "RECO_AGENT_VISIBLE_SECTION_VERSION", "RECO_AGENT_IMMEDIATE_PANEL_VERSION", "render_agent_cycle"), ("agent_cycle", "agent_order", "folder_output", "agent_availability_page")),
        ("autostat", (
            "AUTOSTAT_VERSION", "AUTOSTAT_REMARK_BY", "PREFERRED_ACCOUNT_HEADER", "EMAIL_HEADERS", "VIBER_HEADERS", "build_autostat_exports",
            "AUTOSTAT_24_HOUR_TIME_VERSION",
            "build_bel_exports", "has_bel_primary_keys",
            "ACCOUNT_DEDUP_VERSION", "ACCOUNT_IDENTITY_VERSION", "unique_account_contacts",
            "COMPILED_OUTPUT_VERSION", "BLANK_TIMESTAMP_VERSION",
            "AUTOSTAT_EXCLUSIONS_VERSION", "autostat_source_excluded",
            "CONTACT_VALIDATION_VERSION", "_validated_candidates",
            "ACCOUNT_TEXT_EXPORT_VERSION", "build_autostat_excel",
            "ACCOUNT_NUMBER_NORMALIZATION_VERSION", "account_number_text",
            "ACCOUNT_INTEGRITY_VERSION", "EXCEL_CHANNEL_OUTPUT_VERSION",
            "SOURCE_ROW_SKIP_VERSION", "ContactRows.__dataclass_fields__.source_error_rows",
            "build_autostat_channel_excel", "build_autostat_excel_zip",
            "ContactRows.__dataclass_fields__.email_ob_values", "ContactRows.__dataclass_fields__.phone_ob_values",
        ), ("automation", "email_validation", "export_config", "exports")),
        ("autostat_downloads", ("AUTOSTAT_DOWNLOADS_VERSION", "ZIP_FORMAT_VERSION", "render_autostat_downloads"), ("autostat", "automation", "exports")),
        ("geo_autostat", (
            "GEO_AUTOSTAT_VERSION", "GEO_FINANCING_ACCOUNT_VERSION", "mapped_contact_source", "render_geo_autostat",
            "GEO_AUTOSTAT_CHANNEL_DOWNLOADS_VERSION",
            "GEO_AUTOSTAT_SHARED_DOWNLOADS_VERSION", "render_autostat_downloads",
            "AUTOSTAT_VISIBLE_DOWNLOADS_VERSION",
        ), ("autostat", "autostat_downloads", "automation", "email_validation", "folder_output")),
        ("b2_matching", (
            "B2_MATCHING_VERSION", "build_b2_exports", "DATAGRID_ROUTING_VERSION", "is_datagrid_sheet",
            "B2_TEMPLATE_VERSION", "B2_TEMPLATE_HEADERS",
            "B2_CONTACT_PRIORITY_VERSION",
            "B2_B4_CONTACT_HEADERS_VERSION",
            "B2_BUCKET_DIAGNOSTICS_VERSION",
            "B2_UPLOAD_ROLES_VERSION", "is_sbc_database_sheet", "validate_b2_source",
            "B4_MATCHING_VERSION", "build_b4_exports", "B2_B4_AGENTS_VERSION",
            "B2_B4_OB_SORT_VERSION",
            "B2_B4_AVAILABILITY_VERSION",
            "SPM_ENDORSEMENT_VERSION",
        ), ("autostat", "automation", "export_config")),
        ("b2_page", ("render_b2_page", "BROWSER_DOWNLOADS_ONLY_VERSION", "B4_PAGE_VERSION", "B2_COMPACT_DOWNLOADS_VERSION", "ZIP_FORMAT_VERSION", "B2_B4_CSV_PROTECTION_VERSION", "upload_widget_key", "render_agent_availability"), ("b2_matching", "autostat", "automation", "exports", "folder_output", "export_config", "upload_state", "agent_availability_page")),
        ("b2_autostat", (
            "B2_AUTOSTAT_VERSION", "B2_EMAIL_HEADERS", "B2_PHONE_HEADERS", "B2ActivitySource",
            "b2_metadata_columns", "build_b2_autostat_exports", "B2_PRIMARY_KEY_VERSION", "b2_account_sheet",
            "B2_COLLECTOR_TERMS_VERSION",
            "B2_B4_REMARK_DATE_VERSION",
            "B4_AUTOSTAT_VERSION", "build_b4_autostat_exports",
        ), ("autostat", "automation")),
        ("autostat_page", (
            "render_autostat", "AUTOSTAT_PAGE_VERSION", "COMPACT_SUMMARY_VERSION", "GROUPED_SOURCE_SETTINGS_VERSION",
            "B2_UPLOAD_VERSION", "DEDICATED_BUCKET_UPLOADS_VERSION", "upload_widget_key",
            "AUTOSTAT_AUTOMATIC_UPLOADS_VERSION", "BROWSER_DOWNLOADS_ONLY_VERSION",
            "AUTOSTAT_VISIBLE_SETTINGS_VERSION",
        ), ("autostat", "autostat_downloads", "b2_autostat", "automation", "folder_output", "export_config", "upload_state")),
    )
    for name, required, dependencies in modules:
        module = importlib.import_module(name)
        if any(not _has_required_name(module, attribute) for attribute in required) or refreshed.intersection(dependencies):
            module = importlib.reload(module)
            refreshed.add(name)
        missing = [attribute for attribute in required if not _has_required_name(module, attribute)]
        if missing:
            raise ImportError(
                f"{name} is missing required names: {', '.join(missing)}. "
                "Deploy all files from the latest GitHub commit and reboot the Streamlit app."
            )
    return refreshed
