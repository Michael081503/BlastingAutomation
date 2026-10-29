"""Refresh cached application modules after a Streamlit code update."""

from __future__ import annotations

import importlib


def _has_required_name(module: object, name: str) -> bool:
    value = module
    try:
        for part in name.split("."):
            value = value[part] if isinstance(value, dict) else getattr(value, part)
    except (AttributeError, KeyError):
        return False
    return True


def refresh_application_modules() -> set[str]:
    """Reload missing interfaces and their dependents in dependency order.

    A running cloud process can retain modules from before a deployment. Only
    refresh modules when required names are missing, so routine reruns keep
    exception and dataclass identities stable.
    """
    modules = (
        ("upload_state", ("CLEAR_UPLOADS_VERSION", "upload_widget_key", "clear_all_files", "render_clear_all_files"), ()),
        ("branding", (
            "render_bank_branding", "BANK_ICON", "HALLOWEEN_THEME_VERSION", "HALLOWEEN_GHOST_VERSION", "HALLOWEEN_SPIDER_VERSION",
            "render_clear_all_files",
        ), ("upload_state",)),
        ("email_validation", ("EMAIL_VALIDATION_VERSION", "validated_email"), ()),
        ("export_config", (
            "CHANNEL_SUFFIXES", "OUTPUT_FILENAMES", "OUTPUT_FOLDER", "ZIP_FILENAME",
            "GEO_OUTPUT_FILENAMES", "GEO_ZIP_FILENAME", "AUTOSTAT_BEL_FILENAMES",
            "AUTOSTAT_B2_FILENAMES",
            "GEO_OUTPUT_FILENAMES.mindanao_qualified_cards", "GEO_OUTPUT_FILENAMES.r3_qualified_cards",
        ), ()),
        ("automation", (
            "MAPPING_VERSION", "OPTIONAL_FIELD_MAPPING", "CSV_OUTPUT_ENCODINGS",
            "create_default_csv_template",
            "DERIVED_TEMPLATE_HEADERS", "geo_doclass",
            "SheetInfo.available_fields", "inspect_source_sheet",
            "SheetInfo.headers", "source_column_positions", "preview_source_rows",
            "WORKBOOK_READER_VERSION",
            "OB_SORT_VERSION", "ob_sort_key",
            "write_text_csv_sheet",
            "MappedData.valid_email_rows", "MappedData.missing_email_rows",
        ), ("email_validation",)),
        ("exports", (
            "GEO_GROUP_LABELS", "GEO_PRODUCT_GROUP_LABELS", "GEO_REVIEW_GROUP_KEYS", "build_geo_exports",
            "GEO_FILENAME_CONFIG_VERSION",
            "GROUPED_SHEET_EXPORT_VERSION",
            "EXCEL_EXPORT_VERSION", "build_export_excel", "build_excel_zip",
        ), ("email_validation", "export_config", "automation")),
        ("folder_output", ("save_files_to_folder", "FOLDER_OUTPUT_VERSION"), ("automation", "exports")),
        ("autostat", (
            "AUTOSTAT_VERSION", "PREFERRED_ACCOUNT_HEADER", "EMAIL_HEADERS", "VIBER_HEADERS", "build_autostat_exports",
            "build_bel_exports", "has_bel_primary_keys",
            "ACCOUNT_DEDUP_VERSION", "ACCOUNT_IDENTITY_VERSION", "unique_account_contacts",
            "COMPILED_OUTPUT_VERSION", "BLANK_TIMESTAMP_VERSION",
            "AUTOSTAT_EXCLUSIONS_VERSION", "autostat_source_excluded",
            "CONTACT_VALIDATION_VERSION", "_validated_candidates",
            "ACCOUNT_TEXT_EXPORT_VERSION", "build_autostat_excel",
            "ACCOUNT_NUMBER_NORMALIZATION_VERSION", "account_number_text",
            "ACCOUNT_INTEGRITY_VERSION", "EXCEL_CHANNEL_OUTPUT_VERSION",
            "build_autostat_channel_excel", "build_autostat_excel_zip",
            "ContactRows.__dataclass_fields__.email_ob_values", "ContactRows.__dataclass_fields__.phone_ob_values",
        ), ("automation", "email_validation", "export_config", "exports")),
        ("geo_autostat", (
            "GEO_AUTOSTAT_VERSION", "GEO_PRIMARY_KEY_VERSION", "mapped_contact_source", "render_geo_autostat",
        ), ("autostat", "automation", "email_validation", "folder_output")),
        ("b2_matching", (
            "B2_MATCHING_VERSION", "build_b2_exports", "DATAGRID_ROUTING_VERSION", "is_datagrid_sheet",
            "B2_TEMPLATE_VERSION", "B2_TEMPLATE_HEADERS",
            "B2_CONTACT_PRIORITY_VERSION",
            "B2_UPLOAD_ROLES_VERSION", "is_sbc_database_sheet", "validate_b2_source",
        ), ("autostat", "automation", "export_config")),
        ("b2_page", ("render_b2_page", "upload_widget_key"), ("b2_matching", "autostat", "automation", "folder_output", "export_config", "upload_state")),
        ("b2_autostat", (
            "B2_AUTOSTAT_VERSION", "B2_EMAIL_HEADERS", "B2_PHONE_HEADERS", "B2ActivitySource",
            "b2_metadata_columns", "build_b2_autostat_exports", "B2_PRIMARY_KEY_VERSION", "b2_account_sheet",
        ), ("autostat", "automation")),
        ("autostat_page", (
            "render_autostat", "AUTOSTAT_PAGE_VERSION", "COMPACT_SUMMARY_VERSION", "GROUPED_SOURCE_SETTINGS_VERSION",
            "B2_UPLOAD_VERSION", "B2_ACTIVITY_PANEL_VERSION", "render_b2_autostat_uploads", "upload_widget_key",
        ), ("autostat", "b2_autostat", "automation", "folder_output", "export_config", "upload_state")),
    )
    refreshed = set()
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
