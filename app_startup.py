"""Refresh cached application modules after a Streamlit code update."""

from __future__ import annotations

import importlib


def refresh_application_modules() -> set[str]:
    """Reload missing interfaces and their dependents in dependency order.

    A running cloud process can retain modules from before a deployment. Only
    refresh modules when required names are missing, so routine reruns keep
    exception and dataclass identities stable.
    """
    modules = (
        ("email_validation", ("EMAIL_VALIDATION_VERSION", "validated_email"), ()),
        ("export_config", (
            "CHANNEL_SUFFIXES", "OUTPUT_FILENAMES", "OUTPUT_FOLDER", "ZIP_FILENAME",
            "GEO_AREA_SUFFIXES", "GEO_OUTPUT_FILENAMES", "GEO_ZIP_FILENAME",
        ), ()),
        ("automation", (
            "MAPPING_VERSION", "OPTIONAL_FIELD_MAPPING", "CSV_OUTPUT_ENCODINGS",
            "create_default_csv_template",
        ), ("email_validation",)),
        ("exports", (
            "GEO_GROUP_LABELS", "GEO_REVIEW_GROUP_KEYS", "build_geo_exports",
        ), ("email_validation", "export_config", "automation")),
        ("folder_output", ("save_files_to_folder",), ("automation", "exports")),
    )
    refreshed = set()
    for name, required, dependencies in modules:
        module = importlib.import_module(name)
        if any(not hasattr(module, attribute) for attribute in required) or refreshed.intersection(dependencies):
            module = importlib.reload(module)
            refreshed.add(name)
        missing = [attribute for attribute in required if not hasattr(module, attribute)]
        if missing:
            raise ImportError(
                f"{name} is missing required names: {', '.join(missing)}. "
                "Deploy all files from the latest GitHub commit and reboot the Streamlit app."
            )
    return refreshed
