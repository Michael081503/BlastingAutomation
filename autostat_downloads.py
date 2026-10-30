"""Shared activity-log downloads for GEO/RECO and the AUTOSTAT upload page."""

from __future__ import annotations

from functools import partial
from pathlib import Path

import streamlit as st

from automation import XLSX_MIME
from autostat import (
    AutostatBundle, build_autostat_channel_excel, build_autostat_excel, build_autostat_excel_zip,
)


AUTOSTAT_DOWNLOADS_VERSION = 1


def render_autostat_downloads(
    bundle: AutostatBundle, *, state_prefix: str, workbook_name: str = "AUTOSTAT",
    excel_output: bool = True, show_counts: bool = True, compact: bool = False,
    channel_keys: tuple[str, ...] | None = None,
) -> None:
    """Offer separate channel files first, with ZIP and combined Excel alternatives."""
    channels = [Path(file.filename).stem.rsplit("_", 1)[-1] for file in bundle.files]

    def combined():
        st.download_button(
            f"Download {workbook_name}.xlsx" if compact else
            f"Download {workbook_name}.xlsx ({' + '.join(channels)} sheets)",
            partial(build_autostat_excel, bundle), f"{workbook_name}.xlsx",
            mime=XLSX_MIME, on_click="ignore", key=f"{state_prefix}excel",
            type="primary" if compact else "secondary",
        )

    def archive():
        st.download_button(
            f"Download {'both' if len(bundle.files) == 2 else 'all three'} AUTOSTAT files (ZIP)",
            partial(build_autostat_excel_zip, bundle) if excel_output else bundle.zip_data,
            f"{workbook_name}_FILES.zip", mime="application/zip", on_click="ignore",
            key=f"{state_prefix}zip",
        )

    if compact:
        combined()
        archive()
    columns = [None] * len(bundle.files) if compact else st.columns(len(bundle.files))
    for index, (column, channel, file) in enumerate(zip(columns, channels, bundle.files)):
        with st.container() if compact else column:
            if show_counts:
                st.caption(f"{channel}: {file.count:,} valid accounts")
            filename = Path(file.filename).with_suffix(".xlsx").name if excel_output else file.filename
            st.download_button(
                f"Download {filename}", partial(build_autostat_channel_excel, file) if excel_output else file.data,
                filename, mime=XLSX_MIME if excel_output else "text/csv", on_click="ignore",
                key=channel_keys[index] if channel_keys else f"{state_prefix}download_{index}",
                type="secondary" if compact else "primary",
            )
    if not compact:
        archive()
        combined()
