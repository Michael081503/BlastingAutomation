"""Shared Security Bank presentation for the application pages."""

from __future__ import annotations

from base64 import b64encode
from functools import lru_cache
from pathlib import Path

import streamlit as st

from upload_state import render_clear_all_files


ASSETS_DIRECTORY = Path(__file__).resolve().parent / "assets"
BANK_LOGO = ASSETS_DIRECTORY / "security-bank-logo.png"
BANK_ICON = ASSETS_DIRECTORY / "security-bank-icon.png"
CHRISTMAS_THEME_VERSION = 1
SNOWFALL_VERSION = 1
CHRISTMAS_SPIDERMAN_VERSION = 1


@lru_cache(maxsize=8)
def _asset_data(filename: str) -> str:
    return b64encode((ASSETS_DIRECTORY / filename).read_bytes()).decode("ascii")


def render_bank_branding() -> None:
    st.logo(str(BANK_LOGO), size="large", icon_image=str(BANK_ICON))
    styles = (ASSETS_DIRECTORY / "security-bank.css").read_text(encoding="utf-8")
    styles = styles.replace("__WINTER_SCENE__", _asset_data("christmas-winter.svg"))
    styles = styles.replace("__CHRISTMAS_GARLAND__", _asset_data("christmas-garland.svg"))
    styles = styles.replace("__SNOWFLAKE_ICON__", _asset_data("christmas-snowflake.svg"))
    st.html(f"<style>{styles}</style>")
    with st.container(key="sbc_workspace_header", horizontal=True, vertical_alignment="center", gap="large"):
        st.html(
            '<div class="sbc-brand-header">'
            '<div class="sbc-bank-mark">'
            f'<img src="data:image/png;base64,{_asset_data("security-bank-logo.png")}" alt="Security Bank" class="sbc-brand-logo">'
            '</div>'
            '<div class="sbc-header-mascot" aria-hidden="true">'
            f'<img src="data:image/svg+xml;base64,{_asset_data("christmas-spiderman.svg")}" alt="" class="sbc-spiderman">'
            '</div>'
            '</div>'
        )
        with st.container(key="sbc_workspace_actions", width="content", gap="small"):
            st.html(
                '<div class="sbc-workspace-meta">'
                '<span class="sbc-workspace-name">Blasting Automation</span>'
                '<span class="sbc-season-tag"><span aria-hidden="true">❄</span> Christmas season</span>'
                '</div>'
            )
            render_clear_all_files()
