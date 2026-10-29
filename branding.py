"""Shared Security Bank presentation for the three application pages."""

from __future__ import annotations

import re
from base64 import b64encode
from functools import lru_cache
from pathlib import Path

import streamlit as st


ASSETS_DIRECTORY = Path(__file__).resolve().parent / "assets"
BANK_LOGO = ASSETS_DIRECTORY / "security-bank-logo.png"
BANK_ICON = ASSETS_DIRECTORY / "security-bank-icon.png"
HALLOWEEN_THEME_VERSION = 1
HALLOWEEN_GHOST_VERSION = 1
HALLOWEEN_SPIDER_VERSION = 1


@lru_cache(maxsize=8)
def _asset_data(filename: str) -> str:
    return b64encode((ASSETS_DIRECTORY / filename).read_bytes()).decode("ascii")


@lru_cache(maxsize=1)
def _still_sidebar_data() -> str:
    illustration = (ASSETS_DIRECTORY / "halloween-sidebar.svg").read_text(encoding="utf-8")
    illustration = re.sub(r"<style\b[^>]*>.*?</style>", "", illustration, flags=re.DOTALL)
    return b64encode(illustration.encode("utf-8")).decode("ascii")


def render_bank_branding() -> None:
    st.logo(str(BANK_LOGO), size="large", icon_image=str(BANK_ICON))
    styles = (ASSETS_DIRECTORY / "security-bank.css").read_text(encoding="utf-8")
    styles = styles.replace("__HALLOWEEN_SIDEBAR__", _asset_data("halloween-sidebar.svg"))
    styles = styles.replace("__HALLOWEEN_SIDEBAR_STILL__", _still_sidebar_data())
    styles = styles.replace("__HALLOWEEN_WEBS__", _asset_data("halloween-webs.svg"))
    st.html(f"<style>{styles}</style>")
    st.html(
        '<div class="sbc-brand-header">'
        '<svg class="sbc-cobweb" viewBox="0 0 150 140" aria-hidden="true">'
        '<g fill="none" stroke="currentColor" stroke-width="1">'
        '<path d="M150 0L0 0M150 0L0 140M150 0L70 140M150 0L150 140M150 0L0 65"/>'
        '<path d="M115 0Q116 15 126 21Q135 31 150 32M80 0Q83 29 103 43Q123 62 150 64'
        'M44 0Q49 42 80 65Q109 93 150 96M8 0Q15 57 56 88Q95 123 150 128"/>'
        '</g></svg>'
        '<div class="sbc-bank-mark">'
        f'<img src="data:image/png;base64,{_asset_data("security-bank-logo.png")}" alt="Security Bank" class="sbc-brand-logo">'
        '</div>'
        '<div class="sbc-header-mascot" aria-hidden="true">'
        f'<img src="data:image/svg+xml;base64,{_asset_data("halloween-spiderman.svg")}" alt="" class="sbc-spiderman">'
        '</div><div class="sbc-workspace-meta">'
        '<span class="sbc-workspace-name">Blasting Automation</span>'
        '<span class="sbc-season-tag"><span aria-hidden="true">&#127875;</span> Halloween edition</span>'
        '</div>'
        '</div>'
    )
