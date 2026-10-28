"""Shared Security Bank presentation for the three application pages."""

from __future__ import annotations

from base64 import b64encode
from functools import lru_cache
from pathlib import Path

import streamlit as st


ASSETS_DIRECTORY = Path(__file__).resolve().parent / "assets"
BANK_LOGO = ASSETS_DIRECTORY / "security-bank-logo.png"
BANK_ICON = ASSETS_DIRECTORY / "security-bank-icon.png"


@lru_cache(maxsize=1)
def _logo_data() -> str:
    return b64encode(BANK_LOGO.read_bytes()).decode("ascii")


def render_bank_branding() -> None:
    st.logo(str(BANK_LOGO), size="large", icon_image=str(BANK_ICON))
    st.html(f"<style>{(ASSETS_DIRECTORY / 'security-bank.css').read_text(encoding='utf-8')}</style>")
    st.html(
        '<div class="sbc-brand-header">'
        f'<img src="data:image/png;base64,{_logo_data()}" alt="Security Bank" class="sbc-brand-logo">'
        '<span class="sbc-workspace-name">Blasting Automation</span>'
        '</div>'
    )
