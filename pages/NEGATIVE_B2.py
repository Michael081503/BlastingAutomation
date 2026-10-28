"""Dedicated page for B2 blasting templates."""

import streamlit as st

from b2_page import render_b2_page
from branding import render_bank_branding

render_bank_branding()
st.title("NEGATIVE B2")
render_b2_page()
