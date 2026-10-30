"""Dedicated page for B2 blasting templates."""

import streamlit as st

from automation import CSV_OUTPUT_ENCODINGS
from autostat_page import render_b2_autostat_uploads
from b2_page import render_b2_page
from branding import render_bank_branding

render_bank_branding()
st.title("NEGATIVE B2")
with st.sidebar:
    csv_file_type = st.selectbox("CSV file type", list(CSV_OUTPUT_ENCODINGS), key="b2_csv_type")
encoding = CSV_OUTPUT_ENCODINGS[csv_file_type]
render_b2_page(encoding=encoding)
st.divider()
render_b2_page(encoding=encoding, bucket=4)
with st.expander("B2 AutoStat — Email, SMS and Viber"):
    st.subheader("B2 files")
    render_b2_autostat_uploads(encoding=encoding)
