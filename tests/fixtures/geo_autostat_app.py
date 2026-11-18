"""Run the embedded GEO/RECO AUTOSTAT section with a test's selected worksheet."""

import streamlit as st

from geo_autostat import render_geo_autostat


render_geo_autostat(
    st.session_state["source_name"], st.session_state["source_data"], "NEGATIVE", 1,
    st.session_state["header_count"],
    {"primary_key": 0, "cardnumber": 1, "indexoneuid": 2, "email": 3, "phone": 4, "OB": 5},
    encoding="cp1252",
)
