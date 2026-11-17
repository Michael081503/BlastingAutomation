"""Six-file upload workspace and the inbound reply report layout."""

from __future__ import annotations

from datetime import date, datetime, timedelta, timezone
from html import escape

import streamlit as st

from upload_state import upload_widget_key


INBOUND_REPLY_PAGE_VERSION = 1
CHANNELS = ("SMS", "EMAIL", "VIBER")


def build_report_html(campaign: str, report_date: date, sent_to_ops: str) -> str:
    """Display the requested grouped headers without inventing report totals."""
    date_label = f"{report_date.strftime('%B').upper()} {report_date.day}, {report_date.year} DATA"
    channel_headers = "".join(
        f'<th scope="col">{channel}</th><th scope="col">%</th>'
        for channel in ("EMAIL", "SMS", "VIBER")
    )
    cells = [campaign.strip() or "—", *(["—"] * 10), sent_to_ops, *(["—"] * 6)]
    body = "".join(f"<td>{escape(value)}</td>" for value in cells)
    return f"""
    <style>
      .inbound-reply-report {{
        width: 100%; overflow-x: auto; border: 1px solid #303744;
        border-radius: 4px; background: #fff;
      }}
      .inbound-reply-report:focus-visible {{ outline: 2px solid #f6923f; outline-offset: 3px; }}
      .inbound-reply-report table {{
        width: 100%; min-width: 1510px; border-collapse: collapse;
        color: #242831; font-family: Arial, sans-serif; font-size: 12px;
      }}
      .inbound-reply-report th, .inbound-reply-report td {{
        border: 1px solid #303744; padding: 9px 8px;
        text-align: center; vertical-align: middle; white-space: nowrap;
      }}
      .inbound-reply-report th {{ background: #e2eafb; font-weight: 700; }}
      .inbound-reply-report td {{ background: #fff; font-variant-numeric: tabular-nums; }}
      .inbound-reply-report td:first-child {{ min-width: 155px; white-space: normal; overflow-wrap: anywhere; }}
    </style>
    <div class="inbound-reply-report" role="region" aria-label="Inbound reply summary" tabindex="0">
      <table aria-label="Inbound reply campaign report">
        <thead>
          <tr>
            <th colspan="11" scope="colgroup">{date_label}</th>
            <th rowspan="3" scope="col">SENT TO OPS</th>
            <th colspan="6" rowspan="2" scope="colgroup">AGENT WORKED</th>
          </tr>
          <tr>
            <th rowspan="2" scope="col">CAMPAIGN</th>
            <th rowspan="2" scope="col">ACTIVE TICKET</th>
            <th colspan="6" scope="colgroup">DC WORKED</th>
            <th colspan="3" scope="colgroup">UNIQUE RESULT</th>
          </tr>
          <tr>
            {channel_headers}
            <th scope="col">EMAIL REPLY</th>
            <th scope="col">SMS REPLY</th>
            <th scope="col">VIBER REPLY</th>
            {channel_headers}
          </tr>
        </thead>
        <tbody><tr>{body}</tr></tbody>
      </table>
    </div>
    """


def render_inbound_reply() -> None:
    st.title("Inbound reply")
    st.caption("Upload six files: two each for SMS, Email, and Viber.")

    campaign_column, date_column, ops_column = st.columns([2, 1, 1])
    with campaign_column:
        campaign = st.text_input("Campaign", placeholder="Campaign name", key="inbound_reply_campaign")
    with date_column:
        report_date = st.date_input(
            "Report date", value=datetime.now(timezone(timedelta(hours=8))).date(),
            key="inbound_reply_date",
        )
    with ops_column:
        sent_to_ops = st.selectbox(
            "Sent to OPS", ("—", "YES", "NO"), key="inbound_reply_sent_to_ops",
        )

    st.subheader("Report")
    st.html(build_report_html(campaign, report_date, sent_to_ops))
    st.caption("Totals are pending your instructions about the uploaded files.")

    st.subheader("Upload files")
    uploaded_count = 0
    for column, channel in zip(st.columns(3), CHANNELS):
        with column:
            st.markdown(f"**{channel}**")
            channel_count = 0
            for number in (1, 2):
                upload = st.file_uploader(
                    f"{channel} file {number}", type=["xlsx", "csv"],
                    key=upload_widget_key(f"inbound_reply_{channel.lower()}_{number}_upload"),
                )
                channel_count += upload is not None
            uploaded_count += channel_count
            st.caption(f"{channel_count} of 2 files uploaded")
    st.caption(f"{uploaded_count} of 6 files uploaded")
    if uploaded_count == 6:
        st.info("All six files are uploaded. Report totals are pending your file instructions.")
