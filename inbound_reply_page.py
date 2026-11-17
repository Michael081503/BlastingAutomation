"""Six-file upload workspace and the inbound reply report layout."""

from __future__ import annotations

from datetime import date, datetime, timedelta, timezone
from html import escape
from hashlib import sha256

import streamlit as st

from inbound_reply import (
    VIBER_REPLY_VERSION, VIBER_AGENT_REPLIED_SOURCE, InboundReplyError, ViberReplyReport,
    build_people_csv, match_viber_sources, people_rows, read_reply_source,
)
from upload_state import upload_widget_key


INBOUND_REPLY_PAGE_VERSION = 3
VIBER_MATCHING_PAGE_VERSION = 2
CHANNELS = ("SMS", "EMAIL", "VIBER")


def build_report_html(campaign: str, report_date: date, sent_to_ops: str,
                      viber_report: ViberReplyReport | None = None) -> str:
    """Fill Viber reply columns; leave unprovided channel and ticket totals blank."""
    date_label = f"{report_date.strftime('%B').upper()} {report_date.day}, {report_date.year} DATA"
    channel_headers = "".join(
        f'<th scope="col">{channel}</th><th scope="col">%</th>'
        for channel in ("EMAIL", "SMS", "VIBER")
    )
    body = []
    for counts in (viber_report.campaigns if viber_report and viber_report.campaigns else (None,)):
        cells = [campaign.strip() or "—", *(["—"] * 10), sent_to_ops, *(["—"] * 6)]
        if counts is not None:
            cells[0] = counts.campaign
            cells[10] = f"{counts.unique_people:,}"
            cells[16] = f"{counts.agent_replied:,}"
            cells[17] = f"{counts.agent_percent:.2%}" if counts.agent_percent is not None else "—"
        body.append("<tr>" + "".join(f"<td>{escape(value)}</td>" for value in cells) + "</tr>")
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
        <tbody>{''.join(body)}</tbody>
      </table>
    </div>
    """


def prepare_viber_report(log_upload, campaign_upload) -> ViberReplyReport | None:
    """Keep prepared files session-local and invalidate on upload content changes."""
    if log_upload is None or campaign_upload is None:
        st.session_state.pop("inbound_reply_viber_result", None)
        return None
    log_bytes, campaign_bytes = log_upload.getvalue(), campaign_upload.getvalue()
    signature = (VIBER_REPLY_VERSION, VIBER_AGENT_REPLIED_SOURCE,
                 sha256(log_bytes).hexdigest(), sha256(campaign_bytes).hexdigest(),
                 log_upload.name, campaign_upload.name)
    cached = st.session_state.get("inbound_reply_viber_result")
    if cached is None or cached["signature"] != signature:
        st.session_state.pop("inbound_reply_viber_result", None)
        try:
            if log_bytes == campaign_bytes:
                raise InboundReplyError("The two Viber uploads are identical. Upload message logs and campaign replies separately.")
            logs = read_reply_source(log_upload.name, log_bytes)
            campaign_replies = read_reply_source(campaign_upload.name, campaign_bytes)
            # The user's corrected file roles determine status; Message Type is source data.
            report = match_viber_sources(logs, campaign_replies,
                                         agent_replied_source=VIBER_AGENT_REPLIED_SOURCE, match_mode="name")
            cached = {"signature": signature, "report": report}
        except InboundReplyError as error:
            cached = {"signature": signature, "error": str(error)}
        st.session_state["inbound_reply_viber_result"] = cached
    if "error" in cached:
        st.error(cached["error"])
        return None
    return cached["report"]


def render_viber_details(report: ViberReplyReport) -> None:
    totals = report.totals
    for column, label, value in zip(st.columns(4),
                                    ("Unique people", "Replied by agents", "Awaiting agent reply", "Found in both files"),
                                    (totals.unique_people, totals.agent_replied, totals.awaiting_reply, totals.matched_both_files)):
        column.metric(label, f"{value:,}")
    st.caption(
        "Viber Reply counts unique people across both files. Agent Worked counts people in campaign replies, "
        "divided by Viber Reply for the percentage. Each campaign counts a person once; "
        "the totals above count each person once across all campaigns."
    )
    st.download_button(
        "Download matched Viber people (.csv)", data=build_people_csv(report),
        file_name="VIBER_Inbound_reply_people.csv", mime="text/csv",
        key="inbound_reply_viber_download", on_click="ignore",
    )
    with st.expander("Viber matching details"):
        st.write(
            f"Message logs: {len(report.message_logs.records):,} rows. "
            f"Campaign replies: {len(report.campaign_replies.records):,} rows. "
            f"Combined into {totals.unique_people:,} unique people."
        )
        st.caption(
            "Matching combines identical full names, ignoring case, punctuation and extra spaces, "
            "even when phone numbers differ. A missing name uses the phone number when available. "
            "An agent reply takes priority over pending entries for the same person and campaign. "
            "Campaign replies identify people replied to by agents; people found only in message logs await a reply. "
            "All uploaded rows are included; Report date labels the report."
        )
        excluded = report.message_logs.excluded_rows + report.campaign_replies.excluded_rows
        if excluded:
            st.warning(f"Skipped {excluded:,} rows without a usable name or phone number.")
        rows = people_rows(report)
        headers = list(rows[0]) if rows else []
        header = "".join(f"<th scope='col'>{escape(key)}</th>" for key in headers)
        body = "".join("<tr>" + "".join(f"<td>{escape(str(row[key]))}</td>" for key in headers) + "</tr>"
                       for row in rows[:100])
        if rows:
            st.html(
                '<div class="inbound-reply-report" role="region" aria-label="Matched Viber people" tabindex="0">'
                f'<table><thead><tr>{header}</tr></thead><tbody>{body}</tbody></table></div>'
            )
            st.caption(f"Showing {min(len(rows), 100):,} of {len(rows):,} people. Download the CSV for all people.")


def render_inbound_reply() -> None:
    st.title("Inbound reply")
    st.caption("Upload six files: two each for SMS, Email, and Viber.")

    campaign_column, date_column, ops_column = st.columns([2, 1, 1])
    with campaign_column:
        campaign = st.text_input("Report label (optional)", placeholder="Report name", key="inbound_reply_campaign")
    with date_column:
        report_date = st.date_input(
            "Report date", value=datetime.now(timezone(timedelta(hours=8))).date(),
            key="inbound_reply_date",
        )
    with ops_column:
        sent_to_ops = st.selectbox(
            "Sent to OPS", ("—", "YES", "NO"), key="inbound_reply_sent_to_ops",
        )

    report_area = st.container()

    st.subheader("Upload files")
    uploaded_count = 0
    viber_uploads = []
    for column, channel in zip(st.columns(3), CHANNELS):
        with column:
            st.markdown(f"**{channel}**")
            channel_count = 0
            for number in (1, 2):
                label = f"{channel} file {number}"
                if channel == "VIBER":
                    label = "Viber message logs — comparison list" if number == 1 else "Viber campaign replies — replied by agents"
                upload = st.file_uploader(
                    label, type=["xlsx", "csv"],
                    key=upload_widget_key(f"inbound_reply_{channel.lower()}_{number}_upload"),
                )
                channel_count += upload is not None
                if channel == "VIBER":
                    viber_uploads.append(upload)
            uploaded_count += channel_count
            st.caption(f"{channel_count} of 2 files uploaded")
    st.caption(f"{uploaded_count} of 6 files uploaded")
    with report_area:
        st.subheader("Report")
        if campaign.strip():
            st.caption(campaign.strip())
        report = prepare_viber_report(*viber_uploads)
        st.html(build_report_html(campaign, report_date, sent_to_ops, report))
        if report is not None:
            render_viber_details(report)
            st.caption("Active Ticket, DC Worked, SMS and Email totals are pending their source data and rules.")
        else:
            st.caption("Upload both Viber files to show matched, deduplicated reply totals. SMS and Email rules are pending.")
