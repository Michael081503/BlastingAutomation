"""Six-file upload workspace and the inbound reply report layout."""

from __future__ import annotations

from datetime import date, datetime, timedelta, timezone
from html import escape
from hashlib import sha256
from functools import partial

import streamlit as st

from inbound_reply import (
    VIBER_REPLY_VERSION, VIBER_AGENT_REPLIED_SOURCE, VIBER_MATCH_MODE, InboundReplyError, ViberReplyReport,
    build_people_csv, build_unreplied_workbook, match_viber_sources, normalized_campaign, people_rows, read_reply_source,
)
from inbound_email import (
    EMAIL_REPLY_VERSION, EmailReplyReport, build_email_unreplied_workbook,
    match_email_sources, read_email_source,
)
from upload_state import upload_widget_key


INBOUND_REPLY_PAGE_VERSION = 9
INBOUND_REPLY_TITLE = "INBOUND REPLY"
VIBER_MATCHING_PAGE_VERSION = 4
EMAIL_MATCHING_PAGE_VERSION = 1
MANUAL_CAMPAIGN_TOTALS_VERSION = 1
CHANNELS = ("SMS", "EMAIL", "VIBER")


def campaign_groups(viber_report=None, email_report=None) -> list[dict]:
    campaigns = {}
    for channel, report in (("viber", viber_report), ("email", email_report)):
        if report is not None:
            for counts in report.campaigns:
                group = campaigns.setdefault(normalized_campaign(counts.campaign), {"label": counts.campaign})
                group[channel] = counts
    return sorted(campaigns.values(), key=lambda group: group["label"].casefold())


def campaign_totals_key(report_date: date, campaign: str) -> str:
    return f"{report_date.isoformat()}|{normalized_campaign(campaign)}"


def _count(value) -> int | None:
    return value if isinstance(value, int) and not isinstance(value, bool) and value >= 0 else None


def render_campaign_totals(report_date: date, campaign: str, viber_report=None,
                           email_report=None, *, include_worked: bool = False) -> dict:
    """Keep manual totals with their campaign/date when uploads or visible rows change."""
    saved = {key: dict(value) for key, value in st.session_state.get("inbound_reply_campaign_totals", {}).items()}
    groups = campaign_groups(viber_report, email_report) or [{"label": campaign.strip() or "—"}]
    with st.expander("Edit campaign totals", expanded=True):
        st.caption(
            "Enter ACTIVE TICKET and the Email, SMS and Viber DC WORKED counts. "
            "DC WORKED percentages are calculated automatically from ACTIVE TICKET. "
            "Values are saved separately for each campaign and report date."
        )
        for group in groups:
            label = group["label"]
            identity = campaign_totals_key(report_date, label)
            values = dict(saved.get(identity, {}))
            widget_prefix = "inbound_reply_total_" + sha256(identity.encode("utf-8")).hexdigest()[:16]
            st.markdown(f"**{escape(label)}**")
            fields = [("active_ticket", "ACTIVE TICKET")]
            if include_worked:
                fields.extend((f"dc_{channel}", f"DC WORKED — {channel.upper()}") for channel in ("email", "sms", "viber"))
            for column, (field, field_label) in zip(st.columns(len(fields)), fields):
                values[field] = column.number_input(
                    field_label, min_value=0, value=_count(values.get(field)), step=1,
                    placeholder="Enter total", key=f"{widget_prefix}_{field}",
                )
            saved[identity] = values
    st.session_state["inbound_reply_campaign_totals"] = saved
    return saved


def build_report_html(campaign: str, report_date: date, sent_to_ops: str,
                      viber_report: ViberReplyReport | None = None,
                      email_report: EmailReplyReport | None = None,
                      campaign_totals: dict | None = None) -> str:
    """Fill each uploaded channel independently, joining campaign labels exactly."""
    date_label = f"{report_date.strftime('%B').upper()} {report_date.day}, {report_date.year} DATA"
    channel_headers = "".join(
        f'<th scope="col">{channel}</th><th scope="col">%</th>'
        for channel in ("EMAIL", "SMS", "VIBER")
    )
    body = []
    for group in campaign_groups(viber_report, email_report) or [None]:
        cells = [campaign.strip() or "—", *(["—"] * 10), sent_to_ops, *(["—"] * 6)]
        if group is not None:
            cells[0] = group["label"]
            for channel, indexes in (("email", (8, 12, 13)), ("viber", (10, 16, 17))):
                counts = group.get(channel)
                if counts is not None:
                    reply_index, agent_index, percent_index = indexes
                    cells[reply_index] = f"{counts.unique_people:,}"
                    cells[agent_index] = f"{counts.agent_replied:,}"
                    cells[percent_index] = f"{counts.agent_percent:.2%}" if counts.agent_percent is not None else "—"
        totals = (campaign_totals or {}).get(campaign_totals_key(report_date, cells[0]), {})
        active = _count(totals.get("active_ticket"))
        if active is not None:
            cells[1] = f"{active:,}"
        for channel, index in (("email", 2), ("sms", 4), ("viber", 6)):
            worked = _count(totals.get(f"dc_{channel}"))
            if worked is not None:
                cells[index] = f"{worked:,}"
                cells[index + 1] = f"{worked / active:.2%}" if active else "—"
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
      .inbound-campaign-report table {{ min-width: 1804px; table-layout: fixed; font-size: 11px; }}
      .inbound-campaign-report th, .inbound-campaign-report td {{ padding: 5px 6px; }}
    </style>
    <div class="inbound-reply-report inbound-campaign-report" role="region" aria-label="Inbound reply summary" tabindex="0">
      <table aria-label="Inbound reply campaign report">
        <colgroup><col style="width:264px">{''.join('<col style="width:88px">' for _ in range(10))}
        <col style="width:132px">{''.join('<col style="width:88px">' for _ in range(6))}</colgroup>
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
    signature = (VIBER_REPLY_VERSION, VIBER_AGENT_REPLIED_SOURCE, VIBER_MATCH_MODE,
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
                                         agent_replied_source=VIBER_AGENT_REPLIED_SOURCE, match_mode=VIBER_MATCH_MODE)
            cached = {"signature": signature, "report": report}
        except InboundReplyError as error:
            cached = {"signature": signature, "error": str(error)}
        st.session_state["inbound_reply_viber_result"] = cached
    if "error" in cached:
        st.error(cached["error"])
        return None
    return cached["report"]


def prepare_email_report(history_upload, dc_upload) -> EmailReplyReport | None:
    if history_upload is None or dc_upload is None:
        st.session_state.pop("inbound_reply_email_result", None)
        return None
    history_bytes, dc_bytes = history_upload.getvalue(), dc_upload.getvalue()
    signature = (EMAIL_REPLY_VERSION, sha256(history_bytes).hexdigest(), sha256(dc_bytes).hexdigest(),
                 history_upload.name, dc_upload.name)
    cached = st.session_state.get("inbound_reply_email_result")
    if cached is None or cached["signature"] != signature:
        st.session_state.pop("inbound_reply_email_result", None)
        try:
            if history_bytes == dc_bytes:
                raise InboundReplyError("The two Email uploads are identical. Upload Email History and DC Checking separately.")
            history = read_email_source(history_upload.name, history_bytes, kind="history")
            dc = read_email_source(dc_upload.name, dc_bytes, kind="dc")
            cached = {"signature": signature, "report": match_email_sources(dc, history)}
        except InboundReplyError as error:
            cached = {"signature": signature, "error": str(error)}
        st.session_state["inbound_reply_email_result"] = cached
    if "error" in cached:
        st.error(cached["error"])
        return None
    return cached["report"]


def render_email_details(report: EmailReplyReport, report_date: date) -> None:
    totals = report.totals
    for column, label, value in zip(st.columns(3),
                                    ("Unique email addresses", "Email replied by agents", "Email awaiting agent reply"),
                                    (totals.unique_people, totals.agent_replied, totals.awaiting_reply)):
        column.metric(label, f"{value:,}")
    st.caption(
        "DC Checking supplies the account list and campaigns. An address counts as replied when it matches "
        "an Email History row with a nonblank sent value. Matching and duplicate removal use email addresses."
    )
    st.subheader("Email unreplied accounts by campaign")
    rows = "".join(f"<tr><td>{escape(campaign.campaign)}</td><td>{campaign.awaiting_reply:,}</td></tr>"
                   for campaign in report.campaigns)
    st.html(
        '<div class="inbound-reply-report" style="max-width:640px" role="region" '
        'aria-label="Email unreplied accounts by campaign" tabindex="0">'
        '<table style="min-width:0"><thead><tr><th scope="col">Campaign</th>'
        f'<th scope="col">Not replied</th></tr></thead><tbody>{rows}</tbody></table></div>'
    )
    st.download_button(
        "Download Email unreplied report (.xlsx)", data=partial(build_email_unreplied_workbook, report),
        file_name=f"EMAIL_Unreplied_{report_date.isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="inbound_reply_email_unreplied_download", on_click="ignore", type="primary",
    )
    st.caption("The Excel workbook contains Summary and Unreplied Details sheets, retaining the original DC Checking fields.")
    with st.expander("Email matching details"):
        st.write(
            f"DC Checking: {len(report.dc.records):,} valid rows. Email History: {len(report.history.records):,} valid rows, "
            f"including {sum(row.sent for row in report.history.records):,} sent entries. "
            f"{report.history_only_addresses:,} history-only email addresses are outside the DC account list and are excluded."
        )
        st.caption(
            "Addresses are matched without case or surrounding whitespace. Received-only messages, agent names, "
            "assignment fields, AI suggestions and DC workflow labels do not mark an account as replied. "
            "History has no campaign column, so a sent reply matches every DC campaign containing that address. "
            "All uploaded dates are included; Report date labels the report."
        )
        excluded = report.dc.excluded_rows + report.history.excluded_rows
        if excluded:
            st.warning(f"Skipped {excluded:,} rows with blank or invalid email addresses.")


def render_viber_details(report: ViberReplyReport, report_date: date | None = None) -> None:
    totals = report.totals
    for column, label, value in zip(st.columns(4),
                                    ("Unique phone numbers", "Replied by agents", "Awaiting agent reply", "Found in both files"),
                                    (totals.unique_people, totals.agent_replied, totals.awaiting_reply, totals.matched_both_files)):
        column.metric(label, f"{value:,}")
    st.caption(
        "Viber Reply counts unique phone numbers across both files. Agent Worked counts phone numbers in campaign replies, "
        "divided by Viber Reply for the percentage. Each campaign counts a phone number once; "
        "the totals above count each phone number once across all campaigns."
    )
    st.subheader("Unreplied accounts by campaign")
    summary_rows = "".join(
        f"<tr><td>{escape(campaign.campaign)}</td><td>{campaign.awaiting_reply:,}</td></tr>"
        for campaign in report.campaigns
    )
    st.html(
        '<div class="inbound-reply-report" style="max-width:640px" role="region" '
        'aria-label="Unreplied accounts by campaign" tabindex="0">'
        '<table style="min-width:0"><thead><tr><th scope="col">Campaign</th>'
        f'<th scope="col">Not replied</th></tr></thead><tbody>{summary_rows}</tbody></table></div>'
    )
    st.caption(
        "The Excel download has a Summary sheet and an Unreplied Details sheet with the original source fields. "
        "Accounts are counted once per phone number within each campaign."
    )
    st.download_button(
        "Download unreplied report (.xlsx)", data=partial(build_unreplied_workbook, report),
        file_name=f"VIBER_Unreplied_{report_date.isoformat()}.xlsx" if report_date else "VIBER_Unreplied.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="inbound_reply_viber_unreplied_download", on_click="ignore", type="primary",
    )
    st.download_button(
        "Download matched Viber phone numbers (.csv)", data=build_people_csv(report),
        file_name="VIBER_Inbound_reply_phones.csv", mime="text/csv",
        key="inbound_reply_viber_download", on_click="ignore",
    )
    with st.expander("Viber matching details"):
        st.write(
            f"Message logs: {len(report.message_logs.records):,} rows. "
            f"Campaign replies: {len(report.campaign_replies.records):,} rows. "
            f"Combined into {totals.unique_people:,} unique phone numbers."
        )
        st.caption(
            "Matching uses phone numbers, normalizing spaces, punctuation and Philippine +63/0 formats. "
            "Names are shown for reference. Rows without a usable phone number are excluded and counted. "
            "An agent reply takes priority over pending entries for the same phone number and campaign. "
            "Campaign replies identify phone numbers replied to by agents; numbers found only in message logs await a reply. "
            "All uploaded rows are included; Report date labels the report."
        )
        excluded = report.message_logs.excluded_rows + report.campaign_replies.excluded_rows + report.excluded_phone_rows
        if excluded:
            st.warning(f"Skipped {excluded:,} rows without a usable phone number.")
        rows = people_rows(report)
        headers = list(rows[0]) if rows else []
        header = "".join(f"<th scope='col'>{escape(key)}</th>" for key in headers)
        body = "".join("<tr>" + "".join(f"<td>{escape(str(row[key]))}</td>" for key in headers) + "</tr>"
                       for row in rows[:100])
        if rows:
            st.html(
                '<div class="inbound-reply-report" role="region" aria-label="Matched Viber phone numbers" tabindex="0">'
                f'<table><thead><tr>{header}</tr></thead><tbody>{body}</tbody></table></div>'
            )
            st.caption(f"Showing {min(len(rows), 100):,} of {len(rows):,} phone numbers. Download the CSV for all numbers.")


def render_inbound_reply() -> None:
    st.title("INBOUND REPLY")
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
    email_uploads = []
    for column, channel in zip(st.columns(3), CHANNELS):
        with column:
            st.markdown(f"**{channel}**")
            channel_count = 0
            for number in (1, 2):
                label = f"{channel} file {number}"
                if channel == "VIBER":
                    label = "Viber message logs — comparison list" if number == 1 else "Viber campaign replies — replied by agents"
                elif channel == "EMAIL":
                    label = "Email History — sent agent replies" if number == 1 else "Email DC Checking — account list"
                upload = st.file_uploader(
                    label, type=["xlsx", "csv"],
                    key=upload_widget_key(f"inbound_reply_{channel.lower()}_{number}_upload"),
                )
                channel_count += upload is not None
                if channel == "VIBER":
                    viber_uploads.append(upload)
                elif channel == "EMAIL":
                    email_uploads.append(upload)
            uploaded_count += channel_count
            st.caption(f"{channel_count} of 2 files uploaded")
    st.caption(f"{uploaded_count} of 6 files uploaded")
    with report_area:
        st.subheader("Report")
        if campaign.strip():
            st.caption(campaign.strip())
        viber_report = prepare_viber_report(*viber_uploads)
        email_report = prepare_email_report(*email_uploads)
        table_area = st.container()
        manual_totals = render_campaign_totals(report_date, campaign, viber_report, email_report, include_worked=True)
        with table_area:
            st.html(build_report_html(campaign, report_date, sent_to_ops, viber_report, email_report, manual_totals))
        if email_report is not None:
            st.subheader("Email")
            render_email_details(email_report, report_date)
        if viber_report is not None:
            st.subheader("Viber")
            render_viber_details(viber_report, report_date)
        if viber_report is not None or email_report is not None:
            st.caption("ACTIVE TICKET and DC WORKED counts are entered manually; percentages are automatic. SMS reply processing is pending its files and rules.")
        else:
            st.caption("Upload both files for Email or Viber to generate that channel's report. SMS rules are pending.")
