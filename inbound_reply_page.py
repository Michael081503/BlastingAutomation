"""Six-file upload workspace and the inbound reply report layout."""

from __future__ import annotations

from datetime import date, datetime, timedelta, timezone
from html import escape
from hashlib import sha256
from functools import partial

import streamlit as st

from inbound_reply import (
    VIBER_REPLY_VERSION, VIBER_AGENT_REPLIED_SOURCE, VIBER_MATCH_MODE, InboundReplyError, ViberReplyReport,
    build_people_csv, build_unreplied_workbook, match_viber_sources, people_rows, read_reply_source,
)
from inbound_email import (
    EMAIL_REPLY_VERSION, EmailReplyReport, build_email_unreplied_workbook,
    match_email_sources, read_email_source,
)
from upload_state import upload_widget_key
from inbound_summary import (
    build_summary_workbook, campaign_groups, campaign_totals_key,
    display_cell, nonnegative_count as _count, summary_rows,
)


INBOUND_REPLY_PAGE_VERSION = 14
INBOUND_REPLY_TITLE = "INBOUND REPLY"
VIBER_MATCHING_PAGE_VERSION = 4
EMAIL_MATCHING_PAGE_VERSION = 1
MANUAL_CAMPAIGN_TOTALS_VERSION = 1
GROUPED_SUMMARY_OUTPUT_VERSION = 1
VISIBLE_WORKSPACE_SECTIONS_VERSION = 1
SMS_MATCHING_PAGE_VERSION = 1
AGENT_FIRST_UPLOADS_VERSION = 1
CHANNELS = ("SMS", "EMAIL", "VIBER")


def render_campaign_totals(report_date: date, campaign: str, viber_report=None,
                           email_report=None, *, include_worked: bool = False, sms_report=None, source_campaigns=()) -> dict:
    """Keep manual totals with their campaign/date when uploads or visible rows change."""
    saved = {key: dict(value) for key, value in st.session_state.get("inbound_reply_campaign_totals", {}).items()}
    groups = campaign_groups(viber_report, email_report, sms_report, source_campaigns=source_campaigns) or [{"label": campaign.strip() or "—"}]
    with st.container(border=True, key="inbound_campaign_totals"):
        st.subheader("Edit campaign totals")
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
                      campaign_totals: dict | None = None, sms_report=None, *, source_campaigns=()) -> str:
    """Fill each uploaded channel independently, joining campaign labels exactly."""
    date_label = f"{report_date.strftime('%B').upper()} {report_date.day}, {report_date.year} DATA"
    channel_headers = "".join(
        f'<th scope="col">{channel}</th><th scope="col">%</th>'
        for channel in ("EMAIL", "SMS", "VIBER")
    )
    body = []
    for cells in summary_rows(campaign, report_date, sent_to_ops, viber_report, email_report, campaign_totals, sms_report,
                              source_campaigns=source_campaigns):
        body.append("<tr>" + "".join(f"<td>{escape(display_cell(value, column))}</td>"
                                    for column, value in enumerate(cells)) + "</tr>")
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
      .inbound-campaign-report table {{ min-width: 1120px; table-layout: fixed; font-size: clamp(8px, 0.6vw, 10px); }}
      .inbound-campaign-report th, .inbound-campaign-report td {{ padding: 4px; line-height: 1.25; }}
    </style>
    <div class="inbound-reply-report inbound-campaign-report" role="region" aria-label="Inbound reply summary" tabindex="0">
      <table aria-label="Inbound reply campaign report">
        <colgroup><col style="width:14.634%">{''.join('<col style="width:4.878%">' for _ in range(10))}
        <col style="width:7.317%">{''.join('<col style="width:4.878%">' for _ in range(6))}</colgroup>
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


def prepare_viber_report(agent_upload, customer_upload) -> ViberReplyReport | None:
    """Keep prepared files session-local and invalidate on upload content changes."""
    if agent_upload is None or customer_upload is None:
        st.session_state.pop("inbound_reply_viber_result", None)
        return None
    agent_bytes, customer_bytes = agent_upload.getvalue(), customer_upload.getvalue()
    signature = (AGENT_FIRST_UPLOADS_VERSION, VIBER_REPLY_VERSION, VIBER_AGENT_REPLIED_SOURCE, VIBER_MATCH_MODE,
                 sha256(agent_bytes).hexdigest(), sha256(customer_bytes).hexdigest(),
                 agent_upload.name, customer_upload.name)
    cached = st.session_state.get("inbound_reply_viber_result")
    if cached is None or cached["signature"] != signature:
        st.session_state.pop("inbound_reply_viber_result", None)
        try:
            if agent_bytes == customer_bytes:
                raise InboundReplyError("The two Viber uploads are identical. Upload agent replies and customer replies separately.")
            customers = read_reply_source(customer_upload.name, customer_bytes)
            agents = read_reply_source(agent_upload.name, agent_bytes)
            # The user's corrected file roles determine status; Message Type is source data.
            report = match_viber_sources(customers, agents,
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


def prepare_sms_report(replied_upload, account_upload) -> ViberReplyReport | None:
    """SMS follows the phone/campaign matching policy used by Viber."""
    if account_upload is None or replied_upload is None:
        st.session_state.pop("inbound_reply_sms_result", None)
        return None
    account_bytes, replied_bytes = account_upload.getvalue(), replied_upload.getvalue()
    signature = (AGENT_FIRST_UPLOADS_VERSION, SMS_MATCHING_PAGE_VERSION, VIBER_REPLY_VERSION,
                 sha256(account_bytes).hexdigest(), sha256(replied_bytes).hexdigest(),
                 account_upload.name, replied_upload.name)
    cached = st.session_state.get("inbound_reply_sms_result")
    if cached is None or cached["signature"] != signature:
        st.session_state.pop("inbound_reply_sms_result", None)
        try:
            accounts = read_reply_source(account_upload.name, account_bytes)
            replied = read_reply_source(replied_upload.name, replied_bytes)
            cached = {"signature": signature, "report": match_viber_sources(
                accounts, replied, agent_replied_source="campaign_replies", match_mode="phone",
            )}
        except InboundReplyError as error:
            cached = {"signature": signature, "error": str(error)}
        st.session_state["inbound_reply_sms_result"] = cached
    if "error" in cached:
        st.error(cached["error"])
        return None
    return cached["report"]


def prepare_campaign_labels(channels) -> tuple[str, ...]:
    """Show source campaign names while a channel is still waiting for its other file."""
    sources = []
    for channel, uploads, report in channels:
        if report is not None:
            continue
        for index, upload in enumerate(uploads):
            # Email History has no campaign field; DC Checking supplies its labels.
            if upload is not None and (channel != "EMAIL" or index == 1):
                data = upload.getvalue()
                sources.append((channel, index, upload.name, data))
    if not sources:
        st.session_state.pop("inbound_reply_campaign_labels", None)
        return ()
    signature = (VIBER_REPLY_VERSION, EMAIL_REPLY_VERSION,
                 tuple((channel, index, name, sha256(data).hexdigest()) for channel, index, name, data in sources))
    cached = st.session_state.get("inbound_reply_campaign_labels")
    if cached is None or cached["signature"] != signature:
        labels = []
        for channel, _, name, data in sources:
            try:
                source = read_email_source(name, data, kind="dc") if channel == "EMAIL" else read_reply_source(name, data)
            except InboundReplyError:
                continue  # Full-pair preparation displays input errors.
            labels.extend(record.campaign for record in source.records)
        cached = {"signature": signature, "labels": tuple(dict.fromkeys(labels))}
        st.session_state["inbound_reply_campaign_labels"] = cached
    return cached["labels"]


def channel_source_status(uploads, report) -> str:
    if report is not None:
        return "Processed"
    count = sum(upload is not None for upload in uploads)
    if count == 0:
        return "Not uploaded"
    if count == 1:
        return "Incomplete: upload both source files"
    return "Could not process: check the upload errors on the page"


def render_sms_details(report: ViberReplyReport) -> None:
    totals = report.totals
    for column, label, value in zip(st.columns(3),
                                    ("SMS unique phone numbers", "SMS replied by agents", "SMS awaiting agent reply"),
                                    (totals.unique_people, totals.agent_replied, totals.awaiting_reply)):
        column.metric(label, f"{value:,}")
    st.caption(
        "SMS matches normalized phone numbers within each campaign. Upload agent replies first "
        "and customer replies second. Each file needs Phone and Campaign or Bank Campaign columns. "
        "SMS Unreplied in the main output workbook contains the full account details."
    )
    excluded = report.message_logs.excluded_rows + report.campaign_replies.excluded_rows + report.excluded_phone_rows
    if excluded:
        st.warning(f"Skipped {excluded:,} SMS rows without a usable phone number.")


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
    with st.container():
        st.markdown("**Email matching details**")
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
        "Viber Reply counts unique phone numbers across both files. Agent Worked counts phone numbers in the agent-replies file, "
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
    with st.container():
        st.markdown("**Viber matching details**")
        st.write(
            f"Agent replies: {len(report.campaign_replies.records):,} rows. "
            f"Customer replies: {len(report.message_logs.records):,} rows. "
            f"Combined into {totals.unique_people:,} unique phone numbers."
        )
        st.caption(
            "Matching uses phone numbers, normalizing spaces, punctuation and Philippine +63/0 formats. "
            "Names are shown for reference. Rows without a usable phone number are excluded and counted. "
            "An agent reply takes priority over pending entries for the same phone number and campaign. "
            "Numbers found only in the customer-replies file await an agent reply. "
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
    output_area = st.container(key="inbound_summary_output")

    with st.container(border=True, key="inbound_report_settings"):
        st.subheader("Report settings")
        campaign_column, date_column, ops_column = st.columns([2, 1, 1])
        with campaign_column:
            campaign = st.text_input("Report label (optional)", placeholder="Report name", key="inbound_reply_campaign",
                                     help="Campaign names are read automatically from the uploaded Campaign, Campaigns or Bank Campaign column.")
        with date_column:
            report_date = st.date_input(
                "Report date", value=datetime.now(timezone(timedelta(hours=8))).date(),
                key="inbound_reply_date",
            )
        with ops_column:
            sent_to_ops = st.selectbox(
                "Sent to OPS", ("—", "YES", "NO"), key="inbound_reply_sent_to_ops",
            )

    viber_uploads = []
    email_uploads = []
    sms_uploads = []
    with st.container(border=True, key="inbound_upload_files"):
        st.subheader("Upload files")
        st.caption("Upload two files for each channel you want to process: file 1 is agent replies; file 2 is customer replies.")
        uploaded_count = 0
        for column, channel in zip(st.columns(3), CHANNELS):
            with column:
                st.markdown(f"**{channel}**")
                channel_count = 0
                for number in (1, 2):
                    role = "Agent replies" if number == 1 else "Customer replies"
                    label = f"{channel} file {number} — {role}"
                    help_text = None
                    if channel == "EMAIL":
                        help_text = "Email History with sent agent replies." if number == 1 else "DC Checking customer account list."
                    elif channel == "VIBER":
                        help_text = "Campaign replies export containing agent-replied accounts." if number == 1 else "Message logs containing customer replies to compare."
                    upload = st.file_uploader(
                        label, type=["xlsx", "csv"], help=help_text,
                        key=upload_widget_key(f"inbound_reply_{channel.lower()}_{number}_upload"),
                    )
                    channel_count += upload is not None
                    if channel == "VIBER":
                        viber_uploads.append(upload)
                    elif channel == "EMAIL":
                        email_uploads.append(upload)
                    elif channel == "SMS":
                        sms_uploads.append(upload)
                uploaded_count += channel_count
                st.caption(f"{channel_count} of 2 files uploaded")
        st.caption(f"{uploaded_count} of 6 files uploaded")

    viber_report = prepare_viber_report(*viber_uploads)
    email_report = prepare_email_report(*email_uploads)
    sms_report = prepare_sms_report(*sms_uploads)
    channels = (("SMS", sms_uploads, sms_report), ("VIBER", viber_uploads, viber_report), ("EMAIL", email_uploads, email_report))
    source_campaigns = prepare_campaign_labels(channels)
    manual_totals = render_campaign_totals(report_date, campaign, viber_report, email_report,
                                           include_worked=True, sms_report=sms_report, source_campaigns=source_campaigns)
    statuses = {channel: channel_source_status(uploads, report) for channel, uploads, report in channels}
    with output_area:
        st.html(build_report_html(campaign, report_date, sent_to_ops, viber_report, email_report, manual_totals, sms_report,
                                  source_campaigns=source_campaigns))
        st.caption("CAMPAIGN is filled automatically from the uploaded campaign column, with one row per campaign.")
        has_manual_counts = any(value is not None for row in summary_rows(
            campaign, report_date, sent_to_ops, viber_report, email_report, manual_totals, sms_report,
            source_campaigns=source_campaigns)
            for value in row[1:11])
        if viber_report is not None or email_report is not None or sms_report is not None or has_manual_counts:
            st.download_button(
                "Download report (.xlsx)",
                data=partial(build_summary_workbook, campaign, report_date, sent_to_ops, viber_report, email_report,
                             manual_totals, sms_report=sms_report, channel_statuses=statuses, source_campaigns=source_campaigns),
                file_name=f"INBOUND_REPLY_{report_date.isoformat()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="inbound_reply_summary_download", on_click="ignore", type="primary",
            )
            st.caption("Includes Summary, SMS Unreplied, Viber Unreplied and Email Unreplied worksheets.")
    if sms_report is not None:
        with st.container(border=True, key="inbound_sms_details"):
            st.subheader("SMS account details")
            render_sms_details(sms_report)
    if email_report is not None:
        with st.container(border=True, key="inbound_email_details"):
            st.subheader("Email account details")
            render_email_details(email_report, report_date)
    if viber_report is not None:
        with st.container(border=True, key="inbound_viber_details"):
            st.subheader("Viber account details")
            render_viber_details(viber_report, report_date)
