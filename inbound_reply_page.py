"""Six-file upload workspace and the inbound reply report layout."""

from __future__ import annotations

from datetime import date, datetime
from html import escape
from hashlib import sha256
from functools import partial

import streamlit as st

from inbound_reply import (
    VIBER_REPLY_VERSION, VIBER_AGENT_REPLIED_SOURCE, VIBER_MATCH_MODE, INBOUND_CSV_DELIMITER_VERSION, SMS_CLIENT_TOUCH_VERSION,
    MANILA_TIMEZONE, SBLEGAL_CAMPAIGN, InboundReplyError, ReplySource, ViberReplyReport,
    build_people_csv, build_unreplied_workbook, deduplicate_client_accounts,
    match_client_touches, match_viber_sources, normalized_phone, people_rows, read_reply_source,
)
from inbound_email import (
    EMAIL_REPLY_VERSION, EmailReplyReport, EmailSource, build_email_unreplied_workbook,
    match_email_sources, read_email_source,
)
from inbound_campaign_lookup import (
    EMAIL_CAMPAIGN_LOOKUP_VERSION, PHONE_CAMPAIGN_LOOKUP_VERSION, SAVED_CAMPAIGN_LOOKUP_PATH, EmailCampaignLookup,
    apply_email_campaign_lookup, apply_phone_campaign_lookup, build_email_campaign_lookup,
    read_saved_campaign_lookup_files, save_campaign_lookup_files,
)
from upload_state import INBOUND_LOOKUP_FALLBACK_CACHE_KEY, INBOUND_TOTALS_UPLOAD_KEY, upload_widget_key
from inbound_summary import (
    BLAST_PERCENTAGE_FORMULAS_VERSION, build_client_accounts_workbook, build_summary_workbook, build_unmatched_workbook,
    campaign_totals_key, display_cell, merge_mindanao_totals, overall_blast_row, read_summary_totals, summary_rows,
)


INBOUND_REPLY_PAGE_VERSION = 31
INBOUND_REPLY_TITLE = "INBOUND REPLY"
VIBER_MATCHING_PAGE_VERSION = 4
EMAIL_MATCHING_PAGE_VERSION = 1
GROUPED_SUMMARY_OUTPUT_VERSION = 1
VISIBLE_WORKSPACE_SECTIONS_VERSION = 1
SMS_MATCHING_PAGE_VERSION = 1
AGENT_FIRST_UPLOADS_VERSION = 1
READABLE_SUMMARY_HEADERS_VERSION = 1
REPORT_SETTINGS_REMOVED_VERSION = 1
PL_RECOVERY_LOOKUP_VERSION = 2
CHANNELS = ("SMS", "EMAIL", "VIBER")


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
    rows = summary_rows(campaign, report_date, sent_to_ops, viber_report, email_report, campaign_totals, sms_report,
                        source_campaigns=source_campaigns)
    for cells in rows:
        body.append("<tr>" + "".join(f"<td>{escape(display_cell(value, column))}</td>"
                                    for column, value in enumerate(cells)) + "</tr>")
    footer = "".join(f"<td>{escape(display_cell(value, column))}</td>"
                     for column, value in enumerate(overall_blast_row(rows)))
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
      .inbound-campaign-report {{ max-width: 100%; box-sizing: border-box; }}
      .inbound-campaign-report table {{ min-width: 1080px; table-layout: auto; font-size: 11px; }}
      .inbound-campaign-report th, .inbound-campaign-report td {{ padding: 6px 4px; line-height: 1.35; }}
      .inbound-campaign-report th {{ white-space: normal; overflow-wrap: normal; word-break: normal; }}
      .inbound-campaign-report tfoot td {{ background: #e8f2ec; font-weight: 700; border-top-width: 2px; }}
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
            <th scope="col" title="Unique customer accounts that replied by email, regardless of receiving mailbox or assigned agent. Counted once per customer email address within each campaign.">EMAIL REPLY</th>
            <th scope="col" title="Unique customer accounts that replied by SMS, regardless of receiving inbox or assigned agent. Repeated replies count once per phone number within each campaign.">SMS REPLY</th>
            <th scope="col" title="Unique customer accounts that replied by Viber, regardless of receiving inbox or assigned agent. Repeated replies count once per phone number within each campaign.">VIBER REPLY</th>
            {channel_headers}
          </tr>
        </thead>
        <tbody>{''.join(body)}</tbody>
        <tfoot aria-label="Campaign totals"><tr>{footer}</tr></tfoot>
      </table>
    </div>
    """


def prepare_viber_report(agent_upload, customer_upload, lookup: EmailCampaignLookup | None = None) -> ViberReplyReport | None:
    """Keep prepared files session-local and invalidate on upload content changes."""
    if agent_upload is None or customer_upload is None:
        st.session_state.pop("inbound_reply_viber_result", None)
        return None
    agent_bytes, customer_bytes = agent_upload.getvalue(), customer_upload.getvalue()
    signature = (PL_RECOVERY_LOOKUP_VERSION, AGENT_FIRST_UPLOADS_VERSION, VIBER_REPLY_VERSION, INBOUND_CSV_DELIMITER_VERSION,
                 VIBER_AGENT_REPLIED_SOURCE, VIBER_MATCH_MODE, PHONE_CAMPAIGN_LOOKUP_VERSION,
                 sha256(agent_bytes).hexdigest(), sha256(customer_bytes).hexdigest(),
                 agent_upload.name, customer_upload.name, lookup.signature if lookup is not None else None)
    cached = st.session_state.get("inbound_reply_viber_result")
    if cached is None or cached["signature"] != signature:
        st.session_state.pop("inbound_reply_viber_result", None)
        try:
            customers = read_reply_source(customer_upload.name, customer_bytes, require_campaign=False)
            agents = read_reply_source(agent_upload.name, agent_bytes, require_campaign=False)
            if lookup is not None:
                customers = apply_phone_campaign_lookup(customers, lookup)
                agents = apply_phone_campaign_lookup(agents, lookup)
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


def prepare_email_lookup(signature, read_files, *, cache_key: str) -> tuple[EmailCampaignLookup | None, bool]:
    signature = (EMAIL_CAMPAIGN_LOOKUP_VERSION, signature)
    cached = st.session_state.get(cache_key)
    if cached is None or cached["signature"] != signature:
        try:
            cached = {"signature": signature, "lookup": build_email_campaign_lookup(read_files())}
        except InboundReplyError as error:
            cached = {"signature": signature, "error": str(error), "lookup": (cached or {}).get("lookup")}
        st.session_state[cache_key] = cached
    if "error" in cached:
        st.error(cached["error"])
        if cached.get("lookup") is not None:
            st.caption("Using the previous valid campaign lookup.")
        return cached.get("lookup"), False
    return cached["lookup"], True


def render_campaign_lookup(uploads) -> EmailCampaignLookup | None:
    """Check saved databases before using current uploads to fill missing matches."""
    path = SAVED_CAMPAIGN_LOOKUP_PATH
    saved_file = None
    saved_lookup = None
    saved_key = "inbound_reply_email_lookup"
    try:
        saved_file = path.stat()
    except FileNotFoundError:
        st.session_state.pop(saved_key, None)
    except OSError:
        st.error("Could not access the saved campaign databases. Check the app's uploads folder.")
        saved_lookup = st.session_state.get(saved_key, {}).get("lookup")

    if saved_file is not None:
        if st.button("Remove saved lookup files", key="inbound_reply_remove_saved_lookup"):
            try:
                path.unlink(missing_ok=True)
            except OSError:
                st.error("Could not remove the saved campaign databases.")
            else:
                st.rerun()
        signature = (path, saved_file.st_ino, saved_file.st_mtime_ns, saved_file.st_size)
        saved_lookup, _ = prepare_email_lookup(signature, partial(read_saved_campaign_lookup_files, path), cache_key=saved_key)
        if saved_lookup is not None:
            st.caption("Using saved campaign lookup files.")

    if not uploads:
        st.session_state.pop(INBOUND_LOOKUP_FALLBACK_CACHE_KEY, None)
        return saved_lookup

    files = [(upload.name, upload.getvalue()) for upload in uploads]
    signature = ("uploads", tuple((name, sha256(data).hexdigest()) for name, data in files))
    uploaded_lookup, upload_valid = prepare_email_lookup(signature, lambda: files, cache_key=INBOUND_LOOKUP_FALLBACK_CACHE_KEY)
    if upload_valid:
        label = "Replace saved lookup files" if saved_file is not None else "Save lookup files"
        if st.button(label, key="inbound_reply_save_lookup",
                     help="Use the full uploaded set as the primary lookup for future sessions."):
            try:
                save_campaign_lookup_files(path, files)
            except InboundReplyError as error:
                st.error(str(error))
            else:
                st.rerun()

    if saved_lookup is None:
        return uploaded_lookup
    if uploaded_lookup is None:
        return saved_lookup
    st.caption("Uploaded lookup files fill contacts without a match in the saved databases.")
    return saved_lookup.with_fallback(uploaded_lookup)


def prepare_email_report(history_upload, dc_upload, lookup: EmailCampaignLookup | None = None) -> EmailReplyReport | None:
    if dc_upload is None:
        st.session_state.pop("inbound_reply_email_result", None)
        return None
    dc_bytes = dc_upload.getvalue()
    history_bytes = history_upload.getvalue() if history_upload is not None else None
    history_signature = (history_upload.name, sha256(history_bytes).hexdigest()) if history_bytes is not None else None
    signature = (PL_RECOVERY_LOOKUP_VERSION, EMAIL_REPLY_VERSION, history_signature, sha256(dc_bytes).hexdigest(),
                 dc_upload.name, lookup.signature if lookup is not None else None)
    cached = st.session_state.get("inbound_reply_email_result")
    if cached is None or cached["signature"] != signature:
        cached = {"signature": signature}
        try:
            dc = read_email_source(dc_upload.name, dc_bytes, kind="dc", require_campaign=False)
            dc = apply_email_campaign_lookup(dc, lookup)
        except InboundReplyError as error:
            cached["error"] = str(error)
        else:
            history = None
            if history_bytes is not None:
                try:
                    if history_bytes == dc_bytes:
                        raise InboundReplyError("The two Email uploads are identical. Upload Email History and DC Checking separately.")
                    history = read_email_source(history_upload.name, history_bytes, kind="history")
                except InboundReplyError as error:
                    cached["error"] = str(error)
            cached["report"] = match_email_sources(dc, history)
        st.session_state["inbound_reply_email_result"] = cached
    if "error" in cached:
        st.error(cached["error"])
    return cached.get("report")


def _merge_sms_sources(sources: list[ReplySource]) -> ReplySource:
    if len(sources) == 1:
        return sources[0]
    return ReplySource(
        ", ".join(source.filename for source in sources),
        tuple(row for source in sources for row in source.records),
        sum(source.excluded_rows for source in sources),
        has_reply_dates=any(source.has_reply_dates for source in sources),
    )


def prepare_sms_report(replied_upload, account_upload, lookup: EmailCampaignLookup | None = None) -> ViberReplyReport | None:
    """Build the SMS reply report or the Client and Agent Touch report."""
    if account_upload is None or not replied_upload:
        st.session_state.pop("inbound_reply_sms_result", None)
        return None
    account_bytes = account_upload.getvalue()
    agent_uploads = replied_upload if isinstance(replied_upload, list) else [replied_upload]
    agent_files = [(upload.name, upload.getvalue()) for upload in agent_uploads]
    signature = (PL_RECOVERY_LOOKUP_VERSION, AGENT_FIRST_UPLOADS_VERSION, SMS_MATCHING_PAGE_VERSION,
                 VIBER_REPLY_VERSION, INBOUND_CSV_DELIMITER_VERSION, SMS_CLIENT_TOUCH_VERSION, PHONE_CAMPAIGN_LOOKUP_VERSION,
                 sha256(account_bytes).hexdigest(), account_upload.name,
                 tuple((name, sha256(content).hexdigest()) for name, content in agent_files),
                 lookup.signature if lookup is not None else None)
    cached = st.session_state.get("inbound_reply_sms_result")
    if cached is None or cached["signature"] != signature:
        st.session_state.pop("inbound_reply_sms_result", None)
        try:
            accounts = read_reply_source(account_upload.name, account_bytes, require_campaign=False)
            agent_sources = [read_reply_source(name, content, require_campaign=False) for name, content in agent_files]
            if accounts.layout == "agent_touch" or any(source.layout == "client" for source in agent_sources):
                raise InboundReplyError("Put Agent Touch workbooks in SMS file 1 and the Client workbook in SMS file 2.")
            client_mode = accounts.layout == "client" or any(source.layout == "agent_touch" for source in agent_sources)
            if client_mode and accounts.layout != "client":
                accounts = deduplicate_client_accounts(accounts)
            replied = _merge_sms_sources([source for source in agent_sources if source.layout != "agent_touch"])
            if lookup is not None:
                accounts = apply_phone_campaign_lookup(accounts, lookup)
            if client_mode:
                touches = _merge_sms_sources([source for source in agent_sources if source.layout == "agent_touch"])
                report = match_client_touches(accounts, touches, replied)
            else:
                if lookup is not None:
                    replied = apply_phone_campaign_lookup(replied, lookup)
                report = match_viber_sources(accounts, replied, agent_replied_source="campaign_replies", match_mode="phone")
            cached = {"signature": signature, "report": report}
        except InboundReplyError as error:
            cached = {"signature": signature, "error": str(error)}
        st.session_state["inbound_reply_sms_result"] = cached
    if "error" in cached:
        st.error(cached["error"])
        return None
    return cached["report"]


def prepare_campaign_inputs(channels, email_lookup: EmailCampaignLookup | None = None) -> tuple[
        tuple[str, ...], dict[str, EmailSource | ReplySource]]:
    """Collect campaign labels and unmatched customer rows, including from pending uploads."""
    sources, unmatched = [], {}
    for channel, (_, upload), report in channels:
        if report is not None:
            source = report.dc if channel == "EMAIL" else report.message_logs
            if source.unmatched_records:
                unmatched[channel] = source
        elif upload is not None:
            sources.append((channel, upload.name, upload.getvalue()))
    if not sources:
        st.session_state.pop("inbound_reply_campaign_labels", None)
        return (), unmatched
    signature = (PL_RECOVERY_LOOKUP_VERSION, VIBER_REPLY_VERSION, EMAIL_REPLY_VERSION, EMAIL_CAMPAIGN_LOOKUP_VERSION,
                 SMS_CLIENT_TOUCH_VERSION, PHONE_CAMPAIGN_LOOKUP_VERSION,
                 email_lookup.signature if email_lookup is not None else None,
                 tuple((channel, name, sha256(data).hexdigest()) for channel, name, data in sources))
    cached = st.session_state.get("inbound_reply_campaign_labels")
    if cached is None or cached["signature"] != signature:
        labels = []
        pending_unmatched = {}
        for channel, name, data in sources:
            try:
                if channel == "EMAIL":
                    source = read_email_source(name, data, kind="dc", require_campaign=False)
                    source = apply_email_campaign_lookup(source, email_lookup)
                    labels.extend(record.campaign for record in source.records)
                else:
                    source = read_reply_source(name, data, require_campaign=False)
                    if channel == "SMS" and source.layout == "agent_touch":
                        continue
                    if email_lookup is not None:
                        source = apply_phone_campaign_lookup(source, email_lookup)
                    labels.extend(record.campaign for record in source.records
                                  if source.layout == "client" or normalized_phone(record.phone))
            except InboundReplyError:
                continue  # Full-pair preparation displays input errors.
            if source.unmatched_records:
                pending_unmatched[channel] = source
        cached = {"signature": signature, "labels": tuple(dict.fromkeys(labels)), "unmatched": pending_unmatched}
        st.session_state["inbound_reply_campaign_labels"] = cached
    return cached["labels"], unmatched | cached["unmatched"]


def channel_source_status(uploads, report) -> str:
    if report is not None:
        return "Processed"
    count = sum(bool(upload) for upload in uploads)
    if count == 0:
        return "Not uploaded"
    if count == 1:
        return "Incomplete: upload both source files"
    return "Could not process: check the upload errors on the page"


def render_sms_details(report: ViberReplyReport, report_date: date) -> None:
    totals = report.totals
    if report.client_touches is not None:
        unique_touches = len(deduplicate_client_accounts(report.campaign_replies).records)
        clients_touched = sum(touched for _, touched in report.client_touches)
        for column, label, count in zip(st.columns(5),
                                       ("Unique client accounts", "Replied by agents", "Awaiting agent reply",
                                        "Clients touched", "Accounts in Agent Touch files"),
                                       (totals.unique_people, totals.agent_replied, totals.awaiting_reply,
                                        clients_touched, unique_touches)):
            column.metric(label, f"{count:,}")
        st.caption(f"Removed {report.message_logs.duplicate_rows:,} duplicate Client rows. "
                   "Agent Touch files record calls. Upload SMS agent replies to confirm replies to these clients.")
        if report.message_logs.lookup_applied:
            st.caption("The report keeps Client campaign labels and checks placements against database phone numbers.")
        else:
            st.info("Upload or save campaign lookup files to verify placements by phone. The current labels come from the Client file.")
        if report.message_logs.unmatched_addresses:
            st.caption(f"{report.message_logs.unmatched_addresses:,} client accounts have an unconfirmed placement. "
                       "See the unmatched accounts download for details.")
        st.download_button(
            "Download unique client accounts (.xlsx)", data=partial(build_client_accounts_workbook, report),
            file_name=f"INBOUND_CLIENT_ACCOUNTS_{report_date.isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="inbound_reply_client_accounts_download", on_click="ignore",
        )
        return
    for column, label, value in zip(st.columns(3),
                                    ("Customer SMS replies", "SMS replied by agents", "SMS awaiting agent reply"),
                                    (totals.unique_people, totals.agent_replied, totals.awaiting_reply)):
        column.metric(label, f"{value:,}")
    st.caption(
        "SMS matches normalized phone numbers within each campaign. SMS Unreplied in the main output workbook "
        "contains the full account details."
    )
    excluded = report.message_logs.excluded_rows + report.campaign_replies.excluded_rows + report.excluded_phone_rows
    if excluded:
        st.warning(f"Skipped {excluded:,} SMS rows without a usable phone number.")
    if report.message_logs.unmatched_addresses:
        st.caption(f"{report.message_logs.unmatched_addresses:,} SMS customer phone numbers without a campaign match "
                   f"are grouped under {SBLEGAL_CAMPAIGN}.")


def render_email_details(report: EmailReplyReport, report_date: date) -> None:
    totals = report.totals
    for column, label, value in zip(st.columns(3),
                                    ("Customer email replies", "Email replied by agents", "Email awaiting agent reply"),
                                    (totals.unique_people, totals.agent_replied, totals.awaiting_reply)):
        column.metric(label, display_cell(value, 8))
    if report.dc.workbook_placements:
        placement_source = "Campaign placements come from the lookup results included in DC Checking. "
    elif report.dc.lookup_applied:
        placement_source = "Campaigns come from the lookup files and any PL RECOVERY labels in DC Checking. "
    else:
        placement_source = "DC Checking supplies incoming customer emails and campaigns. "
    st.caption(
        placement_source +
        "UNIQUE RESULT — EMAIL REPLY counts customer accounts that replied by email, regardless of receiving mailbox "
        "or assigned agent. Repeated replies count once per customer email address within each campaign. "
        "A matching Email History row with a nonblank sent value confirms the agent's reply."
    )
    if report.dc.unmatched_addresses:
        st.caption(f"{report.dc.unmatched_addresses:,} email addresses without a campaign match "
                   f"are grouped under {SBLEGAL_CAMPAIGN}.")
    if report.history is None:
        st.info("Upload a valid Email History file in EMAIL file 1 to see Agent Worked and unreplied accounts.")
        if report.dc.excluded_rows:
            st.warning(f"Skipped {report.dc.excluded_rows:,} rows with blank or invalid email addresses.")
        return
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
    st.caption("The Excel workbook contains Summary and Unreplied Details sheets, retaining the original DC Checking fields."
               + (" Lookup Campaign contains the resolved campaign." if report.dc.lookup_applied else ""))
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
            "Sent replies match the included DC campaigns containing that address on the same day."
        )
        excluded = report.dc.excluded_rows + report.history.excluded_rows
        if excluded:
            st.warning(f"Skipped {excluded:,} rows with blank or invalid email addresses.")


def render_viber_details(report: ViberReplyReport, report_date: date | None = None) -> None:
    totals = report.totals
    for column, label, value in zip(st.columns(4),
                                    ("Customer Viber replies", "Replied by agents", "Awaiting agent reply", "Found in both files"),
                                    (totals.unique_people, totals.agent_replied, totals.awaiting_reply, totals.matched_both_files)):
        column.metric(label, f"{value:,}")
    st.caption(
        "Viber matches normalized phone numbers within each campaign. The totals above count each phone number "
        "once across all campaigns."
    )
    if report.message_logs.unmatched_addresses:
        st.caption(f"{report.message_logs.unmatched_addresses:,} Viber customer phone numbers without a campaign match "
                   f"are grouped under {SBLEGAL_CAMPAIGN}.")
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
            f"{totals.unique_people:,} unique customer phone numbers."
        )
        st.caption(
            "Matching uses phone numbers, normalizing spaces, punctuation and Philippine +63/0 formats. "
            "Names are shown for reference. Rows without a usable phone number are excluded and counted. "
            "An agent reply takes priority over pending entries for the same phone number and campaign. "
            "Numbers found only in the customer-replies file await an agent reply. "
            "Numbers found only in the agent-replies file are outside the customer-reply population."
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


def inbound_report_date() -> date:
    return datetime.now(MANILA_TIMEZONE).date()


def render_inbound_reply() -> None:
    st.title("INBOUND REPLY")
    output_area = st.container(key="inbound_summary_output")

    campaign = ""
    report_date = inbound_report_date()
    sent_to_ops = "—"
    imported_totals = {}

    viber_uploads = []
    email_uploads = []
    sms_uploads = []
    with st.container(border=True, key="inbound_upload_files"):
        st.subheader("Upload files")
        st.caption("Use file 1 for agent replies or SMS Agent Touch and file 2 for customer replies.")
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
                        help_text = ("Email History with sent agent replies." if number == 1 else
                                     "DC Checking supplies Email Reply counts from all receiving mailboxes, matched by sender.")
                    elif channel == "VIBER":
                        help_text = ("Message logs export containing Agent Reply rows." if number == 1 else
                                     "Campaign replies export containing Inbound Reply rows from customers.")
                    elif channel == "SMS":
                        help_text = ("Upload SMS agent replies or one or more Agent Touch workbooks." if number == 1 else
                                     "Upload SMS customer replies or the Client workbook with incoming messages.")
                    upload = st.file_uploader(
                        label, type=["xlsx", "csv"], help=help_text,
                        accept_multiple_files=channel == "SMS" and number == 1,
                        key=upload_widget_key(f"inbound_reply_{channel.lower()}_{number}_upload"),
                    )
                    channel_count += bool(upload)
                    if channel == "VIBER":
                        viber_uploads.append(upload)
                    elif channel == "EMAIL":
                        email_uploads.append(upload)
                    elif channel == "SMS":
                        sms_uploads.append(upload)
                uploaded_count += channel_count
                st.caption(f"{channel_count} of 2 upload slots filled")
        st.caption(f"{uploaded_count} of 6 upload slots filled")

        summary_upload = st.file_uploader(
            "Campaign totals workbook", type=["xlsx"], key=upload_widget_key(INBOUND_TOTALS_UPLOAD_KEY),
            help="Use an Inbound Reply workbook to carry forward its report date, Active Ticket, and DC Worked counts.",
        )
        if summary_upload is not None:
            try:
                report_date, imported_totals = read_summary_totals(summary_upload.getvalue())
            except InboundReplyError as error:
                st.error(str(error))

        st.markdown("**Campaign lookup — Email, SMS and Viber**")
        lookup_uploads = st.file_uploader(
            "Campaign lookup files", type=["xlsx", "csv"], accept_multiple_files=True,
            key=upload_widget_key("inbound_reply_email_lookup_uploads"),
            help="Saved databases are checked first. Upload newer Regular, GEO, B2 or B4 files to fill missing matches. "
                 "Keep B2 or B4 in those filenames. "
                 "Regular uses LEVEL or placement; GEO uses Area. B4 email matching uses only Email (01). "
                 "Email uses email addresses; SMS and Viber use phone numbers. The OVERALL worksheet is used when available.",
        )
        st.caption("Regular: Levels 1–5 → SBC Level 1; Level 6 → SBC Level 6. "
                   "GEO: Region 3 and Mindanao are grouped under SBC Mindanao. "
                   "B2 → SBC B2; B4 → SBC B4 using only Email (01) for email matching. "
                   "B2 takes priority for contacts in both bucket files. "
                   "Area takes priority over product PL within each lookup set. Product PL without an Area match uses PL RECOVERY. "
                   "SMS and Viber keep the campaigns in the reply files. The lookup fills missing campaigns.")
        email_lookup = render_campaign_lookup(lookup_uploads)
        if email_lookup is not None:
            sources = email_lookup.sources
            if email_lookup.fallback is not None:
                sources += email_lookup.fallback.sources
            for source in sources:
                label = f" ({source.campaign_label})" if source.campaign_label else ""
                st.caption(f"{source.filename}{label}: {len(source.campaigns_by_email):,} email addresses and "
                           f"{len(source.campaigns_by_phone):,} phone numbers with campaigns "
                           f"from {', '.join(source.sheets)}.")
                if source.missing_campaign_rows:
                    st.warning(f"{source.filename}: {source.missing_campaign_rows:,} rows have contact details but no usable campaign.")

    viber_report = prepare_viber_report(*viber_uploads, email_lookup)
    email_report = prepare_email_report(*email_uploads, email_lookup)
    sms_report = prepare_sms_report(*sms_uploads, email_lookup)
    channels = (("SMS", sms_uploads, sms_report), ("VIBER", viber_uploads, viber_report), ("EMAIL", email_uploads, email_report))
    source_campaigns, unmatched_sources = prepare_campaign_inputs(channels, email_lookup)
    campaign_totals = merge_mindanao_totals(report_date, st.session_state.get("inbound_reply_campaign_totals"))
    active_ticket_counts = email_lookup.active_ticket_counts if email_lookup is not None else {}
    has_phone_customers = sms_uploads[1] is not None or viber_uploads[1] is not None
    if email_lookup is not None and not imported_totals and (email_uploads[1] is not None or not has_phone_customers):
        source_campaigns = (*source_campaigns, *email_lookup.campaigns, *active_ticket_counts)
    if active_ticket_counts:
        for name, count in active_ticket_counts.items():
            key = campaign_totals_key(report_date, name)
            campaign_totals[key] = {**campaign_totals.get(key, {}), "active_ticket": count}
    for name, counts in imported_totals.items():
        key = campaign_totals_key(report_date, name)
        campaign_totals[key] = {**campaign_totals.get(key, {}), **counts}
    source_campaigns = (*source_campaigns, *imported_totals)
    statuses = {channel: channel_source_status(uploads, report) for channel, uploads, report in channels}
    with output_area:
        st.html(build_report_html(campaign, report_date, sent_to_ops, viber_report, email_report, campaign_totals, sms_report,
                                  source_campaigns=source_campaigns))
        if email_report is not None and email_lookup is None:
            if email_report.dc.workbook_placements:
                st.caption("Email placements use the lookup results included in DC Checking.")
            else:
                st.info("Email uses the source campaign labels. Save account databases under Campaign lookup files "
                        "to look up placements automatically on this app.")
        st.caption(
            "GEO and RECO active tickets come from their OVERALL worksheets. Upload a Campaign totals workbook to use "
            "its report date and entered counts. You can also enter Active Ticket and DC Worked counts in Excel."
        )
        st.caption(
            "For Email, SMS and Viber, UNIQUE RESULT counts customer accounts that replied, regardless of receiving inbox "
            "or assigned agent. AGENT WORKED counts accounts with a matched agent reply. In SMS Agent Touch reports, "
            "it counts clients with a logged touch. Each account counts once per channel and campaign."
        )
        if source_campaigns or any(report is not None for _, _, report in channels):
            st.download_button(
                "Download report (.xlsx)",
                data=partial(build_summary_workbook, campaign, report_date, sent_to_ops, viber_report, email_report,
                             campaign_totals, sms_report=sms_report, channel_statuses=statuses, source_campaigns=source_campaigns),
                file_name=f"INBOUND_REPLY_{report_date.isoformat()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="inbound_reply_summary_download", on_click="ignore", type="primary",
            )
            st.caption("Includes Summary and unreplied sheets for SMS, Viber and Email. Click a highlighted cell to read the full text.")
        if unmatched_sources:
            st.download_button(
                "Download unmatched accounts (.xlsx)",
                data=partial(build_unmatched_workbook, unmatched_sources),
                file_name=f"INBOUND_REPLY_UNMATCHED_{report_date.isoformat()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="inbound_reply_unmatched_download", on_click="ignore",
            )
    if sms_report is not None:
        with st.container(border=True, key="inbound_sms_details"):
            st.subheader("SMS account details")
            render_sms_details(sms_report, report_date)
    if email_report is not None:
        with st.container(border=True, key="inbound_email_details"):
            st.subheader("Email account details")
            render_email_details(email_report, report_date)
    if viber_report is not None:
        with st.container(border=True, key="inbound_viber_details"):
            st.subheader("Viber account details")
            render_viber_details(viber_report, report_date)
