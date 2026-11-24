"""The shared 18-column campaign report for the page and Excel output."""

from __future__ import annotations

from datetime import date
from io import BytesIO
from xml.etree import ElementTree
from zipfile import ZipFile

from inbound_email import EmailSource
from inbound_reply import (
    INBOUND_MESSAGE_COLUMN_WIDTH, MINDANAO_CAMPAIGN, PL_RECOVERY_CAMPAIGN, InboundReplyError, ReplySource,
    append_report_detail_rows, fit_report_text,
    inbound_campaign_label, normalized_campaign, normalized_phone, unreplied_records,
)


INBOUND_SUMMARY_VERSION = 8
BLAST_PERCENTAGE_FORMULAS_VERSION = 1
OVERALL_BLAST_COVERAGE_VERSION = 1
OVERALL_BLAST_LABEL = "TOTAL — ALL CAMPAIGNS"
CUSTOMER_EMAIL_REPLY_LABELS_VERSION = 1
ALL_CHANNEL_REPLY_LABELS_VERSION = 1
CHANNEL_UNREPLIED_SHEETS_VERSION = 1
SOURCE_CAMPAIGN_LABELS_VERSION = 1
PERCENT_COLUMNS = frozenset({3, 5, 7, 13, 15, 17})


def campaign_groups(viber_report=None, email_report=None, sms_report=None, *, source_campaigns=()) -> list[dict]:
    campaigns = {}
    for channel, report in (("viber", viber_report), ("email", email_report), ("sms", sms_report)):
        if report is not None:
            for counts in report.campaigns:
                group = campaigns.setdefault(normalized_campaign(counts.campaign), {"label": counts.campaign})
                group[channel] = counts
    for label in source_campaigns:
        label = inbound_campaign_label(label.strip())
        if label:
            campaigns.setdefault(normalized_campaign(label), {"label": label})
    return sorted(campaigns.values(), key=lambda group: group["label"].casefold())


def campaign_totals_key(report_date: date, campaign: str) -> str:
    return f"{report_date.isoformat()}|{normalized_campaign(campaign)}"


def nonnegative_count(value) -> int | None:
    return value if isinstance(value, int) and not isinstance(value, bool) and value >= 0 else None


def merge_mindanao_totals(report_date: date, campaign_totals: dict | None) -> dict:
    """Fold the previous Region 3 totals into Mindanao for this report date."""
    totals = dict(campaign_totals or {})
    region_totals = totals.pop(campaign_totals_key(report_date, "SBC Region 3"), None)
    if not region_totals:
        return totals
    mindanao_key = campaign_totals_key(report_date, MINDANAO_CAMPAIGN)
    mindanao_totals = totals.get(mindanao_key, {})
    combined = {}
    for field in mindanao_totals.keys() | region_totals.keys():
        counts = [count for source in (mindanao_totals, region_totals)
                  if (count := nonnegative_count(source.get(field))) is not None]
        combined[field] = sum(counts) if counts else None
    totals[mindanao_key] = combined
    return totals


def summary_rows(campaign: str, report_date: date, sent_to_ops: str, viber_report=None,
                  email_report=None, campaign_totals: dict | None = None, sms_report=None,
                  *, source_campaigns=()) -> list[list]:
    campaign_totals = merge_mindanao_totals(report_date, campaign_totals)
    rows = []
    for group in campaign_groups(viber_report, email_report, sms_report, source_campaigns=source_campaigns) or [None]:
        cells = [campaign.strip() or "—", *([None] * 10), sent_to_ops, *([None] * 6)]
        if group is not None:
            cells[0] = group["label"]
            for channel, indexes in (("email", (8, 12, 13)), ("sms", (9, 14, 15)), ("viber", (10, 16, 17))):
                counts = group.get(channel)
                if counts is not None:
                    reply_index, agent_index, percent_index = indexes
                    cells[reply_index] = counts.unique_people
                    cells[agent_index] = counts.agent_replied
                    cells[percent_index] = counts.agent_percent
        totals = campaign_totals.get(campaign_totals_key(report_date, cells[0]), {})
        active = nonnegative_count(totals.get("active_ticket"))
        cells[1] = active
        for channel, index in (("email", 2), ("sms", 4), ("viber", 6)):
            worked = nonnegative_count(totals.get(f"dc_{channel}"))
            cells[index] = worked
            cells[index + 1] = worked / active if worked is not None and active else None
        rows.append(cells)
    return rows


def display_cell(value, column: int) -> str:
    if value is None:
        return "—"
    if column in PERCENT_COLUMNS:
        return f"{value:.2%}"
    return f"{value:,}" if isinstance(value, int) else str(value)


def overall_blast_row(rows: list[list]) -> list:
    """Sum current campaigns; calculate coverage only from complete entered counts."""
    totals = [OVERALL_BLAST_LABEL, *([None] * 17)]
    counts = {column: [nonnegative_count(row[column]) for row in rows] for column in (1, 2, 4, 6)}
    for column, values in counts.items():
        known = [value for value in values if value is not None]
        totals[column] = sum(known) if known else None
    if rows and all(value is not None for value in counts[1]) and totals[1]:
        for column in (2, 4, 6):
            if all(value is not None for value in counts[column]):
                totals[column + 1] = totals[column] / totals[1]
    return totals


def overall_blast_formula(first_row: int, last_row: int, blast_column: str) -> str:
    active = f"$B${first_row}:$B${last_row}"
    blasted = f"{blast_column}${first_row}:{blast_column}${last_row}"
    return (f'=IF(AND(COUNT({active})=ROWS({active}),MIN({active})>=0,SUM({active})>0,'
            f'COUNT({blasted})=ROWS({blasted}),MIN({blasted})>=0),SUM({blasted})/SUM({active}),"")')


def percentage_formula(numerator: str, denominator: str) -> str:
    return (f'=IF(AND(ISNUMBER({denominator}),{denominator}>0,ISNUMBER({numerator}),{numerator}>=0),'
            f'{numerator}/{denominator},"")')


def _cache_summary_percentages(workbook: BytesIO, percentages: dict[str, float]) -> bytes:
    """Keep calculated rates visible in workbook previews before Excel recalculates."""
    if not percentages:
        return workbook.getvalue()
    output = BytesIO()
    with ZipFile(workbook) as original, ZipFile(output, "w") as updated:
        for member in original.infolist():
            data = original.read(member)
            if member.filename == "xl/worksheets/sheet1.xml":
                summary = ElementTree.fromstring(data)
                for cell in summary.iterfind("{*}sheetData/{*}row/{*}c"):
                    rate = percentages.get(cell.get("r"))
                    if rate is not None:
                        cell.find("{*}v").text = str(rate)
                data = ElementTree.tostring(summary, encoding="utf-8")
            updated.writestr(member, data)
    return output.getvalue()


def _channel_details(channel: str, report):
    """Keep a channel's original account fields and validate its unreplied total."""
    fields = []
    records = []
    source_headers = ()
    if report is not None and channel == "EMAIL":
        source_headers = report.dc.headers
        for record in report.accounts:
            if record.email not in report.replied_emails:
                records.append(("EMAIL", record.campaign, "", "", record.email,
                                source_headers, record.source_values))
    elif report is not None:
        source = (report.message_logs if report.agent_replied_source == "campaign_replies"
                  else report.campaign_replies)
        source_headers = source.headers
        for record in unreplied_records(report):
            records.append((channel, record.campaign, record.name, normalized_phone(record.phone), "",
                            source.headers, record.source_values))
    for index, header in enumerate(source_headers):
        # Column position distinguishes repeated headers in an original workbook.
        fields.append((header or f"Column {index + 1}", sum(item == header for item in source_headers[:index])))
    rows = []
    reply_status = "No agent touch" if channel == "SMS" and report is not None and report.client_touches is not None else "Not replied"
    for channel, campaign, name, phone, email, headers, values in sorted(
            records, key=lambda row: (normalized_campaign(row[1]), row[0], row[3] or row[4])):
        source_values = {}
        for index, header in enumerate(headers):
            field = (header or f"Column {index + 1}", sum(item == header for item in headers[:index]))
            source_values[field] = values[index] if index < len(values) else None
        rows.append([channel, campaign, name, phone, email, reply_status,
                     *(source_values.get(field) for field in fields)])
    expected = sum(campaign.awaiting_reply for campaign in report.campaigns) if report is not None else 0
    if len(rows) != expected:
        raise InboundReplyError("Unreplied account details do not match the campaign counts. Upload the source files again.")
    headers = ["Channel", "Campaign", "Name", "Phone", "Email", "Reply Status",
               *(f"Source: {header}" + (f" ({occurrence + 1})" if occurrence else "") for header, occurrence in fields)]
    return headers, rows


def build_client_accounts_workbook(report) -> bytes:
    """Export the full unique Client list and its agent touch status."""
    from openpyxl import Workbook

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Client Accounts"
    headers = ("Reported Campaign", "Agent Touch", *report.message_logs.headers, "Placement Source", "Matched Database Phone")
    sheet.append(headers)
    rows = []
    for row, touched in report.client_touches:
        placement_source = "Client file (unverified)"
        if row.campaign == PL_RECOVERY_CAMPAIGN:
            placement_source = PL_RECOVERY_CAMPAIGN
        elif row.placement_phone:
            placement_source = "Database phone"
        elif row.placement_issue:
            placement_source = row.placement_issue
        values = (row.source_values + (None,) * len(report.message_logs.headers))[:len(report.message_logs.headers)]
        rows.append((row.campaign, "Touched" if touched else "No touch", *values, placement_source, row.placement_phone))
    full_text = append_report_detail_rows(sheet, rows)
    _format_detail_sheet(sheet, headers, header_row=1)
    fit_report_text(workbook, {sheet.title: full_text}, separate_text=True)
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def _format_detail_sheet(sheet, headers, *, header_row: int) -> None:
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    for row in sheet:
        for cell in row:
            if isinstance(cell.value, str):
                cell.data_type = "s"
                cell.number_format = "@"
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if cell.row == header_row:
                cell.font = Font(bold=True)
                cell.fill = PatternFill("solid", fgColor="E2EAFB")
    for index, header in enumerate(headers, 1):
        key = header.casefold()
        sheet.column_dimensions[get_column_letter(index)].width = (
            INBOUND_MESSAGE_COLUMN_WIDTH if "body" in key or "content" in key else 26)
    sheet.freeze_panes = f"A{header_row + 1}"
    sheet.auto_filter.ref = f"A{header_row}:{get_column_letter(len(headers))}{sheet.max_row}"


def build_unmatched_workbook(sources: dict[str, EmailSource | ReplySource]) -> bytes:
    """Export the first original row per customer contact without a lookup match."""
    from openpyxl import Workbook

    workbook = Workbook()
    default = workbook.active
    detail_text = {}
    for channel, title in (("SMS", "SMS Unmatched"), ("VIBER", "Viber Unmatched"), ("EMAIL", "Email Unmatched")):
        source = sources.get(channel)
        if source is None or not source.unmatched_records:
            continue
        original_headers = source.headers[:-1]  # The final header was added by the lookup.
        if isinstance(source, EmailSource):
            contact_header = "Normalized Email"
            contacts = (row.email for row in source.unmatched_records)
        else:
            contact_header = "Normalized Phone"
            contacts = (normalized_phone(row.phone) for row in source.unmatched_records)
        sheet = workbook.create_sheet(title)
        headers = (contact_header, *original_headers)
        client_records = isinstance(source, ReplySource) and source.layout == "client"
        if client_records:
            headers += ("Phones Checked", "Placement Issue")
        sheet.append(headers)
        rows = []
        for contact, row in zip(contacts, source.unmatched_records):
            values = (row.source_values + (None,) * len(original_headers))[:len(original_headers)]
            output_row = (contact, *values)
            if client_records:
                output_row += (", ".join(row.lookup_phones), row.placement_issue)
            rows.append(output_row)
        detail_text[title] = append_report_detail_rows(sheet, rows)
        _format_detail_sheet(sheet, headers, header_row=1)
    if len(workbook.worksheets) > 1:
        workbook.remove(default)
    else:
        default.title = "Unmatched accounts"
        default.append(("No unmatched accounts",))
    fit_report_text(workbook, detail_text)
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def build_summary_workbook(campaign: str, report_date: date, sent_to_ops: str,
                           viber_report=None, email_report=None, campaign_totals: dict | None = None,
                           sms_report=None, channel_statuses: dict | None = None, *, source_campaigns=()) -> bytes:
    from openpyxl import Workbook
    from openpyxl.comments import Comment
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.workbook.properties import CalcProperties
    from openpyxl.utils import get_column_letter

    workbook = Workbook()
    workbook.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True, forceFullCalc=True)
    sheet = workbook.active
    sheet.title = "Summary"
    for cell_range in ("A1:K1", "L1:L3", "M1:R2", "A2:A3", "B2:B3", "C2:H2", "I2:K2"):
        sheet.merge_cells(cell_range)
    labels = {
        "A1": f"{report_date.strftime('%B').upper()} {report_date.day}, {report_date.year} DATA",
        "L1": "SENT TO OPS", "M1": "AGENT WORKED", "A2": "CAMPAIGN", "B2": "ACTIVE TICKET",
        "C2": "DC WORKED", "I2": "UNIQUE RESULT",
        "C3": "EMAIL", "D3": "%", "E3": "SMS", "F3": "%", "G3": "VIBER", "H3": "%",
        "I3": "EMAIL REPLY", "J3": "SMS REPLY", "K3": "VIBER REPLY",
        "M3": "EMAIL", "N3": "%", "O3": "SMS", "P3": "%", "Q3": "VIBER", "R3": "%",
    }
    for address, value in labels.items():
        sheet[address] = value
    sheet["I3"].comment = Comment(
        "Unique customer accounts that replied by email, regardless of receiving mailbox or assigned agent. "
        "Repeated replies count once per customer email address within each campaign from DC Checking. "
        "Customer replies remain counted after an agent responds. Agent responses appear under AGENT WORKED - EMAIL.",
        "INBOUND REPLY")
    for reply_column, channel, count_column, percent_column in (("J", "SMS", "O", "P"), ("K", "Viber", "Q", "R")):
        sheet[f"{reply_column}3"].comment = Comment(
            f"Unique customer accounts that replied through {channel}, regardless of receiving inbox or assigned agent. "
            "Repeated replies count once per phone number within each campaign. "
            "Contacts found only in the agent-replies file are excluded.", "INBOUND REPLY")
        sheet[f"{count_column}3"].comment = Comment(
            f"Customers from {channel} REPLY who received a response from any agent in the same campaign. "
            "Repeated agent replies count once per phone number within each campaign.", "INBOUND REPLY")
        sheet[f"{percent_column}3"].comment = Comment(
            f"Customers answered by agents divided by {channel} REPLY × 100. Recalculates when the counts change.",
            "INBOUND REPLY")
    if sms_report is not None and sms_report.client_touches is not None:
        sheet["O3"].comment = Comment("Unique SMS clients with a matching Agent Touch record.", "INBOUND REPLY")
        sheet["P3"].comment = Comment("Clients touched divided by unique SMS clients.", "INBOUND REPLY")
    sheet["M3"].comment = Comment(
        "Customers from EMAIL REPLY who received a response from any agent. "
        "Repeated agent replies count once per customer email address within each campaign.", "INBOUND REPLY")
    sheet["N3"].comment = Comment(
        "Customers answered by agents divided by EMAIL REPLY × 100. Recalculates when the counts change.",
        "INBOUND REPLY")
    campaign_rows = summary_rows(campaign, report_date, sent_to_ops, viber_report, email_report, campaign_totals, sms_report,
                                 source_campaigns=source_campaigns)
    for row in campaign_rows:
        sheet.append(row)
    last_campaign_row = sheet.max_row
    overall = overall_blast_row(campaign_rows)
    sheet.append(overall)
    total_row = sheet.max_row
    edge = Side(style="thin", color="303744")
    for row in sheet:
        for cell in row:
            cell.border = Border(left=edge, right=edge, top=edge, bottom=edge)
            cell.font = Font(name="Arial", size=9, bold=cell.row <= 3 or cell.row == total_row, color="242831")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.fill = PatternFill("solid", fgColor="E2EAFB" if cell.row <= 3 else
                                    "E8F2EC" if cell.row == total_row else "FFFFFF")
            if isinstance(cell.value, str):
                cell.data_type = "s"
            if cell.row >= 4:
                cell.number_format = "0.00%" if cell.column - 1 in PERCENT_COLUMNS else "#,##0"
    sheet["B2"].comment = Comment("Enter the total active accounts for this campaign manually.", "INBOUND REPLY")
    for column in ("B", "C", "E", "G"):
        values = f"{column}4:{column}{last_campaign_row}"
        sheet[f"{column}{total_row}"] = f'=IF(COUNT({values})=0,"",SUM({values}))'
    sheet[f"A{total_row}"].comment = Comment(
        "Overall blast coverage uses total DC WORKED divided by total ACTIVE TICKET across the campaigns above. "
        "Each channel's percentage stays blank until every campaign has ACTIVE TICKET and that channel's blast count. "
        "Enter zero when a campaign has no blasts for that channel.", "INBOUND REPLY")
    for blast_column, percent_column, channel in (("C", "D", "Email"), ("E", "F", "SMS"), ("G", "H", "Viber")):
        sheet[f"{blast_column}3"].comment = Comment(f"Accounts blasted through {channel}.", "INBOUND REPLY")
        sheet[f"{percent_column}3"].comment = Comment(
            "Accounts blasted divided by ACTIVE TICKET × 100. Recalculates when the counts change. "
            "Stays blank when ACTIVE TICKET is blank or zero, or the blast count is blank.", "INBOUND REPLY")
        for row in range(4, last_campaign_row + 1):
            sheet[f"{percent_column}{row}"] = percentage_formula(f"{blast_column}{row}", f"$B{row}")
        sheet[f"{percent_column}{total_row}"] = overall_blast_formula(4, last_campaign_row, blast_column)
    for reply_column, agent_column, percent_column in (("I", "M", "N"), ("J", "O", "P"), ("K", "Q", "R")):
        for row in range(4, last_campaign_row + 1):
            sheet[f"{percent_column}{row}"] = percentage_formula(f"{agent_column}{row}", f"{reply_column}{row}")
    for column in range(1, 19):
        sheet.column_dimensions[get_column_letter(column)].width = 36 if column == 1 else 18 if column == 12 else 12
    for row in range(1, sheet.max_row + 1):
        sheet.row_dimensions[row].height = 20 if row <= 3 else 18
    sheet.freeze_panes = "C4"
    sheet.sheet_view.showGridLines = False
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.print_title_rows = "1:3"
    sheet.print_area = f"A1:R{sheet.max_row}"

    detail_text = {}
    for channel, title, report in (("SMS", "SMS Unreplied", sms_report),
                                    ("VIBER", "Viber Unreplied", viber_report),
                                    ("EMAIL", "Email Unreplied", email_report)):
        details = workbook.create_sheet(title)
        status = "Processed" if report is not None else (channel_statuses or {}).get(channel, "Not uploaded")
        if channel == "EMAIL" and report is not None and report.history is None:
            status = "Incomplete: upload Email History in EMAIL file 1"
            report = None
        headers, rows = _channel_details(channel, report)
        details.append(["Source status", status])
        details.append(["Unreplied accounts", len(rows) if report is not None else None])
        details.append(headers)
        detail_text[title] = append_report_detail_rows(details, rows)
        _format_detail_sheet(details, headers, header_row=3)
        details["A1"].font = details["A2"].font = Font(bold=True)
    fit_report_text(workbook, detail_text)
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    rates = {}
    for row_number, values in enumerate((*campaign_rows, overall), start=4):
        for column in PERCENT_COLUMNS:
            if values[column] is not None:
                rates[f"{get_column_letter(column + 1)}{row_number}"] = values[column]
    return _cache_summary_percentages(output, rates)
