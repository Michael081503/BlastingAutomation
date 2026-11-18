"""The shared 18-column campaign report for the page and Excel output."""

from __future__ import annotations

from datetime import date
from io import BytesIO

from inbound_reply import InboundReplyError, normalized_campaign, normalized_phone, unreplied_records


INBOUND_SUMMARY_VERSION = 3
BLAST_PERCENTAGE_FORMULAS_VERSION = 1
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
        if label.strip():
            campaigns.setdefault(normalized_campaign(label), {"label": label.strip()})
    return sorted(campaigns.values(), key=lambda group: group["label"].casefold())


def campaign_totals_key(report_date: date, campaign: str) -> str:
    return f"{report_date.isoformat()}|{normalized_campaign(campaign)}"


def nonnegative_count(value) -> int | None:
    return value if isinstance(value, int) and not isinstance(value, bool) and value >= 0 else None


def summary_rows(campaign: str, report_date: date, sent_to_ops: str, viber_report=None,
                  email_report=None, campaign_totals: dict | None = None, sms_report=None,
                  *, source_campaigns=()) -> list[list]:
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
        totals = (campaign_totals or {}).get(campaign_totals_key(report_date, cells[0]), {})
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


def blast_percentage_formula(row: int, blast_column: str) -> str:
    active, blasted = f"$B{row}", f"{blast_column}{row}"
    return (f'=IF(AND(ISNUMBER({active}),{active}>0,ISNUMBER({blasted}),{blasted}>=0),'
            f'{blasted}/{active},"")')


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
    for channel, campaign, name, phone, email, headers, values in sorted(
            records, key=lambda row: (normalized_campaign(row[1]), row[0], row[3] or row[4])):
        source_values = {}
        for index, header in enumerate(headers):
            field = (header or f"Column {index + 1}", sum(item == header for item in headers[:index]))
            source_values[field] = values[index] if index < len(values) else None
        rows.append([channel, campaign, name, phone, email, "Not replied",
                     *(source_values.get(field) for field in fields)])
    expected = sum(campaign.awaiting_reply for campaign in report.campaigns) if report is not None else 0
    if len(rows) != expected:
        raise InboundReplyError("Unreplied account details do not match the campaign counts. Upload the source files again.")
    headers = ["Channel", "Campaign", "Name", "Phone", "Email", "Reply Status",
               *(f"Source: {header}" + (f" ({occurrence + 1})" if occurrence else "") for header, occurrence in fields)]
    return headers, rows


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
        "Unique customers who emailed you, counted once per email address within each campaign from DC Checking. "
        "Customer replies remain counted after an agent responds. Agent responses appear under AGENT WORKED - EMAIL.",
        "INBOUND REPLY")
    for reply_column, channel, count_column, percent_column in (("J", "SMS", "O", "P"), ("K", "Viber", "Q", "R")):
        sheet[f"{reply_column}3"].comment = Comment(
            f"Unique customers who contacted you through {channel}, counted once per phone number within each campaign. "
            "Contacts found only in the agent-replies file are excluded.", "INBOUND REPLY")
        sheet[f"{count_column}3"].comment = Comment(
            f"Customers from {channel} REPLY who received an agent response in the same campaign.", "INBOUND REPLY")
        sheet[f"{percent_column}3"].comment = Comment(
            f"Customers answered by agents divided by {channel} REPLY × 100.", "INBOUND REPLY")
    sheet["M3"].comment = Comment("Customers from EMAIL REPLY who received an agent response.", "INBOUND REPLY")
    sheet["N3"].comment = Comment("Customers answered by agents divided by EMAIL REPLY × 100.", "INBOUND REPLY")
    for row in summary_rows(campaign, report_date, sent_to_ops, viber_report, email_report, campaign_totals, sms_report,
                            source_campaigns=source_campaigns):
        sheet.append(row)
    edge = Side(style="thin", color="303744")
    for row in sheet:
        for cell in row:
            cell.border = Border(left=edge, right=edge, top=edge, bottom=edge)
            cell.font = Font(name="Arial", size=9, bold=cell.row <= 3, color="242831")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.fill = PatternFill("solid", fgColor="E2EAFB" if cell.row <= 3 else "FFFFFF")
            if isinstance(cell.value, str):
                cell.data_type = "s"
            if cell.row >= 4:
                cell.number_format = "0.00%" if cell.column - 1 in PERCENT_COLUMNS else "#,##0"
    sheet["B2"].comment = Comment("Enter the total active accounts for this campaign manually.", "INBOUND REPLY")
    for blast_column, percent_column, channel in (("C", "D", "Email"), ("E", "F", "SMS"), ("G", "H", "Viber")):
        sheet[f"{blast_column}3"].comment = Comment(f"Accounts blasted through {channel}.", "INBOUND REPLY")
        sheet[f"{percent_column}3"].comment = Comment(
            "Accounts blasted divided by ACTIVE TICKET × 100. Recalculates when the counts change. "
            "Stays blank when ACTIVE TICKET is blank or zero, or the blast count is blank.", "INBOUND REPLY")
        for row in range(4, sheet.max_row + 1):
            sheet[f"{percent_column}{row}"] = blast_percentage_formula(row, blast_column)
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

    for channel, title, report in (("SMS", "SMS Unreplied", sms_report),
                                    ("VIBER", "Viber Unreplied", viber_report),
                                    ("EMAIL", "Email Unreplied", email_report)):
        details = workbook.create_sheet(title)
        headers, rows = _channel_details(channel, report)
        status = "Processed" if report is not None else (channel_statuses or {}).get(channel, "Not uploaded")
        details.append(["Source status", status])
        details.append(["Unreplied accounts", len(rows) if report is not None else None])
        details.append(headers)
        for row in rows:
            details.append(row)
        for row in details:
            for cell in row:
                if isinstance(cell.value, str):
                    cell.data_type = "s"
                    cell.number_format = "@"
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                if cell.row == 3:
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill("solid", fgColor="E2EAFB")
        details["A1"].font = details["A2"].font = Font(bold=True)
        for index, header in enumerate(headers, 1):
            details.column_dimensions[get_column_letter(index)].width = 60 if "body" in header.casefold() or "content" in header.casefold() else 26
        details.freeze_panes = "A4"
        details.auto_filter.ref = f"A3:{get_column_letter(len(headers))}{details.max_row}"
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()
