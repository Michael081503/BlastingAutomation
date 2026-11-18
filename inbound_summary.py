"""The shared 18-column campaign report for the page and Excel output."""

from __future__ import annotations

from datetime import date
from io import BytesIO

from inbound_reply import InboundReplyError, normalized_campaign, normalized_phone, unreplied_records


INBOUND_SUMMARY_VERSION = 1
PERCENT_COLUMNS = frozenset({3, 5, 7, 13, 15, 17})


def campaign_groups(viber_report=None, email_report=None, sms_report=None) -> list[dict]:
    campaigns = {}
    for channel, report in (("viber", viber_report), ("email", email_report), ("sms", sms_report)):
        if report is not None:
            for counts in report.campaigns:
                group = campaigns.setdefault(normalized_campaign(counts.campaign), {"label": counts.campaign})
                group[channel] = counts
    return sorted(campaigns.values(), key=lambda group: group["label"].casefold())


def campaign_totals_key(report_date: date, campaign: str) -> str:
    return f"{report_date.isoformat()}|{normalized_campaign(campaign)}"


def nonnegative_count(value) -> int | None:
    return value if isinstance(value, int) and not isinstance(value, bool) and value >= 0 else None


def summary_rows(campaign: str, report_date: date, sent_to_ops: str, viber_report=None,
                  email_report=None, campaign_totals: dict | None = None, sms_report=None) -> list[list]:
    rows = []
    for group in campaign_groups(viber_report, email_report, sms_report) or [None]:
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


def _combined_details(viber_report=None, email_report=None):
    """Keep every original source field under a clearly identified source header."""
    fields = []
    records = []
    if email_report is not None:
        for record in email_report.accounts:
            if record.email not in email_report.replied_emails:
                records.append(("EMAIL", record.campaign, "", "", record.email,
                                email_report.dc.headers, record.source_values))
    if viber_report is not None:
        source = (viber_report.message_logs if viber_report.agent_replied_source == "campaign_replies"
                  else viber_report.campaign_replies)
        for record in unreplied_records(viber_report):
            records.append(("VIBER", record.campaign, record.name, normalized_phone(record.phone), "",
                            source.headers, record.source_values))
    for _, _, _, _, _, headers, _ in records:
        for index, header in enumerate(headers):
            # Column position distinguishes repeated headers in an original workbook.
            field = (header or f"Column {index + 1}", sum(item == header for item in headers[:index]))
            if field not in fields:
                fields.append(field)
    rows = []
    for channel, campaign, name, phone, email, headers, values in sorted(
            records, key=lambda row: (normalized_campaign(row[1]), row[0], row[3] or row[4])):
        source_values = {}
        for index, header in enumerate(headers):
            field = (header or f"Column {index + 1}", sum(item == header for item in headers[:index]))
            source_values[field] = values[index] if index < len(values) else None
        rows.append([channel, campaign, name, phone, email, "Not replied",
                     *(source_values.get(field) for field in fields)])
    expected = sum(campaign.awaiting_reply for report in (viber_report, email_report)
                   if report is not None for campaign in report.campaigns)
    if len(rows) != expected:
        raise InboundReplyError("Unreplied account details do not match the campaign counts. Upload the source files again.")
    headers = ["Channel", "Campaign", "Name", "Phone", "Email", "Reply Status",
               *(f"Source: {header}" + (f" ({occurrence + 1})" if occurrence else "") for header, occurrence in fields)]
    return headers, rows


def build_summary_workbook(campaign: str, report_date: date, sent_to_ops: str,
                           viber_report=None, email_report=None, campaign_totals: dict | None = None) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    workbook = Workbook()
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
    for row in summary_rows(campaign, report_date, sent_to_ops, viber_report, email_report, campaign_totals):
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

    details = workbook.create_sheet("Unreplied Details")
    headers, rows = _combined_details(viber_report, email_report)
    details.append(headers)
    for row in rows:
        details.append(row)
    for row in details:
        for cell in row:
            if isinstance(cell.value, str):
                cell.data_type = "s"
                cell.number_format = "@"
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if cell.row == 1:
                cell.font = Font(bold=True)
                cell.fill = PatternFill("solid", fgColor="E2EAFB")
    for index, header in enumerate(headers, 1):
        details.column_dimensions[get_column_letter(index)].width = 60 if "body" in header.casefold() or "content" in header.casefold() else 26
    details.freeze_panes = "A2"
    details.auto_filter.ref = details.dimensions
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()
