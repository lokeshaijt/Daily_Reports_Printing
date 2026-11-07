"""Render processed sheets into a clean, consistently formatted workbook."""

from __future__ import annotations

import io
import math

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import Outline

from .processor import DASH, Sheet, Table, is_number

FONT_NAME = "Calibri"
NAVY = "1F3864"
HEADER_BLUE = "305496"
SECTION_FILL = "D9E1F2"
TOTAL_FILL = "E7E6E6"
GROUP_FILL = "FFF2CC"
ZEBRA_FILL = "F7F9FC"

thin = Side(style="thin", color="A6A6A6")
medium = Side(style="medium", color="404040")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
TOTAL_BORDER = Border(left=thin, right=thin, top=medium, bottom=medium)

TITLE_FONT = Font(name=FONT_NAME, size=14, bold=True, color="FFFFFF")
SECTION_FONT = Font(name=FONT_NAME, size=12, bold=True, color=NAVY)
HEADER_FONT = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
BODY_FONT = Font(name=FONT_NAME, size=10)
BOLD_FONT = Font(name=FONT_NAME, size=10, bold=True)
DETAIL_FONT = Font(name=FONT_NAME, size=10, color="595959")

NUMBER_FORMAT = "#,##0"
DATA_ROW_HEIGHT = 18


def _fill(color: str) -> PatternFill:
    return PatternFill("solid", start_color=color, end_color=color)


def _column_widths(sheet: Sheet) -> dict[int, float]:
    widths: dict[int, float] = {}
    for table in sheet.tables:
        for c, h in enumerate(table.header):
            longest_word = max((len(w) for w in str(h).split()), default=4)
            widths[c] = max(widths.get(c, 0), longest_word + 3, 10)
        for r in table.rows:
            for c, v in enumerate(r.values):
                if v is None:
                    continue
                text = f"{v:,}" if is_number(v) else str(v)
                extra = 4 if r.kind == "detail" and c == 0 else 2
                widths[c] = max(widths.get(c, 0), len(text) + extra)
    return {c: min(w, 50) for c, w in widths.items()}


def _header_height(table: Table, widths: dict[int, float]) -> float:
    lines = 1
    for c, h in enumerate(table.header):
        lines = max(lines, math.ceil(len(str(h)) / max(widths.get(c, 10) - 2, 1)))
    return 15 * min(lines, 4) + 8


def _write_table(ws, table: Table, start_row: int, widths: dict[int, float], zebra: bool) -> int:
    ncols = len(table.header)
    row = start_row

    if table.title:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
        cell = ws.cell(row=row, column=1, value=table.title)
        cell.font = SECTION_FONT
        cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        for c in range(1, ncols + 1):
            ws.cell(row=row, column=c).fill = _fill(SECTION_FILL)
            ws.cell(row=row, column=c).border = BORDER
        ws.row_dimensions[row].height = 22
        row += 1

    for c, h in enumerate(table.header, start=1):
        cell = ws.cell(row=row, column=c, value=h)
        cell.font = HEADER_FONT
        cell.fill = _fill(HEADER_BLUE)
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = _header_height(table, widths)
    header_row = row
    row += 1

    stripe = False
    for r in table.rows:
        if r.kind != "detail":
            stripe = not stripe
        for c, v in enumerate(r.values, start=1):
            cell = ws.cell(row=row, column=c, value=v)
            cell.border = TOTAL_BORDER if r.kind == "total" else BORDER
            numeric = is_number(v) or v == DASH or (c - 1) in table.numeric_cols
            if is_number(v):
                cell.number_format = NUMBER_FORMAT
            if numeric and c > 1:
                cell.alignment = Alignment(horizontal="right", vertical="center", indent=1)
            else:
                indent = 3 if (r.kind == "detail" and c == 1) else 1
                cell.alignment = Alignment(horizontal="left", vertical="center", indent=indent)

            if r.kind == "total":
                cell.font = BOLD_FONT
                cell.fill = _fill(TOTAL_FILL)
            elif r.kind == "group":
                cell.font = BOLD_FONT
                cell.fill = _fill(GROUP_FILL)
            elif r.kind == "detail":
                cell.font = DETAIL_FONT
            else:
                cell.font = BODY_FONT
                if zebra and not stripe:
                    cell.fill = _fill(ZEBRA_FILL)
        ws.row_dimensions[row].height = DATA_ROW_HEIGHT
        if r.kind == "detail":
            ws.row_dimensions[row].outline_level = 1
            ws.row_dimensions[row].hidden = True
        row += 1
    return header_row, row


def write_workbook(sheets: list[Sheet]) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    for sheet in sheets:
        ws = wb.create_sheet(sheet.name[:31])
        ws.sheet_view.showGridLines = False
        widths = _column_widths(sheet)
        max_cols = max((len(t.header) for t in sheet.tables), default=1)
        row = 1

        if sheet.title:
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_cols)
            cell = ws.cell(row=1, column=1, value=sheet.title)
            cell.font = TITLE_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
            for c in range(1, max_cols + 1):
                ws.cell(row=1, column=c).fill = _fill(NAVY)
            ws.row_dimensions[1].height = 30
            row = 3

        has_groups = any(r.kind == "group" for t in sheet.tables for r in t.rows)
        first_header = None
        last_row = row
        for i, table in enumerate(sheet.tables):
            header_row, row = _write_table(ws, table, row, widths, zebra=not has_groups)
            first_header = first_header or header_row
            last_row = row - 1
            if i < len(sheet.tables) - 1:
                row += 2  # breathing space between sections

        for c, w in widths.items():
            ws.column_dimensions[get_column_letter(c + 1)].width = w

        # Single-table sheets (aging / ink stock): freeze the header and add a filter.
        if len(sheet.tables) == 1 and first_header:
            ws.freeze_panes = ws.cell(row=first_header + 1, column=2 if has_groups else 1)
            if not has_groups:
                ncols = len(sheet.tables[0].header)
                ws.auto_filter.ref = f"A{first_header}:{get_column_letter(ncols)}{last_row}"

        if has_groups:
            # Summary row sits above its lots, so the +/- button appears on the summary row.
            ws.sheet_properties.outlinePr = Outline(summaryBelow=False, summaryRight=False)
            ws.sheet_format.outlineLevelRow = 1

        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_options.horizontalCentered = True
        if first_header and len(sheet.tables) == 1:
            ws.print_title_rows = f"{first_header}:{first_header}"

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
