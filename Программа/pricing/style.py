"""Единое оформление листов: заголовок, шапка, зебра, форматы чисел, статусы."""

from __future__ import annotations

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

FONT = "Calibri"


def argb(color: str) -> str:
    """Цвет для Excel 2016: шесть знаков openpyxl пишет с прозрачностью 00, и заливка пропадает."""
    color = color.removeprefix("#").upper()
    if len(color) == 6:
        return "FF" + color
    return color


INK = argb("1F2937")
MUTED = argb("6B7280")
ACCENT = argb("0F4C5C")
ACCENT_SOFT = argb("E6F0F2")
STRIPE = argb("F7F9FA")
LINE = argb("D9DEE3")
GOOD = argb("DCFCE7")
GOOD_INK = argb("166534")
BAD = argb("FEE2E2")
BAD_INK = argb("991B1B")
WARN = argb("FEF3C7")

THIN = Side(style="thin", color=LINE)
BORDER = Border(bottom=THIN)
HEADER_BORDER = Border(bottom=Side(style="medium", color=ACCENT))

NUMBER = '#,##0.00'
INTEGER = '#,##0'
PERCENT = '0%'

STATUS_FILLS = {
    "да": (GOOD, GOOD_INK),
    "нет": (BAD, BAD_INK),
    "пропущен": (WARN, INK),
}


def new_workbook() -> openpyxl.Workbook:
    wb = openpyxl.Workbook()
    wb.active.sheet_view.showGridLines = False
    return wb


def _prepare(ws):
    ws.sheet_view.showGridLines = False


def write_table(
    ws,
    title: str,
    headers: list[str],
    rows: list[list],
    *,
    widths: list[int] | None = None,
    start_row: int = 1,
    status_col: int | None = None,
    percent_cols: list[int] | None = None,
    number_cols: list[int] | None = None,
):
    """Таблица с заголовком над шапкой. Возвращает номер строки шапки."""
    _prepare(ws)
    ws.cell(start_row, 1, title).font = Font(name=FONT, size=14, bold=True, color=INK)
    ws.row_dimensions[start_row].height = 24
    head = start_row + 1
    for col, text in enumerate(headers, 1):
        c = ws.cell(head, col, text)
        c.font = Font(name=FONT, bold=True, color=argb("FFFFFF"))
        c.fill = PatternFill("solid", fgColor=ACCENT)
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        c.border = HEADER_BORDER
    ws.row_dimensions[head].height = 34

    percent_cols = set(percent_cols or [])
    number_cols = set(number_cols or [])
    for i, values in enumerate(rows):
        r = head + 1 + i
        fill = PatternFill("solid", fgColor=STRIPE) if i % 2 else None
        for col, value in enumerate(values, 1):
            c = ws.cell(r, col, value)
            c.font = Font(name=FONT, color=INK)
            c.border = BORDER
            c.alignment = Alignment(vertical="top", wrap_text=isinstance(value, str) and len(value) > 60)
            if fill:
                c.fill = fill
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                if col in percent_cols:
                    c.number_format = PERCENT
                elif col in number_cols or isinstance(value, float):
                    c.number_format = NUMBER
                else:
                    c.number_format = INTEGER
        if status_col and status_col <= len(values):
            c = ws.cell(r, status_col)
            colors = STATUS_FILLS.get(str(values[status_col - 1]).lower())
            if colors:
                c.fill = PatternFill("solid", fgColor=colors[0])
                c.font = Font(name=FONT, bold=True, color=colors[1])
                c.alignment = Alignment(horizontal="center", vertical="top")

    if start_row == 1:
        ws.freeze_panes = ws.cell(head + 1, 1)
        if rows:
            ws.auto_filter.ref = f"A{head}:{get_column_letter(len(headers))}{head + len(rows)}"
    for col, width in enumerate(widths or [], 1):
        current = ws.column_dimensions[get_column_letter(col)].width
        ws.column_dimensions[get_column_letter(col)].width = max(width, current or 0) if current and current != 13 else width
    return head
