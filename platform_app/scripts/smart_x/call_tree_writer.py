"""Writes extracted call-tree rows into the Call Tree workbook.

Lifted out of `extraction_code/gui_simple.py` in the upstream project, where
this logic lived on the tkinter window as `_write_to_excel` and
`_write_rows_to_call_tree`. The layout, styling and column mapping are kept
exactly as they were — headers on row 2 in columns C..I, data from row 3,
banded fills, autosized columns — so a workbook produced here opens looking the
same as one produced by the desktop tool.

What changed: the `self._log(...)` calls became a returned row count, and the
"copy the empty template if the target is missing" step is the caller's job.
"""

from __future__ import annotations

from pathlib import Path
from typing import List, Sequence

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

SHEET_NAME = "Call Tree"

HEADERS = [
    "CSC Calling", "CSU Calling", "Calling Function",
    "CSC Called", "CSU Called", "Called Function", "Condition",
]

# The empty workbook the desktop tool ships as its starting structure.
EMPTY_TEMPLATE = Path(__file__).resolve().parent / "assets" / "Call_Tree_EMPTY.xlsx"


def write_call_tree(excel_path: Path, rows: Sequence[Sequence], 
                    sheet_name: str = SHEET_NAME) -> int:
    """Write `rows` into `excel_path`, creating the workbook if needed.

    Returns the number of data rows written.
    """
    excel_path = Path(excel_path)

    try:
        wb = openpyxl.load_workbook(excel_path)
    except Exception:
        wb = openpyxl.Workbook()
        if wb.sheetnames:
            del wb[wb.sheetnames[0]]

    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.create_sheet(sheet_name)

    # Clear anything already below the header row before refilling.
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.value = None

    _write_rows(ws, rows)

    # The desktop tool keeps only the target sheet in the output.
    for name in list(wb.sheetnames):
        if name != sheet_name:
            wb.remove(wb[name])

    wb.save(excel_path)
    return len(rows)


def _write_rows(ws, rows: Sequence[Sequence]) -> None:
    try:
        ws.views.sheetView[0].showGridLines = True
    except Exception:
        pass

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    header_border = Border(
        left=Side(style="thin", color="1B365D"),
        right=Side(style="thin", color="1B365D"),
        top=Side(style="medium", color="1F4E78"),
        bottom=Side(style="medium", color="1F4E78"),
    )

    ws.row_dimensions[2].height = 26

    for col_idx, header in enumerate(HEADERS, start=3):
        cell = ws.cell(row=2, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_align
        cell.border = header_border

    data_font = Font(name="Segoe UI", size=10)
    even_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    odd_fill = PatternFill(start_color="F7F9FB", end_color="F7F9FB", fill_type="solid")

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    max_widths = {col: len(HEADERS[col - 3]) for col in range(3, 10)}

    for r_idx, row in enumerate(rows, start=3):
        ws.row_dimensions[r_idx].height = 20
        row_fill = even_fill if (r_idx % 2 == 0) else odd_fill

        for col_offset, value in enumerate(row, start=3):
            if col_offset > 9:
                break
            text = str(value).strip() if (value is not None and str(value).strip() != "") else "None"
            cell = ws.cell(row=r_idx, column=col_offset, value=text)
            cell.fill = row_fill
            cell.font = data_font
            cell.border = thin_border
            cell.alignment = align_center if col_offset in (3, 4, 6, 7) else align_left

            if len(text) > max_widths[col_offset]:
                max_widths[col_offset] = len(text)

    for col_idx in range(3, 10):
        ws.column_dimensions[get_column_letter(col_idx)].width = max(max_widths[col_idx] + 5, 14)
