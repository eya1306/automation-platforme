#!/usr/bin/env python3
"""
compare_excel.py

Compares two Excel files that share the same sheet/column structure
(e.g. "Variable", "Constant", "Type definition") and produces a single
report workbook showing exactly what changed.

For every input sheet, two output sheets are created:
    <SheetName>_before   -> rows that were removed or modified (RED)
    <SheetName>_after    -> rows that were added   or modified (GREEN)

Rows are matched between the two files using a "key" column (by default
the first column, e.g. "Name"). A row is considered:
    - Removed  : key exists in BEFORE but not in AFTER
    - Added    : key exists in AFTER  but not in BEFORE
    - Modified : key exists in both, but at least one cell differs
    - Unchanged: identical in both -> not included in the report

Usage:
    python compare_excel.py before.xlsx after.xlsx report.xlsx
    python compare_excel.py before.xlsx after.xlsx report.xlsx --key 0
"""

import sys
import argparse
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
RED_FONT = Font(color="9C0006")
GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
GREEN_FONT = Font(color="006100")
HEADER_FILL = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
HEADER_FONT = Font(bold=True)
CHANGED_CELL_FONT_MOD = Font(bold=True, color="9C0006")   # for changed cells in "before"
CHANGED_CELL_FONT_MOD_A = Font(bold=True, color="006100")  # for changed cells in "after"


def sheet_to_rows(ws):
    """Return (header, data_rows) for a worksheet, skipping fully empty rows."""
    all_rows = list(ws.iter_rows(values_only=True))
    if not all_rows:
        return [], []
    header = list(all_rows[0])
    data = [list(r) for r in all_rows[1:] if any(c is not None and str(c).strip() != "" for c in r)]
    return header, data


def build_key_map(data, key_idx):
    """Map key value -> row, skipping rows with an empty key."""
    m = {}
    for row in data:
        if key_idx >= len(row):
            continue
        key = row[key_idx]
        if key is None or str(key).strip() == "":
            continue
        m[str(key).strip()] = row
    return m


def normalize(value):
    """Normalize a cell value for comparison (so '1' vs 1, or trailing spaces, don't count as changes)."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()


def rows_equal(row_a, row_b, ncols):
    for i in range(ncols):
        va = row_a[i] if i < len(row_a) else None
        vb = row_b[i] if i < len(row_b) else None
        if normalize(va) != normalize(vb):
            return False
    return True


def diff_cells(row_before, row_after, ncols):
    """Return set of column indices where the two rows differ."""
    changed = set()
    for i in range(ncols):
        vb = row_before[i] if i < len(row_before) else None
        va = row_after[i] if i < len(row_after) else None
        if normalize(vb) != normalize(va):
            changed.add(i)
    return changed


def autofit(ws):
    for col_cells in ws.columns:
        length = 0
        col_letter = None
        for c in col_cells:
            col_letter = c.column_letter
            if c.value is not None:
                length = max(length, len(str(c.value)))
        if col_letter:
            ws.column_dimensions[col_letter].width = min(max(length + 2, 10), 60)
    ws.freeze_panes = "A2"


def write_sheet(wb, name, header, entries, base_fill, base_font, changed_font):
    """
    entries: list of tuples (row_values, status, changed_col_indices, whole_row)
    changed_col_indices: set of column indices that differ from the counterpart row.
    whole_row: True for Added/Removed rows (entire row is new/gone, so the
               whole row is colored); False for Modified rows (only the
               specific changed_col_indices cells get colored, the rest of
               the row is left with no fill).
    """
    ws = wb.create_sheet(name[:31])
    ncols = len(header)
    ws.append(list(header) + ["Change"])
    for c in ws[1]:
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center")

    for row_values, status, changed_cols, whole_row in entries:
        padded = list(row_values) + [None] * (ncols - len(row_values))
        ws.append(padded[:ncols] + [status])
        r = ws.max_row
        for col_idx in range(ncols):
            cell = ws.cell(row=r, column=col_idx + 1)
            if whole_row:
                cell.fill = base_fill
                cell.font = base_font
            elif col_idx in changed_cols:
                cell.fill = base_fill
                cell.font = changed_font
        # Status column stays colored as a quick visual indicator of the row's status
        status_cell = ws.cell(row=r, column=ncols + 1)
        status_cell.fill = base_fill
        status_cell.font = base_font

    autofit(ws)
    return ws


def compare_files(before_path, after_path, output_path, key_column_index=0):
    wb_before = openpyxl.load_workbook(before_path, data_only=True)
    wb_after = openpyxl.load_workbook(after_path, data_only=True)

    common_sheets = [s for s in wb_before.sheetnames if s in wb_after.sheetnames]
    only_in_before = [s for s in wb_before.sheetnames if s not in wb_after.sheetnames]
    only_in_after = [s for s in wb_after.sheetnames if s not in wb_before.sheetnames]

    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)

    summary_rows = []

    for sheet_name in common_sheets:
        header_b, data_b = sheet_to_rows(wb_before[sheet_name])
        header_a, data_a = sheet_to_rows(wb_after[sheet_name])
        header = header_a if header_a else header_b
        ncols = max(len(header_a), len(header_b), 1)

        map_b = build_key_map(data_b, key_column_index)
        map_a = build_key_map(data_a, key_column_index)

        before_entries = []
        after_entries = []
        n_removed = n_added = n_modified = 0

        for key, row_b in map_b.items():
            if key not in map_a:
                before_entries.append((row_b, "Removed", set(), True))
                n_removed += 1
            else:
                row_a = map_a[key]
                if not rows_equal(row_a, row_b, ncols):
                    changed_cols = diff_cells(row_b, row_a, ncols)
                    before_entries.append((row_b, "Modified", changed_cols, False))
                    after_entries.append((row_a, "Modified", changed_cols, False))
                    n_modified += 1

        for key, row_a in map_a.items():
            if key not in map_b:
                after_entries.append((row_a, "Added", set(), True))
                n_added += 1

        if before_entries or after_entries:
            write_sheet(out_wb, f"{sheet_name}_before", header, before_entries, RED_FILL, RED_FONT, CHANGED_CELL_FONT_MOD)
            write_sheet(out_wb, f"{sheet_name}_after", header, after_entries, GREEN_FILL, GREEN_FONT, CHANGED_CELL_FONT_MOD_A)

        summary_rows.append((sheet_name, len(map_b), len(map_a), n_removed, n_added, n_modified))

    summary_ws = out_wb.create_sheet("Summary", 0)
    summary_ws.append(["Sheet", "Rows Before", "Rows After", "Removed", "Added", "Modified"])
    for c in summary_ws[1]:
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
    for row in summary_rows:
        summary_ws.append(row)
    if only_in_before:
        summary_ws.append([])
        summary_ws.append([f"Sheets only in BEFORE file: {', '.join(only_in_before)}"])
    if only_in_after:
        summary_ws.append([f"Sheets only in AFTER file: {', '.join(only_in_after)}"])
    autofit(summary_ws)

    out_wb.save(output_path)
    return summary_rows

def main():
    parser = argparse.ArgumentParser(description="Compare two Excel files and highlight changes.")
    parser.add_argument("before", help="Path to the BEFORE .xlsx file")
    parser.add_argument("after", help="Path to the AFTER .xlsx file")
    parser.add_argument("output", help="Path to save the report .xlsx file")
    parser.add_argument("--key", type=int, default=0, help="0-based index of the key/identifier column (default: 0, i.e. first column, e.g. 'Name')")
    args = parser.parse_args()

    summary = compare_files(args.before, args.after, args.output, key_column_index=args.key)

    print(f"Report saved to: {args.output}\n")
    print(f"{'Sheet':<25}{'Removed':>10}{'Added':>10}{'Modified':>10}")
    for row in summary:
        name, _, _, removed, added, modified = row
        print(f"{name:<25}{removed:>10}{added:>10}{modified:>10}")


if __name__ == "__main__":
    main()