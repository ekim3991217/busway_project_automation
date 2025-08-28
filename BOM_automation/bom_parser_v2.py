#!/usr/bin/env python3
"""
bom_parser.py

Usage:
  python bom_parser.py

What it does:
  1) Prompts user to paste an Excel file path.
  2) Deletes worksheet named "REFERENCE" (case-insensitive match).
  3) On the "BOM" sheet, UNMERGES all merged cells and resets horizontally centered cells to 'general'.
  4) On the "BOM" sheet, removes columns that are completely empty between rows 7–77 (inclusive).
  5) On the "BOM" sheet, removes rows 1:3.
  6) Writes an edited copy next to the original with "_editedBOM" appended to the filename.
"""

import sys
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment


def is_cell_empty(value):
    """Treat None or whitespace-only strings as empty."""
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def main():
    # --- Filename prompt section ---
    file_input = input("COPY AND PASTE EXCEL FILEPATH HERE: ").strip('"').strip()
    src_path = Path(file_input).expanduser().resolve()

    if not src_path.exists():
        print(f"[ERROR] File not found: {src_path}")
        sys.exit(1)
    if src_path.suffix.lower() != ".xlsx":
        print(f"[ERROR] This script expects a .xlsx file. Got: {src_path.suffix}")
        sys.exit(1)

    # --- Workbook processing ---
    try:
        wb = load_workbook(filename=str(src_path), data_only=False)
    except Exception as e:
        print(f"[ERROR] Failed to open workbook: {e}")
        sys.exit(1)

    # 1) Delete worksheet "REFERENCE" (case-insensitive)
    ref_sheet_name = None
    for name in wb.sheetnames:
        if name.casefold() == "reference":
            ref_sheet_name = name
            break
    if ref_sheet_name:
        del wb[ref_sheet_name]
        print(f"[INFO] Deleted worksheet: {ref_sheet_name}")
    else:
        print("[INFO] No 'REFERENCE' worksheet found (skipping).")

    # 2) & 3) Operate on "BOM" sheet
    bom_sheet_name = None
    for name in wb.sheetnames:
        if name.casefold() == "bom":
            bom_sheet_name = name
            break

    if not bom_sheet_name:
        print("[ERROR] 'BOM' worksheet not found. No changes made to data sheets.")
        sys.exit(1)

    ws = wb[bom_sheet_name]

    # --- NEW: Undo all merged cells and horizontal centering on BOM sheet ---
    # Unmerge all merged ranges
    merged_count = len(ws.merged_cells.ranges)
    if merged_count:
        for m in list(ws.merged_cells.ranges):
            ws.unmerge_cells(range_string=m.coord)
        print(f"[INFO] Unmerged {merged_count} merged cell range(s).")
    else:
        print("[INFO] No merged cells found to unmerge.")

    # Reset horizontally centered cells to 'general' (keep vertical alignment as-is)
    center_fixed = 0
    max_row = ws.max_row
    max_col = ws.max_column
    for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            al = cell.alignment
            # If horizontal is 'center' or 'centerContinuous', reset to general
            if al and al.horizontal and ("center" in al.horizontal):
                cell.alignment = Alignment(horizontal="general", vertical=al.vertical)
                center_fixed += 1
    print(f"[INFO] Reset horizontal centering on {center_fixed} cell(s).")

    # --- UPDATED: Remove columns empty ONLY within rows 7..77 (inclusive) ---
    window_start = 7
    window_end = min(77, ws.max_row)  # don't exceed actual last row
    cols_deleted = []
    # Iterate right->left to avoid index shifting
    for col_idx in range(ws.max_column, 0, -1):
        # If the BOM has fewer than 7 rows, there's no window to check
        if ws.max_row < window_start:
            break
        empty_in_window = True
        for row_idx in range(window_start, window_end + 1):
            cell_val = ws.cell(row=row_idx, column=col_idx).value
            if not is_cell_empty(cell_val):
                empty_in_window = False
                break
        if empty_in_window:
            ws.delete_cols(col_idx, 1)
            cols_deleted.append(get_column_letter(col_idx))

    if cols_deleted:
        print(f"[INFO] Deleted columns empty in rows 7–77: {', '.join(reversed(cols_deleted))}")
    else:
        print("[INFO] No columns empty across rows 7–77 found to delete.")

    # Remove rows 1:3
    ws.delete_rows(idx=1, amount=3)
    print("[INFO] Deleted rows 1:3 on 'BOM'.")

    # 4) Save edited file alongside original
    out_path = src_path.with_name(f"{src_path.stem}_editedBOM{src_path.suffix}")
    try:
        wb.save(str(out_path))
    except Exception as e:
        print(f"[ERROR] Failed to save edited workbook: {e}")
        sys.exit(1)

    print(f"[SUCCESS] Edited workbook saved to:\n{out_path}")


if __name__ == "__main__":
    main()
