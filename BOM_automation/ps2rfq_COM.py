#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
ps_cleaner_preserve_images.py

Usage:
  python ps_cleaner_preserve_images.py "C:\path\to\250826 PS-USA-2444 EGLINTON PJT-A.xlsm"

If you omit the argument, you’ll be prompted for a path.

What it does:
  1) Deletes worksheet "COVER" (case-insensitive).
  2) Converts ALL cells on ALL worksheets to values (removes formulas/dependencies) while preserving formatting.
  3) On "PS" sheet, clears columns L through XFD (content + formats), leaving images/shapes untouched.
  4) Saves a copy with filename prefix changed from "yymmdd PS-" to "yymmdd RFP-", preserving extension.
     - .xlsx stays .xlsx (FileFormat=51)
     - .xlsm stays .xlsm and macros are preserved (FileFormat=52)

Requires:
  - Windows with Microsoft Excel installed
  - pywin32 (win32com)
"""

import sys
import re
from pathlib import Path

try:
    import win32com.client as win32
except ImportError:
    print("[ERROR] pywin32 is required. Install with: pip install pywin32")
    sys.exit(1)

# Excel FileFormat constants
XL_OPENXML_WORKBOOK = 51       # .xlsx
XL_OPENXML_WORKBOOK_MACRO = 52 # .xlsm


def get_src_path():
    if len(sys.argv) > 1:
        raw = sys.argv[1]
    else:
        raw = input("COPY AND PASTE EXCEL FILEPATH HERE: ")
    p = Path(raw.strip().strip('"').strip("'")).expanduser().resolve()
    if not p.exists():
        print(f"[ERROR] File not found: {p}")
        sys.exit(1)
    if p.suffix.lower() not in (".xlsx", ".xlsm"):
        print(f"[ERROR] Expected .xlsx or .xlsm file. Got: {p.suffix}")
        sys.exit(1)
    return p


def rename_to_rfp(src_path: Path) -> Path:
    """
    Replace leading 'yymmdd PS-' with 'yymmdd RFP-'.
    If not matched exactly, try inserting ' RFP-' after a 6-digit date at start.
    Fallback: append _RFP before extension.
    """
    name = src_path.name
    m = re.match(r"^(\d{6})\sPS-(.*)$", name, flags=re.IGNORECASE)
    if m:
        return src_path.with_name(f"{m.group(1)} RFP-{m.group(2)}")

    m2 = re.match(r"^(\d{6})(\s?)(.*)$", name)
    if m2 and "RFP-" not in name.upper():
        return src_path.with_name(f"{m2.group(1)} RFP-{m2.group(3)}")

    return src_path.with_name(f"{src_path.stem}_RFP{src_path.suffix}")


def convert_all_used_ranges_to_values(xl_wb):
    """
    For every worksheet: freeze formulas to values by writing UsedRange.Value back to itself.
    This preserves formatting (fonts, fills, borders, column widths, row heights) and images.
    """
    for ws in xl_wb.Worksheets:
        used = ws.UsedRange
        # Handle 1-cell vs multi-cell ranges
        try:
            vals = used.Value  # tuple-of-tuples for multi-cell, scalar for single-cell
            used.Value = vals
        except Exception as e:
            # Some sheets (e.g., charts) or rare cases may not expose UsedRange.Value cleanly.
            # Silently continue; those sheets likely don't need conversion.
            pass


def delete_cover_if_present(xl_wb):
    # Find COVER case-insensitively
    cover = None
    for ws in xl_wb.Worksheets:
        if str(ws.Name).lower() == "cover":
            cover = ws
            break
    if cover:
        cover.Delete()


def clear_ps_columns_from_L_right(xl_wb):
    # Find "PS" sheet case-insensitively
    ps_ws = None
    for ws in xl_wb.Worksheets:
        if str(ws.Name).lower() == "ps":
            ps_ws = ws
            break

    if ps_ws is None:
        print("[ERROR] 'PS' worksheet not found.")
        sys.exit(1)

    # Clear columns L to XFD (content + formats), leaving images/shapes untouched.
    # Range.Clear clears values, formats, comments, etc., but not shapes.
    ps_ws.Range("L:XFD").Clear()


def main():
    src = get_src_path()
    out = rename_to_rfp(src)

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    wb = None

    try:
        keep_macros = src.suffix.lower() == ".xlsm"
        fileformat = XL_OPENXML_WORKBOOK_MACRO if keep_macros else XL_OPENXML_WORKBOOK

        wb = excel.Workbooks.Open(str(src))

        # 1) Delete COVER
        delete_cover_if_present(wb)

        # 2) Convert formulas to values on ALL sheets
        convert_all_used_ranges_to_values(wb)

        # 3) PS: clear L..XFD
        clear_ps_columns_from_L_right(wb)

        # 4) Save a copy with new name (preserve extension and macros)
        wb.SaveAs(Filename=str(out), FileFormat=fileformat)
        print(f"[SUCCESS] Saved edited file to:\n{out}")

    except Exception as e:
        print(f"[ERROR] {e}")
        sys.exit(1)
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        # restore settings (best effort)
        try:
            excel.ScreenUpdating = True
            excel.DisplayAlerts = True
            excel.Visible = False
        except Exception:
            pass
        excel.Quit()


if __name__ == "__main__":
    main()
