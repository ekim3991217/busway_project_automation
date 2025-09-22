#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
ps_cleaner_preserve_images.py

Usage:
  python ps_cleaner_preserve_images.py "C:\\path\\to\\250826 PS-USA-2444 EGLINTON PJT-A.xlsm"

If you omit the argument, you’ll be prompted for a path.

What it does:
  1) Deletes worksheet "COVER" (case-insensitive).
  2) Converts ALL cells on ALL worksheets to values (preserves formatting).
  3) Replaces all occurrences of "DDP" with "CIF" across all worksheets
     (partial match safe: only the substring "DDP" is replaced).
  4) On "PS" sheet:
       - Temporarily moves all shapes/pictures to a safe area in column A
       - Clears columns to the right dynamically:
         * Find any cell(s) containing "DDP" (before replacement) and compute the
           earliest column where clearing should start: (match_col + 3)
           e.g., if a "DDP" cell is in column G (7), start clearing at J (10).
         * If no "DDP" was present originally, default to L:XFD as before.
       - Restores shapes to their original coordinates
  5) Saves a copy with filename prefix changed from "yymmdd PS-" to "yymmdd RFP-",
     preserving extension and macros when applicable (.xlsm).
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
XL_OPENXML_WORKBOOK       = 51  # .xlsx
XL_OPENXML_WORKBOOK_MACRO = 52  # .xlsm

# Excel Find/Replace constants
XL_PART       = 2   # LookAt: xlPart (substring match)
XL_BYROWS     = 1   # SearchOrder: xlByRows
XL_NEXT       = 1   # SearchDirection: xlNext
XL_FORMULAS   = -4123  # xlFormulas (Find's Within arg if needed)

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
    name = src_path.name
    m = re.match(r"^(\d{6})\sPS-(.*)$", name, flags=re.IGNORECASE)
    if m:
        return src_path.with_name(f"{m.group(1)} RFP-{m.group(2)}")
    m2 = re.match(r"^(\d{6})(\s?)(.*)$", name)
    if m2 and "RFP-" not in name.upper():
        return src_path.with_name(f"{m2.group(1)} RFP-{m2.group(3)}")
    return src_path.with_name(f"{src_path.stem}_RFP{src_path.suffix}")

def delete_cover_if_present(xl_wb):
    for ws in xl_wb.Worksheets:
        if str(ws.Name).lower() == "cover":
            ws.Delete()
            return

def convert_all_used_ranges_to_values(xl_wb):
    """Freeze formulas to values (formats/images stay intact)."""
    for ws in xl_wb.Worksheets:
        try:
            used = ws.UsedRange
            vals = used.Value
            used.Value = vals
        except Exception:
            # Charts or special sheets may not support UsedRange.Value -> safe to skip
            pass

def get_ps_sheet(xl_wb):
    for ws in xl_wb.Worksheets:
        if str(ws.Name).lower() == "ps":
            return ws
    return None

def snapshot_and_move_shapes_to_safe_area(ps_ws):
    """
    Snapshot all shapes' geometry and move them temporarily into col A
    so clearing won't touch them even if anchored there.
    Returns a list of dicts with shape properties for restoration.
    """
    shapes_info = []
    safe_left = ps_ws.Range("A1").Left
    safe_top  = ps_ws.Range("A1").Top
    y_cursor  = safe_top

    try:
        count = ps_ws.Shapes.Count
    except Exception:
        count = 0

    for i in range(1, count + 1):
        shp = ps_ws.Shapes.Item(i)
        try:
            info = {
                "Name":    shp.Name,
                "Left":    shp.Left,
                "Top":     shp.Top,
                "Width":   shp.Width,
                "Height":  shp.Height,
                "Locked":  getattr(shp, "Locked", False),
            }
            shapes_info.append(info)

            # Temporarily unlock & move to safe area (column A), stacking vertically
            try:
                shp.Locked = False
            except Exception:
                pass
            shp.Left = safe_left
            shp.Top  = y_cursor
            y_cursor += shp.Height + 10  # add some spacing
        except Exception:
            # If a specific shape can't be moved, keep going
            continue

    return shapes_info

def restore_shapes(ps_ws, shapes_info):
    """Restore shape positions after clearing."""
    by_name = {s["Name"]: s for s in shapes_info}
    try:
        count = ps_ws.Shapes.Count
    except Exception:
        count = 0

    for i in range(1, count + 1):
        shp = ps_ws.Shapes.Item(i)
        name = shp.Name
        if name in by_name:
            meta = by_name[name]
            try:
                shp.Left  = meta["Left"]
                shp.Top   = meta["Top"]
                shp.Width = meta["Width"]
                shp.Height= meta["Height"]
                try:
                    shp.Locked = meta["Locked"]
                except Exception:
                    pass
            except Exception:
                # Best-effort restoration per shape
                continue

def find_earliest_clear_start_col_for_ps(ps_ws):
    """
    On PS sheet, search for any cell containing "DDP" (substring, case-insensitive).
    Return the earliest (leftmost) column index + 3 (i.e., start clearing from col+3).
    If no matches, return None.
    """
    try:
        used = ps_ws.UsedRange
    except Exception:
        return None

    # Use Range.Find to search within UsedRange
    try:
        first = used.Find(What="DDP", LookAt=XL_PART, SearchOrder=XL_BYROWS, SearchDirection=XL_NEXT, MatchCase=False)
    except Exception:
        first = None

    if not first:
        return None

    min_col_plus3 = first.Column + 3
    # Loop FindNext until we circle back
    cur = first
    while True:
        try:
            cur = used.FindNext(cur)
        except Exception:
            break
        if not cur:
            break
        if cur.Address == first.Address:
            break
        candidate = cur.Column + 3
        if candidate < min_col_plus3:
            min_col_plus3 = candidate

    return min_col_plus3

def clear_ps_columns_from_dynamic_start(ps_ws, start_col_index=None):
    """
    Clear from start_col_index to XFD (values + formats).
    If start_col_index is None, default to column L (12).
    """
    if start_col_index is None or start_col_index < 1:
        start_col_index = 12  # L

    # Excel's last column is XFD (16384)
    last_col = 16384
    if start_col_index > last_col:
        return

    # Build an A1-style column label for start_col_index
    col_label = column_index_to_label(start_col_index)
    rng_addr = f"{col_label}:XFD"
    ps_ws.Range(rng_addr).Clear()

def column_index_to_label(idx):
    """Convert 1-based column index to Excel column letters."""
    label = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        label = chr(65 + rem) + label
    return label

def replace_ddp_with_cif_all_sheets(xl_wb):
    """
    Replace substring "DDP" -> "CIF" on all worksheets, case-insensitive,
    without disturbing other text in the cell.
    """
    for ws in xl_wb.Worksheets:
        try:
            # Use Cells.Replace for partial substring replace
            ws.Cells.Replace(What="DDP", Replacement="CIF", LookAt=XL_PART, SearchOrder=XL_BYROWS, MatchCase=False)
        except Exception:
            # Some sheet types may not support Replace; ignore safely
            pass

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

        # 2) Convert all formulas to values
        convert_all_used_ranges_to_values(wb)

        # 3a) Determine dynamic clear start on PS using occurrences of "DDP" BEFORE we replace
        ps_ws = get_ps_sheet(wb)
        if ps_ws is None:
            print("[ERROR] 'PS' worksheet not found.")
            sys.exit(1)

        dynamic_start = find_earliest_clear_start_col_for_ps(ps_ws)  # may be None

        # 3b) Replace "DDP" -> "CIF" on ALL sheets (substring-safe)
        replace_ddp_with_cif_all_sheets(wb)

        # 4) PS sheet shape-preserving clear using dynamic start
        shapes_info = snapshot_and_move_shapes_to_safe_area(ps_ws)
        clear_ps_columns_from_dynamic_start(ps_ws, dynamic_start)
        restore_shapes(ps_ws, shapes_info)

        # 5) Save with new name / format
        wb.SaveAs(Filename=str(out), FileFormat=fileformat)
        print(f"[SUCCESS] Saved edited file to:\n{out}")

    except Exception as e:
        print(f"[ERROR] {e}")
        sys.exit(1)
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        try:
            excel.ScreenUpdating = True
            excel.DisplayAlerts  = True
            excel.Visible        = False
        except Exception:
            pass
        excel.Quit()

if __name__ == "__main__":
    main()
