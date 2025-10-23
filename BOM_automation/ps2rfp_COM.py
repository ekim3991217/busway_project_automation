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
  3) Replaces all occurrences of "DDP" with "FOB" across all worksheets
     (partial match safe: only the substring "DDP" is replaced).
     Then, for any cell containing BOTH "FOB" and "USA" (case-insensitive),
     the entire cell is overwritten with "FOB, BUSAN PORT".
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

  ### UPDATED 10/23/2025 (new behaviors):
  A) After DDP->FOB replacements, find the upper-most cell containing "FOB"
     in the entire workbook and CLEAR the cell directly below it.
  B) Read text from PS!G4, then find the NEXT occurrence of that text in the workbook.
  C) For that found cell, clear contents of the SAME COLUMN beginning 5 rows below
     (i.e., from row = found.Row + 6) down to the last used row.
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
XL_PART       = 2    # LookAt: xlPart (substring match)
XL_BYROWS     = 1    # SearchOrder: xlByRows
XL_NEXT       = 1    # SearchDirection: xlNext
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

def replace_ddp_fob_and_normalize_fob_usa_all_sheets(xl_wb):
    """
    Step 1: Replace substring "DDP" -> "FOB" on all worksheets (case-insensitive).
    Step 2: For any cell that contains BOTH "FOB" and "USA" (case-insensitive),
            overwrite the cell with exactly "FOB, BUSAN PORT".
    """
    # Step 1: Global DDP -> FOB
    for ws in xl_wb.Worksheets:
        try:
            ws.Cells.Replace(
                What="DDP",
                Replacement="FOB",
                LookAt=XL_PART,
                SearchOrder=XL_BYROWS,
                MatchCase=False
            )
        except Exception:
            pass

    # Step 2: Normalize any cell containing both "FOB" and "USA"
    for ws in xl_wb.Worksheets:
        try:
            used = ws.UsedRange
        except Exception:
            continue

        # Find all occurrences of "FOB" and then check for "USA" in the same cell
        try:
            first = used.Find(What="FOB", LookAt=XL_PART, SearchOrder=XL_BYROWS, SearchDirection=XL_NEXT, MatchCase=False)
        except Exception:
            first = None

        if not first:
            continue

        # Collect addresses to avoid infinite loop if we edit while iterating
        cells_to_fix = []
        cur = first
        while True:
            try:
                val = cur.Value
            except Exception:
                val = None

            if isinstance(val, (str, bytes)):
                s = str(val)
                if ("usa" in s.lower()) and ("fob" in s.lower()):
                    cells_to_fix.append(cur.Address)

            try:
                cur = used.FindNext(cur)
            except Exception:
                break
            if not cur or cur.Address == first.Address:
                break

        # Now set those cells to the normalized string
        for addr in cells_to_fix:
            try:
                ws.Range(addr).Value = "FOB, BUSAN PORT"
            except Exception:
                continue

# --- HELPER ADDED ---
def find_uppermost_match_across_workbook(xl_wb, needle, lookat=XL_PART):
    """### UPDATED: Find the upper-most (smallest row, then smallest column) cell containing 'needle' across all sheets."""
    needle_lc = str(needle).lower()
    best = None  # (ws, cell_range, row, col)
    for ws in xl_wb.Worksheets:
        try:
            used = ws.UsedRange
        except Exception:
            continue
        try:
            first = used.Find(What=needle, LookAt=lookat, SearchOrder=XL_BYROWS,
                              SearchDirection=XL_NEXT, MatchCase=False)
        except Exception:
            first = None
        if not first:
            continue
        # Walk through all hits on this sheet
        cur = first
        while True:
            r = cur.Row
            c = cur.Column
            if best is None or (r < best[2] or (r == best[2] and c < best[3])):
                best = (ws, cur, r, c)
            try:
                cur = used.FindNext(cur)
            except Exception:
                break
            if not cur or cur.Address == first.Address:
                break
    return best  # may be None

def find_next_occurrence_after_anchor(xl_wb, text, anchor_ws, anchor_row, anchor_col, lookat=XL_PART):
    """### UPDATED: Find the NEXT occurrence of 'text' after the given anchor (ws,row,col) in workbook order."""
    # First, try the same worksheet, after the anchor cell
    try:
        used = anchor_ws.UsedRange
    except Exception:
        used = None

    if used is not None:
        try:
            after_cell = anchor_ws.Cells(anchor_row, anchor_col)  # After param for Find
            hit = used.Find(What=text, After=after_cell, LookAt=lookat,
                            SearchOrder=XL_BYROWS, SearchDirection=XL_NEXT, MatchCase=False)
        except Exception:
            hit = None
        if hit and not (hit.Row == anchor_row and hit.Column == anchor_col):
            return anchor_ws, hit

    # If not found on same sheet, scan following worksheets then wrap around
    sheets = list(xl_wb.Worksheets)
    try:
        start_ix = sheets.index(anchor_ws)
    except Exception:
        start_ix = 0

    # Scan sheets after anchor first
    for ws in sheets[start_ix+1:] + sheets[:start_ix]:
        try:
            used2 = ws.UsedRange
        except Exception:
            continue
        try:
            first = used2.Find(What=text, LookAt=lookat, SearchOrder=XL_BYROWS,
                               SearchDirection=XL_NEXT, MatchCase=False)
        except Exception:
            first = None
        if not first:
            continue
        # If the first hit on this sheet is okay, return it
        return ws, first

    return None, None

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

        # 3b) Replace DDP->FOB, then normalize any "FOB" & "USA" cells
        replace_ddp_fob_and_normalize_fob_usa_all_sheets(wb)

        # A) Find UPPERMOST "FOB" and clear cell directly BELOW it
        # ----------------------------------------------------------
        best = find_uppermost_match_across_workbook(wb, "FOB")  # ### UPDATED: new call
        if best:
            ws_top, cell_top, r_top, c_top = best
            try:
                below = ws_top.Cells(r_top + 1, c_top)          # ### UPDATED
                below.Clear()                                    # ### UPDATED: clear contents+formats of the cell directly below
            except Exception:
                pass

        # B) Read PS!G4 text
        # ------------------
        try:
            text_g4 = ps_ws.Range("G4").Text                     # ### UPDATED: capture text in G4
            if text_g4 is None or str(text_g4).strip() == "":
                text_g4 = ps_ws.Range("G4").Value
        except Exception:
            text_g4 = None

        # C) Find NEXT occurrence of text_g4 after G4 across workbook
        # -----------------------------------------------------------
        found_ws, found_cell = (None, None)
        if text_g4 is not None and str(text_g4).strip() != "":
            found_ws, found_cell = find_next_occurrence_after_anchor(
                wb, str(text_g4), ps_ws, 4, 7, lookat=XL_PART  # G4 = row 4, col 7
            )                                                   # ### UPDATED

        # D) If found, clear contents of same column beyond 5 rows under that cell
        # ------------------------------------------------------------------------
        if found_ws is not None and found_cell is not None:
            try:
                start_row = found_cell.Row + 6                  # 5 rows below means start clearing from row+6
                col_idx   = found_cell.Column
                # Determine last used row in this worksheet
                last_row = found_ws.Cells(found_ws.Rows.Count, col_idx).End(-4162).Row  # xlUp = -4162
                if start_row <= last_row:
                    rng = found_ws.Range(found_ws.Cells(start_row, col_idx),
                                          found_ws.Cells(last_row, col_idx))
                    rng.ClearContents                           # ### UPDATED: clear contents (not formats) per requirement
            except Exception:
                pass
        # === END NEW BEHAVIORS ===

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
