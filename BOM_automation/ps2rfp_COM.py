#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
ps_cleaner_preserve_images.py

Usage:
  python ps_cleaner_preserve_images.py "C:\path\to\250826 PS-USA-2444 EGLINTON PJT-A.xlsm"

If you omit the argument, you’ll be prompted for a path.

What it does:
  1) Deletes worksheet "COVER" (case-insensitive).
  2) Converts ALL cells on ALL worksheets to values (preserves formatting).
  3) On "PS" sheet:
       - Temporarily moves all shapes/pictures to a safe area in column A
       - Clears columns L through XFD (content + formats)
       - Restores shapes to their original coordinates
  4) Saves a copy with filename prefix changed from "yymmdd PS-" to "yymmdd RFP-",
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
XL_OPENXML_WORKBOOK        = 51  # .xlsx
XL_OPENXML_WORKBOOK_MACRO  = 52  # .xlsm


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
    so clearing L:XFD won't touch them even if anchored there.
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


def clear_ps_columns_from_L_right(ps_ws):
    """Clear L:XFD (values + formats) — *after* shapes are moved away."""
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

        # 2) Convert all formulas to values
        convert_all_used_ranges_to_values(wb)

        # 3) PS sheet shape-preserving clear
        ps_ws = get_ps_sheet(wb)
        if ps_ws is None:
            print("[ERROR] 'PS' worksheet not found.")
            sys.exit(1)

        # Move shapes safely to column A area, remember original geometry
        shapes_info = snapshot_and_move_shapes_to_safe_area(ps_ws)

        # Now clear L:XFD
        clear_ps_columns_from_L_right(ps_ws)

        # Restore shape positions
        restore_shapes(ps_ws, shapes_info)

        # 4) Save with new name / format
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
