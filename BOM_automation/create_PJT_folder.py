#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# """
# Script: create_PJT_folder.py
# Usage: python create_PJT_folder.py
# Description:
#   1) Prompts for PJT TITLE and project type (C/I/DC).
#   2) Creates folder: "ddmmyy_PJT TITLE_C/I/DC" (today's date; mmddyy).
#   3) Creates subfolders: "from KOR" and "to KOR" inside it.
#   4) Copies the correct Excel template into the project folder and renames it:
#        - If C or I -> copy "...\\yymmdd PS-USA-EXWAY-A_TEMPLATE.xlsm"
#        - If DC     -> copy "...\\yymmdd PS-USA-DATAWAY-A_TEMPLATE.xlsm"
#      New file name: "yymmdd PS-USA-<PJT TITLE>-A.xlsm"
#   5) Root path is fixed to:
#      C:\\Users\\EKim\\OneDrive - LS Cable\\PM - EugeneKim\\2_QUOTATION_PO&RFQ
# """

from pathlib import Path
from datetime import datetime
import re
import sys
import shutil

# Windows-invalid filename chars
INVALID_CHARS_PATTERN = r'[\\/:*?"<>|]+'

# Fixed root path
BASE_PATH = Path(r"C:\Users\EKim\OneDrive - LS Cable\PM - EugeneKim\2_QUOTATION_PO&RFQ")

def prompt_nonempty(prompt: str) -> str:
    while True:
        val = input(prompt).strip()
        if val:
            return val
        print("  [!] Please enter a non-empty value.")

def prompt_project_type() -> str:
    while True:
        val = input("COMMERCIAL/INDUSTRIAL/DATA CENTER? (C, I, DC): ").strip().lower()
        if val in {"c", "i", "dc"}:
            return val.upper()
        print("  [!] Invalid input. Type C, I, or DC.")

def sanitize_title(title: str) -> str:
    title = re.sub(INVALID_CHARS_PATTERN, " ", title)
    title = re.sub(r"\s+", " ", title).strip()
    return title

def main():
    try:
        # 1) Inputs
        pjt_title_raw = prompt_nonempty("ENTER PJT TITLE: ")
        pjt_title = sanitize_title(pjt_title_raw)
        pjt_type = prompt_project_type()  # guarantees C / I / DC

        # Dates for folder vs file naming
        folder_date = datetime.now().strftime("%m%d%y")   # mmddyy for folder (kept as in original)
        file_date   = datetime.now().strftime("%y%m%d")   # yymmdd for file (per your instruction)

        # 2) Build folder name
        folder_name = f"{folder_date}_{pjt_title}_{pjt_type}"

        # 3) Build full path under fixed base directory
        project_dir = BASE_PATH / folder_name

        # Create directories
        project_dir.mkdir(parents=True, exist_ok=True)
        (project_dir / "from KOR").mkdir(exist_ok=True)
        (project_dir / "to KOR").mkdir(exist_ok=True)

        # 4) Determine template source based on project type
        #    Source files are expected to be in BASE_PATH and prefixed with today's yymmdd.
        if pjt_type in {"C", "I"}:
            template_name = f"{file_date} PS-USA-EXWAY-A_TEMPLATE.xlsm"
        elif pjt_type == "DC":
            template_name = f"{file_date} PS-USA-DATAWAY-A_TEMPLATE.xlsm"
        else:
            # Should never hit because prompt enforces, but keep as safety.
            print("  [!] Invalid project type. Please enter C, I, or DC.")
            sys.exit(1)

        src_template = BASE_PATH / template_name
        if not src_template.exists():
            print(f"[ERROR] Template file not found:\n  {src_template}\n"
                  f"Please verify the template exists and is dated '{file_date}'.")
            sys.exit(1)

        # 5) Copy and rename inside the project folder
        #    New filename: "yymmdd PS-USA-<PJT TITLE>-A.xlsm"
        dest_filename = f"{file_date} PS-USA-{pjt_title} PJT-A.xlsm"
        dest_path = project_dir / dest_filename

        shutil.copy2(src_template, dest_path)

        print("\n[OK] Created project structure and template file:")
        print(f"  {project_dir}")
        print(f"  {project_dir / 'from KOR'}")
        print(f"  {project_dir / 'to KOR'}")
        print(f"  {dest_path}")

    except KeyboardInterrupt:
        print("\n[Cancelled]")
        sys.exit(1)
    except Exception as e:
        print(f"[ERROR] {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
