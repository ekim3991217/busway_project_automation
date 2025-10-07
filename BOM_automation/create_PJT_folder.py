#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from pathlib import Path
from datetime import datetime
import re
import sys
import shutil

# Windows-invalid filename chars
INVALID_CHARS_PATTERN = r'[\\/:*?"<>|]+'

# Fixed root path
BASE_PATH = Path(r"C:\Users\EKim\OneDrive - LS Cable\PM - EugeneKim\2_QUOTATION_PO&RFQ")
TEMPLATES_PATH = BASE_PATH / "_TEMPLATES"

VALID_TYPES = {"EX", "DATA", "NSPB"}

# NSPB intentionally uses the same template file as EX (EXWAY)
TEMPLATE_MAP = {
    "EX":   "yymmdd PS-USA-EXWAY-A_TEMPLATE.xlsm",
    "DATA": "yymmdd PS-USA-DATAWAY-A_TEMPLATE.xlsm",
    "NSPB": "yymmdd PS-USA-EXWAY-A_TEMPLATE.xlsm",  # reuse EXWAY template for NSPB
}

def prompt_nonempty(prompt: str) -> str:
    while True:
        val = input(prompt).strip()
        if val:
            return val
        print("  [!] Please enter a non-empty value.")

def prompt_product_type() -> str:
    while True:
        val = input("choose product type (EX, DATA, NSPB): ").strip().upper()
        if val in VALID_TYPES:
            return val
        print("  [!] Invalid input. Must be one of: EX, DATA, NSPB.")

def sanitize_title(title: str) -> str:
    title = re.sub(INVALID_CHARS_PATTERN, " ", title)
    title = re.sub(r"\s+", " ", title).strip()
    return title

def main():
    try:
        # 1) Inputs
        pjt_title_raw = prompt_nonempty("ENTER PJT TITLE: ")
        # Sanitize then force ALL CAPS for consistent folder/file naming
        pjt_title = sanitize_title(pjt_title_raw).upper()
        pjt_type = prompt_product_type()

        # Dates for folder vs file naming
        folder_date = datetime.now().strftime("%m%d%y")   # mmddyy for folder
        file_date   = datetime.now().strftime("%y%m%d")   # yymmdd for file

        # 2) Build folder name
        folder_name = f"{folder_date}_{pjt_title}_{pjt_type}"

        # 3) Build full path under fixed base directory
        project_dir = BASE_PATH / folder_name

        # Create directories
        project_dir.mkdir(parents=True, exist_ok=True)
        (project_dir / "from KOR").mkdir(exist_ok=True)
        (project_dir / "to KOR").mkdir(exist_ok=True)
        (project_dir / "FINAL").mkdir(exist_ok=True)

        # 4) Determine template source based on product type
        template_name = TEMPLATE_MAP[pjt_type]
        src_template = TEMPLATES_PATH / template_name
        if not src_template.exists():
            print(f"[ERROR] Template file not found:\n  {src_template}\n"
                  f"Please verify the template exists and is dated '{file_date}'.")
            sys.exit(1)

        # 5) Copy and rename inside the project folder
        dest_filename = f"{file_date} PS-USA-{pjt_title} PJT-A.xlsm"
        dest_path = project_dir / dest_filename

        shutil.copy2(src_template, dest_path)

        print("\n[OK] Created project structure and template file:")
        print(f"  {project_dir}")
        print(f"  {project_dir / 'from KOR'}")
        print(f"  {project_dir / 'to KOR'}")
        print(f"  {project_dir / 'FINAL'}")
        print(f"  {dest_path}")

    except KeyboardInterrupt:
        print("\n[Cancelled]")
        sys.exit(1)
    except Exception as e:
        print(f"[ERROR] {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
