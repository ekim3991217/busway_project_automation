#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# """
# Script: create_PJT_folder.py
# Usage: python create_PJT_folder.py
# Description:
#   1) Prompts for PJT TITLE and project type (C/I/DC).
#   2) Creates folder: "ddmmyy_PJT TITLE_C/I/DC" (today's date).
#   3) Creates subfolders: "from KOR" and "to KOR" inside it.
#   4) Root path is fixed to:
#      C:\Users\EKim\OneDrive - LS Cable\PM - EugeneKim\2_QUOTATION_PO&RFQ
# """

from pathlib import Path
from datetime import datetime
import re
import sys

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
        pjt_type = prompt_project_type()

        # 2) Build folder name
        today = datetime.now().strftime("%M%D%y")
        folder_name = f"{today}_{pjt_title}_{pjt_type}"

        # 3) Build full path under fixed base directory
        project_dir = BASE_PATH / folder_name

        # Create directories
        project_dir.mkdir(parents=True, exist_ok=True)
        (project_dir / "from KOR").mkdir(exist_ok=True)
        (project_dir / "to KOR").mkdir(exist_ok=True)

        print("\n[OK] Created project structure:")
        print(f"  {project_dir}")
        print(f"  {project_dir / 'from KOR'}")
        print(f"  {project_dir / 'to KOR'}")

    except KeyboardInterrupt:
        print("\n[Cancelled]")
        sys.exit(1)
    except Exception as e:
        print(f"[ERROR] {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
