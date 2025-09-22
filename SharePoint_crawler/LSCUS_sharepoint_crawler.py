import os
import csv
import json
from datetime import datetime, timedelta

# Hardcoded base path to your synced SharePoint library
BASE_PATH = r"C:\Users\EKim\LS Cable\LSCUS Busway Division - General\06. LS Busway Library"
OUT_DIR = "./LSUCS_Busway_Library_inventory"

# Define 1 yr in dates
ONE_YEAR = timedelta(days=365) 

def iso(ts: float) -> str:
    return datetime.fromtimestamp(ts).isoformat()

def crawl_local(base_path):
    rows = []
    for root, dirs, files in os.walk(base_path):
        rel_root = os.path.relpath(root, base_path)
        if rel_root == ".":
            rel_root = ""

        # Folders
        for d in dirs:
            full_path = os.path.join(root, d)
            try:
                stat = os.stat(full_path)
            except (FileNotFoundError, PermissionError):
                continue  # skip transient/unavailable entries
            rows.append({
                "name": d,
                "is_folder": True,
                "parent_path": rel_root,
                "full_path": os.path.join(rel_root, d),
                "absolute_path": full_path,
                "size_bytes": "",
                "created": iso(stat.st_ctime),
                "modified": iso(stat.st_mtime),
                "modified_epoch": stat.st_mtime,
                "created_epoch": stat.st_ctime,
                "days_since_modified": None,
            })

        # Files
        for f in files:
            full_path = os.path.join(root, f)
            try:
                stat = os.stat(full_path)
            except (FileNotFoundError, PermissionError):
                continue
            days_since_mod = (datetime.now() - datetime.fromtimestamp(stat.st_mtime)).days
            rows.append({
                "name": f,
                "is_folder": False,
                "parent_path": rel_root,
                "full_path": os.path.join(rel_root, f),
                "absolute_path": full_path,
                "size_bytes": stat.st_size,
                "created": iso(stat.st_ctime),
                "modified": iso(stat.st_mtime),
                "modified_epoch": stat.st_mtime,
                "created_epoch": stat.st_ctime,
                "days_since_modified": days_since_mod,
            })
    return rows

def save_outputs(rows):
    os.makedirs(OUT_DIR, exist_ok=True)

    # Sort for stable outputs (folders first, then files; alphabetical by path)
    rows_sorted = sorted(
        rows,
        key=lambda r: (r["full_path"].lower(), not r["is_folder"])
    )

    # CSV inventory
    csv_path = os.path.join(OUT_DIR, "inventory.csv")
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=rows_sorted[0].keys())
        writer.writeheader()
        writer.writerows(rows_sorted)

    # JSON inventory
    json_path = os.path.join(OUT_DIR, "inventory.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(rows_sorted, f, ensure_ascii=False, indent=2)

    # Tree.txt (folders and files in a simple indented listing)
    tree_path = os.path.join(OUT_DIR, "tree.txt")
    with open(tree_path, "w", encoding="utf-8") as f:
        for r in rows_sorted:
            # Build indent by folder depth
            depth = r["full_path"].count(os.sep)
            indent = "    " * depth
            if r["is_folder"]:
                f.write(f"{indent}📁 {os.path.basename(r['full_path']) or r['full_path']}\n")
            else:
                f.write(f"{indent}└── {r['name']}\n")

    print(f"[OK] Wrote:\n  {csv_path}\n  {json_path}\n  {tree_path}")

def save_older_than_1yr(rows):
    cutoff_dt = datetime.now() - ONE_YEAR
    cutoff_epoch = cutoff_dt.timestamp()

    # Filter only files (not folders) where modified < cutoff
    aged_files = [
        r for r in rows
        if (not r["is_folder"]) and (r["modified_epoch"] < cutoff_epoch)
    ]

    if not aged_files:
        print("[INFO] No files older than 1 year were found.")
        # Still emit an empty file with header for consistency
        header = [
            "name","is_folder","parent_path","full_path","absolute_path",
            "size_bytes","created","modified","modified_epoch","created_epoch","days_since_modified"
        ]
        out_path = os.path.join(OUT_DIR, "files_older_than_1yr.csv")
        with open(out_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=header)
            writer.writeheader()
        print(f"[OK] Wrote empty report: {out_path}")
        return

    # Sort aged files by oldest first
    aged_files.sort(key=lambda r: r["modified_epoch"])

    out_path = os.path.join(OUT_DIR, "files_older_than_1yr.csv")
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=aged_files[0].keys())
        writer.writeheader()
        writer.writerows(aged_files)

    print(f"[OK] Wrote aged-file report:\n  {out_path} (count: {len(aged_files)})")

if __name__ == "__main__":
    rows = crawl_local(BASE_PATH)
    save_outputs(rows)
    save_older_than_1yr(rows)
    print(f"[DONE] Scanned {len(rows)} items under:\n  {BASE_PATH}")
