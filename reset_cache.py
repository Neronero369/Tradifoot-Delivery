#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tradifoot Logistics - Local Cache & Program Archive Purge
Performs a hard reset of local cache files and archived Excel delivery programs
to start fresh from November 1, 2026.
"""

import os
import glob
import sys

WORKSPACE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(WORKSPACE_DIR, "data")
PROGRAMS_DIR = os.path.join(DATA_DIR, "programs")
DB_PATH = os.path.join(DATA_DIR, "delivery_db.json")


def reset_cache():
    print("=" * 65)
    print(" TRADIFOOT LOGISTICS - CACHE & ARCHIVE PURGE (NOVEMBER 2026 RESET)")
    print("=" * 65)

    # 1. Forcefully delete delivery_db.json
    print("\n[1/2] Checking local delivery database cache (delivery_db.json)...")
    if os.path.exists(DB_PATH):
        try:
            size_bytes = os.path.getsize(DB_PATH)
            os.remove(DB_PATH)
            print(f"  [+] Purged file: {DB_PATH}")
            print(f"      Size freed: {size_bytes:,} bytes")
        except Exception as e:
            print(f"  [-] Error deleting {DB_PATH}: {e}")
    else:
        print("  [i] No existing delivery_db.json found (already clean).")

    # 2. Delete all existing .xlsx / .xls files in data/programs/
    print("\n[2/2] Purging archived Excel delivery programs in data/programs/...")
    if not os.path.exists(PROGRAMS_DIR):
        os.makedirs(PROGRAMS_DIR, exist_ok=True)
        print("  [i] Created fresh data/programs/ directory.")
        deleted_count = 0
        deleted_bytes = 0
    else:
        # Match all .xlsx, .XLSX, .xls, .XLS files in data/programs/
        excel_patterns = ["*.xlsx", "*.XLSX", "*.xls", "*.XLS"]
        all_excels = set()
        for pat in excel_patterns:
            all_excels.update(glob.glob(os.path.join(PROGRAMS_DIR, pat)))

        deleted_count = 0
        deleted_bytes = 0
        for fpath in sorted(all_excels):
            fname = os.path.basename(fpath)
            try:
                fsize = os.path.getsize(fpath)
                os.remove(fpath)
                deleted_count += 1
                deleted_bytes += fsize
                print(f"  [-] Deleted program: {fname} ({fsize:,} bytes)")
            except Exception as e:
                print(f"  [!] Failed to delete {fname}: {e}")

        print(f"\n  [+] Successfully purged {deleted_count} Excel file(s) ({deleted_bytes:,} bytes) from data/programs/.")

    print("\n" + "=" * 65)
    print(" SUMMARY OF PURGE OPERATION")
    print(f" - Local DB ({DB_PATH}): {'DELETED' if not os.path.exists(DB_PATH) else 'STILL PRESENT'}")
    print(f" - Archived Excel Programs : {deleted_count} files deleted")
    print(" - Environment Status     : Clean and ready for fresh November 2026 sync")
    print("=" * 65 + "\n")


if __name__ == "__main__":
    reset_cache()
