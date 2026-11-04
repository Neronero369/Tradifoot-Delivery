#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tradifoot Logistics - Production Environment Hard Reset
Purges all cached data files and archived delivery program spreadsheets
to prepare the local system for clean production ingestion.
"""

import os
from pathlib import Path


def production_reset():
    base_dir = Path(__file__).resolve().parent
    data_dir = base_dir / "data"
    programs_dir = data_dir / "programs"
    db_file = data_dir / "delivery_db.json"

    print("=" * 65)
    print(" TRADIFOOT LOGISTICS - PRODUCTION SYSTEM RESET")
    print("=" * 65)

    # 1. Forcefully delete data/delivery_db.json if it exists
    print("\n[1/2] Checking local delivery database cache (delivery_db.json)...")
    if db_file.exists():
        try:
            file_size = db_file.stat().st_size
            db_file.unlink()
            print(f"  [+] Deleted: {db_file.name} ({file_size:,} bytes)")
        except Exception as e:
            print(f"  [-] Error deleting {db_file}: {e}")
    else:
        print("  [i] delivery_db.json does not exist (already clean).")

    # 2. Iterate through data/programs/ and delete all .xlsx files
    print("\n[2/2] Scanning data/programs/ for Excel delivery files...")
    deleted_excel_count = 0
    deleted_excel_bytes = 0

    if programs_dir.exists():
        excel_files = [
            f for f in programs_dir.iterdir()
            if f.is_file() and f.suffix.lower() == ".xlsx"
        ]

        if excel_files:
            for f in sorted(excel_files):
                try:
                    fsize = f.stat().st_size
                    f.unlink()
                    deleted_excel_count += 1
                    deleted_excel_bytes += fsize
                    print(f"  [+] Deleted program: {f.name} ({fsize:,} bytes)")
                except Exception as e:
                    print(f"  [-] Failed to delete {f.name}: {e}")
            print(f"\n  [+] Total purged: {deleted_excel_count} file(s) ({deleted_excel_bytes:,} bytes).")
        else:
            print("  [i] No .xlsx files found in data/programs/ (already clean).")
    else:
        programs_dir.mkdir(parents=True, exist_ok=True)
        print("  [+] Created empty data/programs/ directory.")

    print("\n" + "=" * 65)
    print(" SUMMARY OF PRODUCTION RESET")
    print(f" - Local DB ({db_file.name}) : {'PURGED' if not db_file.exists() else 'STILL EXISTS'}")
    print(f" - Programs Purged        : {deleted_excel_count} .xlsx file(s)")
    print(" - Production Readiness   : System ready for production ingestion.")
    print("=" * 65 + "\n")


if __name__ == "__main__":
    production_reset()
