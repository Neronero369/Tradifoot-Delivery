#!/usr/bin/env python3
"""
Tradifoot Logistics - Program Ingestion Parser & Validator
Demonstrates reading, sanitizing, schema-validating, and transforming
Excel delivery schedules into structured JSON ready for Google Apps Script & Dashboard sync.
"""

import sys
import json
from datetime import datetime, date
import openpyxl

AUTHORIZED_CHAUFFEURS = {
    "SAADI BILLEL",
    "SNOUCI SEIF-EDDINE",
    "BOUKHALFA ADEL",
    "BALLA YOUCEF",
    "BAHLOULI ADEL"
}

REQUIRED_HEADERS = [
    "ID", "Date", "Magasin", "Chauffeur", "Matricule", "Quantite", "Status", "Remarques"
]

def parse_delivery_excel(file_path):
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb.active

    # 1. Header Validation (Row 1)
    actual_headers = [str(cell.value or "").strip() for cell in ws[1]]
    # Match the first 8 columns
    actual_headers_prefix = actual_headers[:len(REQUIRED_HEADERS)]
    if actual_headers_prefix != REQUIRED_HEADERS:
        raise ValueError(
            f"Invalid Excel Header Schema.\nExpected: {REQUIRED_HEADERS}\nFound:    {actual_headers_prefix}"
        )

    records = []
    errors = []

    # 2. Iterate data rows starting at Row 2
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        # Skip completely empty rows
        if not any(row):
            continue

        raw_id, raw_date, raw_store, raw_driver, raw_mat, raw_qty, raw_status, raw_notes = row[:8]

        # Row ID
        delivery_id = str(raw_id).strip() if raw_id is not None else f"AUTO-{row_idx}"

        # Date normalization -> YYYY-MM-DD
        normalized_date = ""
        if isinstance(raw_date, (datetime, date)):
            normalized_date = raw_date.strftime("%Y-%m-%d")
        elif isinstance(raw_date, str) and raw_date.strip():
            raw_str = raw_date.strip()
            # Try parsing various common formats (ISO, French DD/MM/YYYY, etc.)
            parsed = False
            for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
                try:
                    dt = datetime.strptime(raw_str, fmt)
                    normalized_date = dt.strftime("%Y-%m-%d")
                    parsed = True
                    break
                except ValueError:
                    pass
            if not parsed:
                errors.append(f"Row {row_idx} ({raw_store}): Format de date invalide: '{raw_date}'")
                normalized_date = raw_str
        else:
            errors.append(f"Row {row_idx}: Date manquante")

        # Magasin / Destination
        store_name = str(raw_store or "").strip().upper()
        if not store_name:
            errors.append(f"Row {row_idx}: Nom de magasin manquant")

        # Chauffeur validation & normalization
        driver_name = str(raw_driver or "").strip().upper()
        if driver_name and driver_name not in AUTHORIZED_CHAUFFEURS:
            # Check for close match or soft warning
            errors.append(f"Row {row_idx}: Chauffeur non autorisé '{raw_driver}' (Magasin: {store_name})")

        # Matricule
        matricule = str(raw_mat or "").strip().upper()

        # Quantite
        try:
            qty = int(raw_qty) if raw_qty is not None else 0
            if qty < 1:
                errors.append(f"Row {row_idx}: Quantité ({qty}) doit être >= 1")
        except (ValueError, TypeError):
            errors.append(f"Row {row_idx}: Quantité non numérique '{raw_qty}'")
            qty = 0

        # Status normalization
        raw_status_str = str(raw_status or "Non Livré").strip().lower()
        if "livr" in raw_status_str and "non" not in raw_status_str:
            status = "Livré"
        else:
            status = "Non Livré"

        # Remarques
        notes = str(raw_notes or "").strip()

        records.append({
            "id": delivery_id,
            "date": normalized_date,
            "magasin": store_name,
            "chauffeur": driver_name,
            "matricule": matricule,
            "quantite": qty,
            "status": status,
            "remarques": notes
        })

    return {
        "success": len(errors) == 0,
        "total_records": len(records),
        "records": records,
        "errors": errors
    }

if __name__ == "__main__":
    test_file = "template_programme_de_livraison.xlsx"
    print(f"Parsing test file: {test_file}")
    res = parse_delivery_excel(test_file)
    print(f"Success: {res['success']}")
    print(f"Total Rows: {res['total_records']}")
    print(f"Errors ({len(res['errors'])}): {res['errors']}")
    print("Parsed Sample Records:")
    print(json.dumps(res['records'], indent=2, ensure_ascii=False))
