#!/usr/bin/env python3
"""
Tradifoot Logistics - Program Ingestion Parser, Validator & Upsert Engine
Demonstrates reading, sanitizing, schema-validating, and transforming
Excel delivery schedules into structured JSON ready for Google Apps Script & Dashboard sync
with strict Status Preservation (Never Downgrade 'Livré').
"""

import sys
import os
import json
import urllib.request
import urllib.error
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

DEFAULT_GOOGLE_APP_URL = "https://script.google.com/macros/s/AKfycbwqLs9WYjdjESCnX9avFBABv2bZdTXh1a43rYDC6N51l4uIO6em0wS51OaTyTuMAOY/exec"


def normalize_composite_key(date_str, store_str, driver_str):
    """Generates a canonical composite key [Date + Magasin + Chauffeur]."""
    d = str(date_str or "").strip()
    m = " ".join(str(store_str or "").strip().upper().split())
    c = " ".join(str(driver_str or "").strip().upper().split())
    return f"{d}|||{m}|||{c}"


def parse_delivery_excel(file_path):
    """
    Parses and schema-validates a Programme de Livraison Excel workbook.
    """
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb.active

    # 1. Header Validation (Row 1)
    actual_headers = [str(cell.value or "").strip() for cell in ws[1]]
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


def upsert_manifest_records(incoming_records, existing_records):
    """
    Performs in-memory Upsert logic with strict Status Preservation.

    CRITICAL STATUS PRESERVATION CONTRACT:
    - If a delivery already exists and its current status is 'Livré', PRESERVE 'Livré'.
      Never downgrade to 'Non Livré' or empty values.
    - If a delivery exists with status 'Non Livré', update mutable fields (Quantité, Matricule)
      but keep status as 'Non Livré' (unless incoming explicitly provides 'Livré').
    - If a delivery is brand new, insert as 'Non Livré'.
    """
    id_map = {}
    comp_map = {}

    for rec in existing_records:
        r_id = str(rec.get("id") or "").strip().upper()
        if r_id:
            id_map[r_id] = rec

        ckey = normalize_composite_key(rec.get("date"), rec.get("magasin"), rec.get("chauffeur"))
        if ckey:
            comp_map[ckey] = rec

    stats = {
        "total": len(incoming_records),
        "preserved_livre": 0,
        "upgraded_livre": 0,
        "updated_non_livre": 0,
        "inserted_new": 0
    }

    merged = list(existing_records)

    for item in incoming_records:
        in_id = str(item.get("id") or "").strip().upper()
        ckey = normalize_composite_key(item.get("date"), item.get("magasin"), item.get("chauffeur"))

        matched = id_map.get(in_id) or comp_map.get(ckey)

        if matched:
            curr_status = matched.get("status", "Non Livré")
            incoming_status = item.get("status", "Non Livré")

            if curr_status == "Livré":
                # PRESERVE 'Livré'
                final_status = "Livré"
                stats["preserved_livre"] += 1
            elif incoming_status == "Livré":
                final_status = "Livré"
                stats["upgraded_livre"] += 1
            else:
                final_status = "Non Livré"
                stats["updated_non_livre"] += 1

            matched["status"] = final_status
            matched["completed"] = (final_status == "Livré")
            if item.get("quantite"):
                matched["quantite"] = item["quantite"]
            if item.get("matricule"):
                matched["matricule"] = item["matricule"]
        else:
            # Brand new row
            new_item = dict(item)
            if new_item.get("status") != "Livré":
                new_item["status"] = "Non Livré"
            new_item["completed"] = (new_item["status"] == "Livré")
            merged.append(new_item)
            stats["inserted_new"] += 1

            if in_id:
                id_map[in_id] = new_item
            if ckey:
                comp_map[ckey] = new_item

    return merged, stats


def push_manifest_to_google_sheets(excel_path_or_records, web_app_url=None):
    """
    Ingests an Excel delivery manifest directly into Google Sheets
    via the Google Apps Script Web App 'batch_upsert' action.
    """
    url = web_app_url or DEFAULT_GOOGLE_APP_URL

    if isinstance(excel_path_or_records, str):
        parsed = parse_delivery_excel(excel_path_or_records)
        records = parsed["records"]
    else:
        records = excel_path_or_records

    payload = {
        "action": "batch_upsert",
        "deliveries": records
    }

    req_data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=req_data,
        headers={"Content-Type": "application/json"}
    )

    try:
        with urllib.request.urlopen(req, timeout=45) as resp:
            resp_body = resp.read().decode("utf-8")
            return json.loads(resp_body)
    except urllib.error.URLError as e:
        return {"status": "error", "message": str(e)}


if __name__ == "__main__":
    test_file = "template_programme_de_livraison.xlsx"
    print(f"Parsing test file: {test_file}")
    res = parse_delivery_excel(test_file)
    print(f"Success: {res['success']}")
    print(f"Total Rows: {res['total_records']}")
    print(f"Errors ({len(res['errors'])}): {res['errors']}")
    print(f"Sample Parsed Record: {res['records'][0] if res['records'] else None}")
