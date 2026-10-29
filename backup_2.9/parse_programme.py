#!/usr/bin/env python3
"""
Tradifoot Logistics - Program Ingestion Parser, Validator & Upsert Engine
Demonstrates reading, sanitizing, schema-validating, and transforming
Excel delivery schedules into structured JSON ready for Supabase & Dashboard sync
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

DEFAULT_SUPABASE_URL = "https://gmdgnzokgqyivlfeosdy.supabase.co"
DEFAULT_SUPABASE_ANON_KEY = "sb_publishable_JPqzQ4b11U9BU8HzG16k2g_O2cHl5ur"
SUPABASE_DELIVERIES_ENDPOINT = f"{DEFAULT_SUPABASE_URL}/rest/v1/deliveries"


def build_delivery_key(date_str, store_str, quantite):
    """Deterministic Composite Fingerprint Matching (Date + Magasin + Quantite)."""
    d = str(date_str or "").strip()
    m = " ".join(str(store_str or "").strip().upper().split())
    try:
        q = int(quantite)
    except (ValueError, TypeError):
        q = 0
    if not d or not m or q <= 0:
        return ""
    return f"{d}|{m}|{q}"


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

        d_key = build_delivery_key(normalized_date, store_name, qty)
        if not normalized_date or not store_name or qty <= 0:
            errors.append(f"Row {row_idx}: Ignorée (données incomplètes: date='{normalized_date}', magasin='{store_name}', quantite={qty})")
            continue

        records.append({
            "id": delivery_id,
            "delivery_id": delivery_id,
            "delivery_key": d_key,
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


def push_manifest_to_supabase(excel_path_or_records, supabase_url=None, anon_key=None):
    """
    Ingests an Excel delivery manifest directly into the Supabase 'deliveries' table
    using HTTPS REST API (PostgREST upsert on conflict delivery_key).
    """
    base_url = (supabase_url or DEFAULT_SUPABASE_URL).rstrip("/")
    url = f"{base_url}/rest/v1/deliveries?on_conflict=delivery_key"
    key = anon_key or DEFAULT_SUPABASE_ANON_KEY

    if isinstance(excel_path_or_records, str):
        parsed = parse_delivery_excel(excel_path_or_records)
        records = parsed["records"]
    else:
        records = excel_path_or_records

    rows = []
    now_iso = datetime.now().isoformat()
    for r in records:
        del_date = str(r.get("date") or "").strip()
        del_store = str(r.get("magasin") or "").strip()
        try:
            del_qty = int(r.get("quantite") or 0)
        except (ValueError, TypeError):
            del_qty = 0

        # Filter corrupted / empty rows
        if not del_date or not del_store or del_qty <= 0:
            continue

        dkey = r.get("delivery_key") or build_delivery_key(del_date, del_store, del_qty)
        del_id = str(r.get("id") or r.get("delivery_id") or "").strip()

        rows.append({
            "delivery_id": del_id,
            "delivery_key": dkey,
            "date": del_date,
            "magasin": del_store,
            "chauffeur": str(r.get("chauffeur") or "").strip().upper(),
            "matricule": str(r.get("matricule") or "").strip().upper(),
            "quantite": del_qty,
            "status": r.get("status") or "Non Livré",
            "delivered_by": r.get("delivered_by"),
            "updated_at": now_iso
        })

    req_data = json.dumps(rows).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=req_data,
        headers={
            "apikey": key,
            "Authorization": f"Bearer {key}",
            "Content-Type": "application/json",
            "Prefer": "resolution=merge-duplicates,return=representation"
        },
        method="POST"
    )

    try:
        with urllib.request.urlopen(req, timeout=45) as resp:
            resp_body = resp.read().decode("utf-8")
            return {
                "status": "success",
                "count": len(rows),
                "data": json.loads(resp_body) if resp_body else []
            }
    except urllib.error.HTTPError as e:
        err_body = e.read().decode("utf-8", errors="ignore")
        # Schema tolerance fallback if remote table does not have on_conflict=delivery_key yet
        if "delivery_key" in err_body or "PGRST204" in err_body:
            fallback_url = f"{base_url}/rest/v1/deliveries"
            fallback_rows = [{
                "date": r["date"],
                "magasin": r["magasin"],
                "matricule": r["matricule"],
                "quantite": r["quantite"],
                "status": r["status"],
                "updated_at": r["updated_at"]
            } for r in rows]
            fallback_req = urllib.request.Request(
                fallback_url,
                data=json.dumps(fallback_rows).encode("utf-8"),
                headers={
                    "apikey": key,
                    "Authorization": f"Bearer {key}",
                    "Content-Type": "application/json"
                },
                method="POST"
            )
            try:
                with urllib.request.urlopen(fallback_req, timeout=45) as fb_resp:
                    return {"status": "success", "count": len(fallback_rows), "note": "inserted via fallback"}
            except Exception:
                pass
        return {"status": "error", "code": e.code, "message": str(e), "details": err_body}
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
