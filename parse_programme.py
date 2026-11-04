#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tradifoot Logistics - Program Ingestion Parser, Validator & Upsert Engine
Production Excel Delivery Schedule Parser & Supabase Synchronization.
Dynamically locates headers (Boutique, Globale, Chauffeur, Livraison)
and maps records directly to the Supabase database schema.
Optimized for minimal RAM usage (openpyxl read-only streaming) and combined API batching.
"""

import os
import re
import json
import urllib.request
import urllib.error
from datetime import datetime, date
import openpyxl
import pandas as pd
import numpy as np


def _load_env():
    """Lightweight zero-dependency .env loader."""
    env_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if os.path.exists(env_file):
        try:
            with open(env_file, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith("#") and "=" in line:
                        k, v = line.split("=", 1)
                        k, v = k.strip(), v.strip().strip("'\"")
                        if k and k not in os.environ:
                            os.environ[k] = v
        except Exception:
            pass


_load_env()

DEFAULT_SUPABASE_REST_URL = os.environ.get(
    "SUPABASE_REST_URL",
    "https://gmdgnzokgqyivlfeosdy.supabase.co/rest/v1/deliveries"
)
DEFAULT_SUPABASE_API_KEY = os.environ.get(
    "SUPABASE_ANON_KEY",
    os.environ.get("SUPABASE_KEY", "sb_publishable_JPqzQ4b11U9BU8HzG16k2g_O2cHl5ur")
)

AUTHORIZED_CHAUFFEURS = {
    "SAADI BILLEL",
    "SNOUCI SEIF-EDDINE",
    "BOUKHALFA ADEL",
    "BALLA YOUCEF"
}

BLACKLISTED_CHAUFFEURS = {
    "BAHLOULI ADEL",
    "BAHLOULI"
}


def build_delivery_key(date_str, store_str, quantite, occurrence=None):
    """Deterministic Composite Key per row."""
    d = str(date_str or "").strip()
    m = " ".join(str(store_str or "").strip().upper().split())
    try:
        q = int(quantite)
    except (ValueError, TypeError):
        q = 0
    if not d or not m:
        return ""
    suffix = f"|OCC{occurrence}" if occurrence is not None else ""
    return f"{d}|{m}|{q}{suffix}"


def normalize_date_to_iso(val, fallback=""):
    """
    Converts date representations (datetime, date, DD/MM/YYYY, etc.) to YYYY-MM-DD.
    """
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d")
    s = str(val or "").strip()
    if not s:
        return fallback or datetime.today().strftime("%Y-%m-%d")

    # Priority 1: Standard date formats
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y", "%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(s[:10], fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass

    # Priority 2: ISO formatted timestamps
    try:
        return datetime.fromisoformat(s[:19]).strftime("%Y-%m-%d")
    except Exception:
        pass

    return s[:10]


def fetch_existing_delivery_ids(supabase_url=None, api_key=None):
    """
    Fetches all existing delivery_id and delivery_key values from the Supabase
    'deliveries' table and returns them as a Python set for fast lookup.
    """
    url = (supabase_url or DEFAULT_SUPABASE_REST_URL).split("?")[0]
    key = api_key or DEFAULT_SUPABASE_API_KEY
    endpoint = f"{url}?select=delivery_id,delivery_key"
    headers = {
        "apikey": key,
        "Authorization": f"Bearer {key}"
    }
    existing_ids = set()
    try:
        req = urllib.request.Request(endpoint, headers=headers)
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            for item in data:
                did = str(item.get("delivery_id") or "").strip().upper()
                dkey = str(item.get("delivery_key") or "").strip().upper()
                if did:
                    existing_ids.add(did)
                if dkey:
                    existing_ids.add(dkey)
    except Exception as e:
        print(f"[!] Warning: Could not fetch existing Supabase IDs ({e})")
    return existing_ids


def parse_delivery_excel(file_path, skip_existing=True, existing_ids=None):
    """
    Parses a production Programme de Livraison Excel workbook dynamically by header name.
    Uses pandas to load the DataFrame, strip header whitespaces, and safely extract columns.

    INSERT-ONLY SAFEGUARD:
    If skip_existing is True (default), fetches existing Supabase IDs before row iteration
    and skips any row whose generated delivery_id is already present, protecting live
    driver data (status, photo_url, recouvrement_url, GPS).

    CRITICAL COLUMN MAPPING:
      - Boutique   -> maps to magasin (String, uppercase, trimmed)
      - Globale    -> maps to quantite (Integer, default 0, skip row if 0/empty)
      - Chauffeur  -> maps to chauffeur (String, uppercase, strictly "" if blank or missing)
      - Livraison  -> maps to date (Convert from DD/MM/YYYY to YYYY-MM-DD)
      - Materiel   -> maps to materiel (String, stripped, None if blank/empty/NaN)
    """
    # 1. Fetch Existing State from Supabase for Insert-Only safe sync
    if skip_existing and existing_ids is None:
        existing_ids = fetch_existing_delivery_ids()
    # 1. Locate the delivery program sheet (or default to first sheet)
    xl = pd.ExcelFile(file_path)
    try:
        prog_sheet_name = None
        for s_name in xl.sheet_names:
            if re.search(r'prog|poro|livr', str(s_name), re.I):
                prog_sheet_name = s_name
                break
        if not prog_sheet_name:
            prog_sheet_name = xl.sheet_names[0]

        # 2. Dynamic Header Resolution (search rows 0 to 8)
        preview_df = pd.read_excel(xl, sheet_name=prog_sheet_name, nrows=8, header=None)
        header_row_idx = 0
        raw_headers = []

        for r_idx, row_vals in preview_df.iterrows():
            clean_row = [str(c or "").strip().lower() for c in row_vals if pd.notna(c)]
            has_boutique = any(h in ("boutique", "magasin") or "boutique" in h or "magasin" in h for h in clean_row)
            has_globale = any(h in ("globale", "global", "quantite", "quantité", "qte") or "globale" in h or "quantite" in h for h in clean_row)

            if has_boutique and (has_globale or any("livraison" in h or "chauffeur" in h for h in clean_row)):
                header_row_idx = r_idx
                raw_headers = [str(c or "").strip() for c in row_vals if pd.notna(c)]
                break

        # Load dataframe starting at detected header row
        df = pd.read_excel(xl, sheet_name=prog_sheet_name, header=header_row_idx)

        # 1. Clean the Excel Headers (Pandas)
        df.columns = df.columns.astype(str).str.strip()

        # Identify primary columns safely from stripped column names
        col_boutique = None
        col_globale = None
        col_chauffeur = None
        col_livraison = None
        col_remarque = None
        col_materiel = None

        for col in df.columns:
            c_clean = str(col).strip()
            c_lower = c_clean.lower()
            if ("boutique" in c_lower or "magasin" in c_lower) and col_boutique is None:
                col_boutique = col
            elif ("globale" in c_lower or "global" in c_lower or "quantite" in c_lower or "quantité" in c_lower or "qte" in c_lower) and col_globale is None:
                col_globale = col
            elif ("chauffeur" in c_lower or "driver" in c_lower) and col_chauffeur is None:
                col_chauffeur = col
            elif ("livraison" in c_lower or "date" in c_lower) and col_livraison is None:
                col_livraison = col
            elif ("remarque" in c_lower or "note" in c_lower or "observation" in c_lower) and col_remarque is None:
                col_remarque = col
            elif ("materiel" in c_lower or "matériel" in c_lower or "vehicule" in c_lower or "véhicule" in c_lower or "camion" in c_lower or "engin" in c_lower) and col_materiel is None:
                col_materiel = col

        if col_boutique is None:
            for cand in ['Boutique', 'Magasin', 'BOUTIQUE', 'MAGASIN', 'Destinataire']:
                if cand in df.columns:
                    col_boutique = cand
                    break

        if col_boutique is None:
            raise ValueError(
                f"Could not locate required 'Boutique' column in {file_path}. Headers detected: {list(df.columns)}"
            )

        records = []
        errors = []
        occurrence_tracker = {}
        skipped_existing_count = 0

        # 2. Iterate data rows
        for row_idx, row in df.iterrows():
            # Column 1: Boutique -> magasin (String, uppercase, trimmed)
            raw_store = ""
            if col_boutique is not None and col_boutique in df.columns:
                val_store = row.get(col_boutique)
                if val_store is not None and pd.notna(val_store):
                    raw_store = str(val_store).strip()
            if not raw_store:
                for k in ['Boutique', 'Magasin', 'BOUTIQUE', 'MAGASIN', 'Destinataire']:
                    if k in df.columns and pd.notna(row.get(k)):
                        raw_store = str(row.get(k)).strip()
                        break

            magasin = " ".join(raw_store.upper().split())

            # Column 2: Globale -> quantite (Integer, default 0 if blank)
            raw_qty = 0
            if col_globale is not None and col_globale in df.columns:
                val_qty = row.get(col_globale)
                if val_qty is not None and pd.notna(val_qty):
                    raw_qty = val_qty
            if not raw_qty:
                for k in ['Globale', 'Global', 'Quantité', 'Quantite', 'QTE', 'Qte', 'Total']:
                    if k in df.columns and pd.notna(row.get(k)):
                        raw_qty = row.get(k)
                        break
            try:
                if raw_qty is not None and str(raw_qty).strip() != "":
                    qty = int(float(str(raw_qty).replace(",", ".")))
                else:
                    qty = 0
            except (ValueError, TypeError):
                qty = 0

            # Skip rule: If empty Boutique or Globale is 0/empty, skip entirely
            if not magasin or magasin.startswith("TOTAL") or magasin == "NONE" or qty <= 0:
                continue

            # Column 3: Chauffeur -> chauffeur
            raw_driver = ""
            if col_chauffeur is not None and col_chauffeur in df.columns:
                val_ch = row.get(col_chauffeur)
                if val_ch is not None and pd.notna(val_ch):
                    raw_driver = str(val_ch).strip()
            if not raw_driver:
                for k in ['Chauffeur', 'Driver', 'CHAUFFEUR', 'DRIVER']:
                    if k in df.columns and pd.notna(row.get(k)):
                        raw_driver = str(row.get(k)).strip()
                        break
            driver_name = " ".join(raw_driver.split())

            # Customer Pickup (Ramassage) Workflow:
            raw_remarque = ""
            if col_remarque is not None and col_remarque in df.columns:
                val_rem = row.get(col_remarque)
                if val_rem is not None and pd.notna(val_rem):
                    raw_remarque = str(val_rem).strip()
            if not raw_remarque:
                for k in ['Remarque', 'Remarques', 'Observation', 'Observations', 'Note', 'Notes', 'REMARQUE']:
                    if k in df.columns and pd.notna(row.get(k)):
                        raw_remarque = str(row.get(k)).strip()
                        break

            remarque_clean = " ".join(raw_remarque.split()).lower()
            driver_lower = driver_name.lower()
            row_vals_list = [str(v).strip().lower() for v in row.values if pd.notna(v)]
            row_text = " ".join(row_vals_list)

            is_ramassage_pickup = (
                "ramassage" in remarque_clean or
                "rammasage" in remarque_clean or
                "ramassage" in row_text or
                "rammasage" in row_text or
                "retrait depot" in row_text or
                "retrait dépôt" in row_text or
                "retrait" in driver_lower or
                ((driver_lower in ("autre", "auter", "autres") or "autre" in driver_lower or "auter" in driver_lower) and
                 ("client" in row_text or "ramassage" in row_text or not remarque_clean))
            )

            if is_ramassage_pickup:
                driver_name = "RETRAIT DÉPÔT"

            # Strict Blacklist Check: Skip "bahlouli adel"
            driver_normalized = driver_name.lower().strip()
            if "bahlouli adel" in driver_normalized or "bahlouli" in driver_normalized or driver_name.upper() in BLACKLISTED_CHAUFFEURS:
                continue

            if driver_name == "RETRAIT DÉPÔT":
                chauffeur = "RETRAIT DÉPÔT"
            else:
                chauffeur = driver_name.upper() if driver_name else ""
                if chauffeur in ("NONE", "NULL", "-", "/", "NON SPÉCIFIÉ", "NON SPECIFIE", "NON ASSIGNÉ", "NON ASSIGNE", "SANS CHAUFFEUR"):
                    chauffeur = ""

            # Column 4: Livraison -> date (DD/MM/YYYY to YYYY-MM-DD)
            raw_date = None
            if col_livraison is not None and col_livraison in df.columns:
                val_date = row.get(col_livraison)
                if val_date is not None and pd.notna(val_date):
                    raw_date = val_date
            if raw_date is None:
                for k in ['Livraison', 'Date', 'LIVRAISON', 'DATE', 'Date Livraison']:
                    if k in df.columns and pd.notna(row.get(k)):
                        raw_date = row.get(k)
                        break
            normalized_date = normalize_date_to_iso(raw_date)

            # Column 5: Materiel -> safe extraction & handling potential invisible spaces
            mat_val = str(row.get('Materiel', str(row.get('Materiel ', '')))).strip()
            if not mat_val or mat_val.lower() == 'nan':
                for key_cand in ['Matériel', 'MATERIEL', 'MATÉRIEL', 'materiel', 'matériel', 'Vehicule', 'Véhicule', 'VEHICULE', 'Camion', 'CAMION', 'Engin', 'ENGIN']:
                    if key_cand in df.columns:
                        val_m = row.get(key_cand)
                        if val_m is not None and not (isinstance(val_m, float) and np.isnan(val_m)):
                            cand_s = str(val_m).strip()
                            if cand_s and cand_s.lower() not in ("nan", "none", "null", "-", "/"):
                                mat_val = cand_s
                                break
                if (not mat_val or mat_val.lower() == 'nan') and col_materiel is not None and col_materiel in df.columns:
                    val_m = row.get(col_materiel)
                    if val_m is not None and not (isinstance(val_m, float) and np.isnan(val_m)):
                        cand_s = str(val_m).strip()
                        if cand_s and cand_s.lower() not in ("nan", "none", "null", "-", "/"):
                            mat_val = cand_s

            materiel = mat_val if mat_val and mat_val.lower() != 'nan' and mat_val.lower() not in ("none", "null", "-", "/") else None

            # Occurrence tracker for deterministic unique delivery_id
            store_name = re.sub(r'[^A-Z0-9]', '', magasin.upper())
            driver_key_part = re.sub(r'[^A-Z0-9]', '', chauffeur.upper()) if chauffeur else "NONE"
            base_key = f"{normalized_date}_{store_name}_{driver_key_part}"
            occurrence_tracker[base_key] = occurrence_tracker.get(base_key, 0) + 1

            raw_delivery_id = f"LIV-{normalized_date}-{store_name}-{driver_key_part}-{occurrence_tracker[base_key]}"
            delivery_id = re.sub(r'[^A-Za-z0-9\-]', '', raw_delivery_id)
            delivery_key = delivery_id

            # CRITICAL FIX 2: Filter Existing Records (INSERT-ONLY)
            # If this delivery ID or key is already in the existing Supabase set,
            # completely SKIP the row to prevent overwriting live driver progress/photos/GPS.
            if skip_existing and existing_ids:
                if delivery_id.upper() in existing_ids or delivery_key.upper() in existing_ids:
                    skipped_existing_count += 1
                    continue

            record = {
                "id": delivery_id,
                "delivery_id": delivery_id,
                "date": normalized_date,
                "magasin": magasin,
                "chauffeur": chauffeur,
                "matricule": materiel or "",
                "materiel": materiel,
                "quantite": qty,
                "status": "Non Livré",
                "delivery_key": delivery_key
            }
            records.append(record)

        if skipped_existing_count > 0:
            print(f"[i] Safe Ingestion: Skipped {skipped_existing_count} existing record(s) in {os.path.basename(file_path)} to protect live Supabase state.")

        return {
            "success": len(errors) == 0,
            "total_records": len(records),
            "skipped_existing": skipped_existing_count,
            "records": records,
            "errors": errors
        }
    finally:
        xl.close()


def upsert_manifest_records(incoming_records, existing_records):
    """
    Appends all incoming valid delivery records directly without grouping
    or merging by Store, Driver, or Quantity. Status preservation relies
    strictly on exact delivery_id.
    """
    existing_by_id = {}
    for rec in existing_records:
        r_id = str(rec.get("delivery_id") or rec.get("id") or "").strip().upper()
        if r_id:
            existing_by_id[r_id] = rec

    stats = {
        "total": len(incoming_records),
        "preserved_livre": 0,
        "upgraded_livre": 0,
        "updated_non_livre": 0,
        "inserted_new": 0
    }

    final_list = []
    for item in incoming_records:
        chauf_check = str(item.get("chauffeur") or "").strip().lower()
        if "bahlouli adel" in chauf_check or "bahlouli" in chauf_check:
            continue
        new_item = dict(item)
        in_id = str(item.get("delivery_id") or item.get("id") or "").strip().upper()
        matched = existing_by_id.get(in_id) if in_id else None

        if matched:
            curr_status = matched.get("status", "Non Livré")
            incoming_status = item.get("status", "Non Livré")

            if curr_status == "Livré":
                final_status = "Livré"
                stats["preserved_livre"] += 1
            elif incoming_status == "Livré":
                final_status = "Livré"
                stats["upgraded_livre"] += 1
            else:
                final_status = "Non Livré"
                stats["updated_non_livre"] += 1

            new_item["status"] = final_status
            new_item["completed"] = (final_status == "Livré")
            if matched.get("quantite_recue") is not None:
                new_item["quantite_recue"] = matched["quantite_recue"]
            if matched.get("quantite_litige") is not None:
                new_item["quantite_litige"] = matched["quantite_litige"]
            if matched.get("statut_reserve"):
                new_item["statut_reserve"] = matched["statut_reserve"]
            if matched.get("photo_url"):
                new_item["photo_url"] = matched["photo_url"]
            if matched.get("recouvrement_url"):
                new_item["recouvrement_url"] = matched["recouvrement_url"]
            if matched.get("materiel") and not new_item.get("materiel"):
                new_item["materiel"] = matched["materiel"]
        else:
            if new_item.get("status") != "Livré":
                new_item["status"] = "Non Livré"
            new_item["completed"] = (new_item["status"] == "Livré")
            stats["inserted_new"] += 1

        final_list.append(new_item)

    return final_list, stats


def fetch_supabase_records(supabase_url=None, api_key=None):
    """
    Fetches all current deliveries from Supabase REST endpoint.
    Retrieves full record state to support status checks and verification.
    """
    url = (supabase_url or DEFAULT_SUPABASE_REST_URL).split("?")[0]
    key = api_key or DEFAULT_SUPABASE_API_KEY
    endpoint = f"{url}?select=delivery_id,delivery_key,date,magasin,chauffeur,quantite,status,delivered_by,materiel,photo_url,recouvrement_url,latitude,longitude,statut_reserve,quantite_recue,quantite_litige"
    headers = {
        "apikey": key,
        "Authorization": f"Bearer {key}"
    }
    try:
        req = urllib.request.Request(endpoint, headers=headers)
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except Exception as e:
        print(f"[!] Note: fetch_supabase_records encountered notice ({e})")
        return []


def push_manifest_to_supabase(records, supabase_url=None, api_key=None, only_new=True):
    """
    STRICTLY INSERT-ONLY SYNC:
    Safely appends only brand-new delivery records to Supabase.
    Existing records are NEVER overwritten, updated, or nullified, completely
    preserving driver photos (BL & recouvrement), live statuses, and GPS stamps.
    """
    url = (supabase_url or DEFAULT_SUPABASE_REST_URL).split("?")[0]
    key = api_key or DEFAULT_SUPABASE_API_KEY

    # 1. Fetch Existing State: Get all existing delivery_id and delivery_key values in a Python set
    existing_ids = fetch_existing_delivery_ids(url, key)

    if isinstance(records, str):
        parsed = parse_delivery_excel(records, skip_existing=True, existing_ids=existing_ids)
        records = parsed.get("records", [])

    if not records:
        print("[i] Insert-Only Sync: No incoming records to evaluate.")
        return {"status": "success", "count": 0, "message": "No incoming records to evaluate."}

    # 2. Filter Existing Records: Discard any record whose ID or key already exists in Supabase
    unpushed_records = []
    for r in records:
        rec_id = str(r.get("delivery_id") or r.get("id") or "").strip().upper()
        dkey = str(r.get("delivery_key") or "").strip().upper()
        if (rec_id and rec_id in existing_ids) or (dkey and dkey in existing_ids):
            continue
        unpushed_records.append(r)

    if not unpushed_records:
        print(f"[i] Insert-Only Sync: All {len(records)} record(s) already exist in Supabase. 0 rows inserted, 0 overwritten.")
        return {
            "status": "success",
            "count": 0,
            "message": "All records already exist in Supabase. Live statuses, photos, and GPS timestamps preserved."
        }

    # 3. Build uniform schema payload strictly for brand-new deliveries
    payload = []
    for r in unpushed_records:
        rec_id = str(r.get("delivery_id") or r.get("id") or "").strip()
        if not rec_id:
            continue
        try:
            qty = int(r.get("quantite", 0))
        except (ValueError, TypeError):
            qty = 0

        deliv_date = str(r.get("date") or "").strip()
        store = str(r.get("magasin") or "").strip()
        driver = str(r.get("chauffeur") or "").strip()
        if "bahlouli adel" in driver.lower() or "bahlouli" in driver.lower() or driver.upper() in BLACKLISTED_CHAUFFEURS:
            continue
        matricule = str(r.get("matricule") or "").strip()
        dkey = str(r.get("delivery_key") or "").strip() or rec_id

        mat_val = str(r.get("materiel") or r.get("Materiel") or r.get("Materiel ") or "").strip()
        materiel_val = mat_val if mat_val and mat_val.lower() != "nan" and mat_val.lower() not in ("none", "null", "-", "/") else None

        mapped = {
            "delivery_id": rec_id,
            "date": deliv_date,
            "magasin": store,
            "chauffeur": driver,
            "matricule": matricule or (materiel_val or ""),
            "materiel": materiel_val,
            "quantite": qty,
            "status": "Non Livré",
            "delivery_key": dkey,
            "delivered_by": None,
            "quantite_recue": None,
            "quantite_litige": None,
            "statut_reserve": None,
            "photo_url": None,
            "recouvrement_url": None,
            "latitude": None,
            "longitude": None
        }
        payload.append(mapped)

    if not payload:
        return {"status": "success", "count": 0, "message": "No new records to insert."}

    # 4. Never Overwrite: POST with resolution=ignore-duplicates
    headers = {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
        "Prefer": "resolution=ignore-duplicates"
    }

    endpoint_url = f"{url}?on_conflict=delivery_key"
    BATCH_SIZE = 100
    total_pushed = 0

    try:
        for i in range(0, len(payload), BATCH_SIZE):
            batch = payload[i:i + BATCH_SIZE]
            req_data = json.dumps(batch).encode("utf-8")
            req = urllib.request.Request(endpoint_url, data=req_data, headers=headers, method="POST")
            try:
                with urllib.request.urlopen(req, timeout=45) as resp:
                    pass
            except urllib.error.HTTPError as he:
                if he.code == 400 and endpoint_url != url:
                    fallback_req = urllib.request.Request(url, data=req_data, headers=headers, method="POST")
                    with urllib.request.urlopen(fallback_req, timeout=45) as resp:
                        pass
                else:
                    raise he
            total_pushed += len(batch)

        print(f"[+] Insert-Only Sync: Successfully inserted {total_pushed} brand-new record(s) to Supabase (0 overwritten).")
        return {"status": "success", "count": total_pushed}
    except urllib.error.HTTPError as e:
        err_body = e.read().decode("utf-8", errors="ignore")
        print(f"[-] Supabase HTTP Error {e.code}: {err_body}")
        return {"status": "error", "code": e.code, "message": str(e), "details": err_body}
    except urllib.error.URLError as e:
        print(f"[-] Supabase URL Error: {e}")
        return {"status": "error", "message": str(e)}
    except Exception as e:
        print(f"[-] Supabase Push Unexpected Error: {e}")
        return {"status": "error", "message": str(e)}


if __name__ == "__main__":
    import sys
    test_files = sys.argv[1:] if len(sys.argv) > 1 else []
    for tf in test_files:
        if os.path.exists(tf):
            print(f"\n--- Testing parse_delivery_excel on: {tf} ---")
            res = parse_delivery_excel(tf)
            print(f"Success: {res['success']}, Total Rows: {res['total_records']}")
            for r in res['records'][:5]:
                print(f"  [{r['delivery_id']}] {r['magasin']} | {r['chauffeur']} | {r['quantite']} p")
