#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tradifoot Delivery Tracker & Intelligence Engine
Automated ingestion, parsing, tracking, and synchronization for
Tradifoot delivery programs (Programme de livraison.xlsx & Partages).
Optimized for minimal RAM usage, fast streaming Excel parsing, and direct index.html synchronization.
"""

import os
import sys
import glob
import re
import json
import datetime
import webbrowser
import openpyxl

WORKSPACE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(WORKSPACE_DIR, "data")
PROGRAMS_DIR = os.path.join(DATA_DIR, "programs")
DB_PATH = os.path.join(DATA_DIR, "delivery_db.json")
INDEX_PATH = os.path.join(WORKSPACE_DIR, "index.html")

os.makedirs(PROGRAMS_DIR, exist_ok=True)

# Supabase REST Ingestion Integration
try:
    from parse_programme import (
        push_manifest_to_supabase,
        parse_delivery_excel,
        fetch_supabase_records,
        build_delivery_key,
        DEFAULT_SUPABASE_REST_URL,
        DEFAULT_SUPABASE_API_KEY
    )
except ImportError:
    sys.path.append(WORKSPACE_DIR)
    from parse_programme import (
        push_manifest_to_supabase,
        parse_delivery_excel,
        fetch_supabase_records,
        build_delivery_key,
        DEFAULT_SUPABASE_REST_URL,
        DEFAULT_SUPABASE_API_KEY
    )


def parse_excel_file_for_supabase(fpath):
    """
    Parses an individual delivery Excel file into flat records
    matching the Supabase deliveries schema.
    """
    # 1. Try template schema (parse_programme.py)
    try:
        res = parse_delivery_excel(fpath)
        if res.get("success") and res.get("records"):
            return res["records"]
    except Exception:
        pass

    # 2. Try Outlook delivery file format
    try:
        prog = parse_delivery_file(fpath)
        if not prog or not prog.get("destinations"):
            return []

        fallback_date = prog.get("received_date", "")[:10]
        m_dd = re.search(r'(\d{1,2})[-_](\d{1,2})[-_](\d{4})', prog.get("filename", ""))
        if m_dd:
            d_val, m_val = int(m_dd.group(1)), int(m_dd.group(2))
            if m_val > 12 and d_val <= 12:
                d_val, m_val = m_val, d_val
            m_val = max(1, min(12, m_val))
            d_val = max(1, min(31, d_val))
            fallback_date = f"{m_dd.group(3)}-{m_val:02d}-{d_val:02d}"
        else:
            m_short = re.search(r'(\d{1,2})[-_](\d{1,2})', prog.get("filename", ""))
            if m_short:
                yr = prog.get("year_month", "2026")[:4]
                d_val, m_val = int(m_short.group(1)), int(m_short.group(2))
                if m_val > 12 and d_val <= 12:
                    d_val, m_val = m_val, d_val
                m_val = max(1, min(12, m_val))
                d_val = max(1, min(31, d_val))
                fallback_date = f"{yr}-{m_val:02d}-{d_val:02d}"

        records = []
        slug = re.sub(r'[^a-zA-Z0-9]', '', prog.get("filename", ""))[:10].upper()
        for idx, d in enumerate(prog.get("destinations", []), 1):
            store = (d.get("store") or "").strip().upper()
            qty = d.get("quantity", 0)
            if not store or store.startswith("TOTAL") or qty <= 0:
                continue

            del_date = str(d.get("delivery_date") or "").strip()
            if not re.match(r'^\d{4}-\d{2}-\d{2}$', del_date):
                del_date = fallback_date

            driver = re.sub(r"\s+", " ", d.get("driver") or "Non spécifié").strip().upper()
            row_notes = " ".join([str(v or "").lower() for v in (d.get("remark"), d.get("remarque"), d.get("notes"), d.get("retour"), d.get("driver"), d.get("store")) if v])
            if "ramassage" in row_notes or "rammasage" in row_notes or "retrait depot" in row_notes or "retrait dépôt" in row_notes or "retrait" in driver.lower():
                driver = "RETRAIT DÉPÔT"
            if "bahlouli adel" in driver.lower() or "bahlouli" in driver.lower():
                continue
            vehicle = d.get("vehicle") or "Camion Porteur"
            if not vehicle or vehicle == "/":
                vehicle = "Camion Porteur"

            rec_id = f"DEL-{slug}-{idx:04d}"
            dkey = build_delivery_key(del_date, store, qty)
            records.append({
                "id": rec_id,
                "delivery_id": rec_id,
                "delivery_key": dkey,
                "date": del_date,
                "magasin": store,
                "chauffeur": driver,
                "matricule": vehicle,
                "materiel": vehicle,
                "quantite": qty,
                "status": "Non Livré"
            })
        records = [r for r in records if str(r.get("date", "")).strip() >= "2026-11-01"]
        return records
    except Exception as e:
        print(f"[-] Error extracting Supabase records from {fpath}: {e}")
        return []


def extract_routes_from_programs(programs, index_path=None):
    """
    Extracts canonical delivery routes from parsed programs with Live Status Preservation.
    Queries Supabase to retrieve the authoritative live status of all deliveries.
    """
    if index_path is None:
        index_path = INDEX_PATH

    all_routes = []
    existing_statuses_by_id = {}

    # 1. Primary Source of Truth: Live Supabase Deliveries
    try:
        supabase_records = fetch_supabase_records()
        for s_rec in supabase_records:
            s_id = str(s_rec.get("delivery_id") or s_rec.get("id") or "").strip()
            s_dkey = str(s_rec.get("delivery_key") or "").strip()
            s_status = str(s_rec.get("status") or "")
            s_completed = bool("livr" in s_status.lower() and "non" not in s_status.lower())
            rec_info = {
                "id": s_id,
                "status": "Livré" if s_completed else "Non Livré",
                "completed": s_completed,
                "delivered_by": s_rec.get("delivered_by")
            }
            if s_id:
                existing_statuses_by_id[s_id.upper()] = rec_info
            if s_dkey:
                existing_statuses_by_id[s_dkey] = rec_info
    except Exception as e:
        print(f"[!] Note: Fallback to local cache for statuses ({e})")

    # 2. Local fallback from index.html (if not found in Supabase)
    if os.path.exists(index_path):
        try:
            with open(index_path, "r", encoding="utf-8", newline="") as f:
                html_content = f.read()

            var_loc_start = html_content.find("const LOCAL_DELIVERY_DATA =")
            if var_loc_start != -1:
                bracket_start = html_content.find("[", var_loc_start)
                next_var = html_content.find("let DELIVERY_DATA =", bracket_start)
                if next_var != -1:
                    semicolon_pos = html_content.rfind("];", bracket_start, next_var)
                    if bracket_start != -1 and semicolon_pos != -1:
                        raw_existing_json = html_content[bracket_start:semicolon_pos + 1]
                        existing_list = json.loads(raw_existing_json)
                        for item in existing_list:
                            i_id = str(item.get("id") or item.get("delivery_id") or "").strip()
                            i_status = str(item.get("status") or "")
                            i_completed = bool(item.get("completed") or ("livr" in i_status.lower() and "non" not in i_status.lower()))
                            rec_info = {
                                "id": i_id,
                                "status": "Livré" if i_completed else "Non Livré",
                                "completed": i_completed
                            }
                            if i_id and i_id.upper() not in existing_statuses_by_id:
                                existing_statuses_by_id[i_id.upper()] = rec_info
        except Exception:
            pass

    occurrence_tracker = {}
    for p in programs:
        m_dd = re.search(r'(\d{1,2})[-_](\d{1,2})[-_](\d{4})', p["filename"])
        if m_dd:
            d_val, m_val = int(m_dd.group(1)), int(m_dd.group(2))
            if m_val > 12 and d_val <= 12:
                d_val, m_val = m_val, d_val
            m_val = max(1, min(12, m_val))
            d_val = max(1, min(31, d_val))
            fallback_date = f"{m_dd.group(3)}-{m_val:02d}-{d_val:02d}"
        else:
            m_short = re.search(r'(\d{1,2})[-_](\d{1,2})', p["filename"])
            if m_short:
                yr = p.get("year_month", "2026")[:4]
                d_val, m_val = int(m_short.group(1)), int(m_short.group(2))
                if m_val > 12 and d_val <= 12:
                    d_val, m_val = m_val, d_val
                m_val = max(1, min(12, m_val))
                d_val = max(1, min(31, d_val))
                fallback_date = f"{yr}-{m_val:02d}-{d_val:02d}"
            else:
                fallback_date = p["received_date"][:10]

        p_title = p.get("display_title", "Programme")

        for d in p.get("destinations", []):
            raw_deliv_date = str(d.get("delivery_date") or "").strip()
            deliv_date = raw_deliv_date if re.match(r'^\d{4}-\d{2}-\d{2}$', raw_deliv_date) else fallback_date

            qty = d.get("quantity", 0)
            store_clean = (d.get("store") or "").strip().upper()
            if not store_clean or store_clean.startswith("TOTAL") or qty <= 0:
                continue

            driver = re.sub(r"\s+", " ", d.get("driver") or "Non spécifié").strip().upper()
            if "bahlouli adel" in driver.lower() or "bahlouli" in driver.lower():
                continue
            vehicle = d.get("vehicle") or "Camion Porteur"
            if not vehicle or vehicle == "/":
                vehicle = "Camion Porteur"

            store_name = re.sub(r'[^A-Z0-9]', '', store_clean.upper())
            chauf_for_id = driver if not re.search(r'non|sans|unassigned|^none$|^null$|^\-$|^\/$', driver, re.I) else ""
            driver_name = re.sub(r'[^A-Z0-9]', '', chauf_for_id.upper()) if chauf_for_id else "NONE"
            base_key = f"{deliv_date}_{store_name}_{driver_name}"
            occurrence_tracker[base_key] = occurrence_tracker.get(base_key, 0) + 1
            raw_delivery_id = f"LIV-{deliv_date}-{store_name}-{driver_name}-{occurrence_tracker[base_key]}"
            route_id = re.sub(r'[^A-Za-z0-9\-]', '', raw_delivery_id)

            matched_rec = existing_statuses_by_id.get(route_id.upper())
            if matched_rec and (matched_rec.get("status") == "Livré" or matched_rec.get("completed") is True):
                final_status = "Livré"
                final_completed = True
            else:
                final_status = "Non Livré"
                final_completed = False

            all_routes.append({
                "id": route_id,
                "delivery_id": route_id,
                "delivery_key": route_id,
                "date": deliv_date,
                "magasin": d["store"],
                "destination": d["store"],
                "chauffeur": driver,
                "driver": driver,
                "matricule": vehicle,
                "vehicle": vehicle,
                "materiel": vehicle,
                "quantite": qty,
                "quantity": qty,
                "status": final_status,
                "completed": final_completed
            })

    return all_routes


# -------------------------------------------------------------------------
# OUTLOOK SYNC MODULE
# -------------------------------------------------------------------------
def sync_from_outlook(days_back=2, full_scan=False):
    """
    Connects to Outlook Classic via COM and searches for delivery programs.
    Strict Date Boundary: Only programs dated on or after November 1, 2026.
    """
    MIN_INGESTION_DATE = datetime.datetime(2026, 11, 1, 0, 0, 0)
    lookback = 90 if full_scan else days_back
    print(f"[1/2] Connecting to Outlook Classic (Window: {lookback} days, Min: 2026-11-01)...")
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
    except Exception as e:
        print(f"[-] Could not connect to Outlook via COM: {e}")
        print("[!] Using existing local programs...")
        return 0

    print(f"[+] Connected to Outlook as: {namespace.CurrentUser.Name}")
    root = namespace.Stores[0].GetRootFolder()

    target_folders = ["GDS", "COMMERCIALS", "Inbox", "Deleted Items"]
    downloaded_count = 0
    downloaded_files = []
    now = datetime.datetime.now()
    existing_programs = set(os.listdir(PROGRAMS_DIR)) if os.path.exists(PROGRAMS_DIR) else set()

    def scan_folder(folder, path):
        nonlocal downloaded_count
        try:
            items = folder.Items
            try:
                items.Sort("[ReceivedTime]", True)
            except Exception:
                pass

            for item in items:
                try:
                    rec_time = item.ReceivedTime
                    dt = datetime.datetime(rec_time.year, rec_time.month, rec_time.day, rec_time.hour, rec_time.minute)
                    if (now - dt).days > lookback or dt < MIN_INGESTION_DATE:
                        break

                    subj = str(item.Subject or "")
                    if item.Attachments and item.Attachments.Count > 0:
                        for i in range(1, item.Attachments.Count + 1):
                            att = item.Attachments.Item(i)
                            fn = str(att.FileName or "")

                            is_prog_file = bool(re.search(r'prog|poro|livr', fn, re.I)) and fn.lower().endswith(('.xlsx', '.xls'))
                            is_prog_subj = bool(re.search(r'programme|programe|livraison', subj, re.I)) and fn.lower().endswith(('.xlsx', '.xls'))

                            if is_prog_file or is_prog_subj:
                                m_d = re.search(r'(\d{1,2})[-_](\d{1,2})[-_](\d{2,4})', fn)
                                if m_d:
                                    try:
                                        d_p, m_p, y_p = int(m_d.group(1)), int(m_d.group(2)), int(m_d.group(3))
                                        if y_p < 100:
                                            y_p += 2000
                                        if m_p > 12 and d_p <= 12:
                                            d_p, m_p = m_p, d_p
                                        file_dt = datetime.datetime(y_p, m_p, d_p)
                                        if file_dt < MIN_INGESTION_DATE:
                                            continue
                                    except Exception:
                                        pass

                                date_str = dt.strftime("%Y%m%d_%H%M")
                                safe_fn = f"{date_str}_{fn}".replace("/", "_").replace("\\", "_").replace(":", "_")
                                save_path = os.path.join(PROGRAMS_DIR, safe_fn)

                                if safe_fn in existing_programs or os.path.exists(save_path):
                                    continue

                                att.SaveAsFile(save_path)
                                downloaded_count += 1
                                downloaded_files.append(save_path)
                                existing_programs.add(safe_fn)
                                print(f"  [+] Downloaded new program: {safe_fn}")
                except Exception:
                    pass
        except Exception:
            pass

        try:
            for sub in folder.Folders:
                scan_folder(sub, f"{path}/{sub.Name}")
        except Exception:
            pass

    for f in root.Folders:
        if f.Name in target_folders:
            scan_folder(f, f.Name)

    if downloaded_count > 0:
        print(f"[+] Outlook sync complete: {downloaded_count} new file(s) found.")
    else:
        print("[+] Outlook sync complete: All programs are up to date.")

    if downloaded_files:
        print(f"[*] Ingesting and pushing {len(downloaded_files)} new delivery file(s) to Supabase...")
        for new_file in downloaded_files:
            try:
                new_records = parse_excel_file_for_supabase(new_file)
                if new_records:
                    res = push_manifest_to_supabase(new_records, only_new=True)
                    pushed_n = res.get("count", 0)
                    print(f"  [+] Synced {pushed_n} new delivery record(s) from '{os.path.basename(new_file)}' to Supabase.")
            except Exception as e:
                print(f"  [-] Error pushing '{os.path.basename(new_file)}' to Supabase: {e}")

    return downloaded_count


# -------------------------------------------------------------------------
# EXCEL PARSER MODULE (STREAMING / READ_ONLY)
# -------------------------------------------------------------------------
def parse_delivery_file(fpath):
    """
    Parses a single delivery program Excel file using openpyxl read_only=True
    for minimum memory footprint.
    """
    fname = os.path.basename(fpath)
    m_date = re.match(r'^(\d{4})(\d{2})(\d{2})_(\d{2})(\d{2})_', fname)
    if m_date:
        rec_date = f"{m_date.group(1)}-{m_date.group(2)}-{m_date.group(3)} {m_date.group(4)}:{m_date.group(5)}"
        year_month = f"{m_date.group(1)}-{m_date.group(2)}"
    else:
        rec_date = datetime.datetime.fromtimestamp(os.path.getmtime(fpath)).strftime("%Y-%m-%d %H:%M")
        year_month = rec_date[:7]

    try:
        wb = openpyxl.load_workbook(fpath, data_only=True, read_only=True)
    except Exception:
        return None

    try:
        prog_sheet_name = None
        for s in wb.sheetnames:
            if re.search(r'prog|poro|livr', s, re.I):
                prog_sheet_name = s
                break
        if not prog_sheet_name:
            prog_sheet_name = wb.sheetnames[0]

        ws_prog = wb[prog_sheet_name]
        rows = list(ws_prog.iter_rows(values_only=True))
        if not rows:
            return None

        h_idx = -1
        for i, r in enumerate(rows[:7]):
            line = ' '.join([str(x) for x in r if x is not None])
            if re.search(r'magasin|boutique|destinataire|wilaya|chauffeur|priorit|p\b|lignes', line, re.I):
                h_idx = i
                break

        if h_idx == -1:
            return None

        headers = [str(x).strip().lower() if x is not None else '' for x in rows[h_idx]]

        col_store = next((i for i, h in enumerate(headers) if any(k in h for k in ['magasin', 'boutique', 'destinataire', 'lignes'])), None)
        col_wilaya = next((i for i, h in enumerate(headers) if 'wilaya' in h), None)
        col_region = next((i for i, h in enumerate(headers) if 'region' in h or 'rgion' in h), None)
        col_driver = next((i for i, h in enumerate(headers) if 'chauffeur' in h), None)
        col_vehicle = next((i for i, h in enumerate(headers) if 'materiel' in h or 'vehicule' in h), None)
        col_prio = next((i for i, h in enumerate(headers) if h == 'p' or 'priorit' in h), None)
        col_total = next((i for i, h in enumerate(headers) if any(k in h for k in ['globale', 'global', 'total', 'qt', 'quant', 'qte'])), None)
        col_ret = next((i for i, h in enumerate(headers) if 'retour' in h), None)
        col_rem = next((i for i, h in enumerate(headers) if 'remarque' in h), None)
        col_deliv_date = next((i for i, h in enumerate(headers) if 'livraison' in h), None)

        cat_cols = {}
        known_cats = ['footwear', 'chaussettes', 'pantoufle', 'slipper', 'accessoires', 'sac a dos', 'textille', 'sachet', 'claquette']
        for idx, h in enumerate(headers):
            for kc in known_cats:
                if kc in h:
                    cat_cols[kc.capitalize()] = idx

        destinations = []
        tot_units = 0
        drivers_in_prog = set()
        vehicles_in_prog = set()
        categories_totals = {}

        for r in rows[h_idx+1:]:
            if not any(r):
                continue
            store_name = str(r[col_store]).strip() if col_store is not None and r[col_store] is not None else ''
            if not store_name or store_name.lower().startswith('total') or store_name == 'None':
                continue

            row_qty = 0
            if col_total is not None and r[col_total] is not None:
                try:
                    row_qty = int(float(r[col_total]))
                except Exception:
                    row_qty = 0

            cat_breakdown = {}
            for cat_name, c_idx in cat_cols.items():
                val = r[c_idx]
                if val is not None:
                    try:
                        q = int(float(val))
                        cat_breakdown[cat_name] = q
                        categories_totals[cat_name] = categories_totals.get(cat_name, 0) + q
                    except Exception:
                        pass

            if row_qty == 0 and cat_breakdown:
                row_qty = sum(cat_breakdown.values())

            tot_units += row_qty

            driver = str(r[col_driver]).strip().upper() if col_driver is not None and r[col_driver] else ''
            rem = str(r[col_rem]).strip() if col_rem is not None and r[col_rem] else ''
            row_text = " ".join([str(c or "").strip().lower() for c in r if c is not None])
            d_low = driver.lower()
            is_ramassage_pickup = (
                "ramassage" in row_text or
                "rammasage" in row_text or
                "retrait depot" in row_text or
                "retrait dépôt" in row_text or
                "retrait" in d_low or
                ((d_low in ('autre', 'auter', 'autres') or 'autre' in d_low or 'auter' in d_low) and
                 ('client' in row_text or 'ramassage' in row_text or not rem))
            )
            if is_ramassage_pickup:
                driver = 'RETRAIT DÉPÔT'
            if "bahlouli adel" in driver.lower() or "bahlouli" in driver.lower():
                continue
            vehicle = str(r[col_vehicle]).strip().upper() if col_vehicle is not None and r[col_vehicle] else ''
            if driver and 'TOTAL' not in driver:
                drivers_in_prog.add(driver)
            if vehicle and 'TOTAL' not in vehicle:
                vehicles_in_prog.add(vehicle)

            deliv_date_str = ""
            if col_deliv_date is not None and r[col_deliv_date]:
                v_date = r[col_deliv_date]
                if isinstance(v_date, (datetime.datetime, datetime.date)):
                    deliv_date_str = v_date.strftime("%Y-%m-%d")
                else:
                    deliv_date_str = str(v_date)[:10]

            destinations.append({
                "store": store_name,
                "wilaya": str(r[col_wilaya]).strip() if col_wilaya is not None and r[col_wilaya] else '',
                "region": str(r[col_region]).strip() if col_region is not None and r[col_region] else '',
                "driver": driver,
                "vehicle": vehicle,
                "priority": r[col_prio] if col_prio is not None and r[col_prio] is not None else '',
                "quantity": row_qty,
                "categories": cat_breakdown,
                "delivery_date": deliv_date_str,
                "retour": str(r[col_ret]).strip() if col_ret is not None and r[col_ret] else '',
                "remark": str(r[col_rem]).strip() if col_rem is not None and r[col_rem] else ''
            })

        # Brands analysis from Partage / System sheet
        brands = {}
        part_sheet = next((s for s in wb.sheetnames if re.search(r'partage|system', s, re.I)), None)
        if part_sheet:
            try:
                ws_p = wb[part_sheet]
                p_rows = list(ws_p.iter_rows(values_only=True))
                if len(p_rows) > 2:
                    p_h = [str(x).strip().lower() if x is not None else '' for x in p_rows[1]]
                    b_col = next((i for i, h in enumerate(p_h) if 'marque' in h), None)
                    q_col = next((i for i, h in enumerate(p_h) if 'quantit' in h or 'qt' in h), None)
                    if b_col is not None and q_col is not None:
                        for pr in p_rows[2:]:
                            b = str(pr[b_col]).strip().upper() if pr[b_col] else ''
                            if b and 'TOTAL' not in b:
                                try:
                                    q = int(float(pr[q_col]))
                                    brands[b] = brands.get(b, 0) + q
                                except Exception:
                                    pass
            except Exception:
                pass

        discrepancies_count = 0
        rem_sheet = next((s for s in wb.sheetnames if re.search(r'enlever', s, re.I)), None)
        if rem_sheet:
            try:
                ws_r = wb[rem_sheet]
                discrepancies_count = max(0, sum(1 for _ in ws_r.iter_rows(values_only=True)) - 1)
            except Exception:
                pass

        title_match = re.search(r'(?:programme|programe|porograme)\s*(.*?)\.xlsx', fname, re.I)
        display_title = title_match.group(1).strip() if title_match else fname
        if not display_title or display_title.lower() == 'xlsx':
            display_title = f"Vague {rec_date[:10]}"

        return {
            "id": fname,
            "filename": fname,
            "filepath": fpath,
            "received_date": rec_date,
            "year_month": year_month,
            "display_title": display_title.replace("_", " ").title(),
            "total_units": tot_units,
            "destinations_count": len(destinations),
            "drivers": sorted(list(drivers_in_prog)),
            "vehicles": sorted(list(vehicles_in_prog)),
            "brands": brands,
            "categories": categories_totals,
            "discrepancies_count": discrepancies_count,
            "destinations": destinations
        }
    finally:
        wb.close()


def build_database():
    """
    Scans data/programs for delivery program Excel files (November 2026 onwards),
    parses them, and returns structured database records.
    """
    print("[2/2] Scanning and indexing delivery programs...")
    all_files = sorted(list(set(glob.glob(os.path.join(PROGRAMS_DIR, "*.xls*")))))

    programs = []
    store_stats = {}
    driver_stats = {}
    brand_stats = {}

    for fpath in all_files:
        fn = os.path.basename(fpath)
        if re.search(r'manque|sachet|evenement|notation|inventaire', fn, re.I):
            continue

        res = parse_delivery_file(fpath)
        if not res or res["destinations_count"] == 0:
            continue

        prog_date = str(res.get("received_date", ""))[:10]
        dest_dates = [str(d.get("delivery_date", "")).strip() for d in res.get("destinations", []) if d.get("delivery_date")]
        max_dest_date = max(dest_dates) if dest_dates else prog_date
        if max_dest_date and max_dest_date < "2026-11-01":
            continue
        programs.append(res)

        for d in res["destinations"]:
            st = d["store"]
            if st:
                if st not in store_stats:
                    store_stats[st] = {"store": st, "total_units": 0, "deliveries_count": 0, "wilaya": d["wilaya"], "region": d["region"], "top_driver": d["driver"]}
                store_stats[st]["total_units"] += d["quantity"]
                store_stats[st]["deliveries_count"] += 1

        for dr in res["drivers"]:
            driver_stats[dr] = driver_stats.get(dr, 0) + 1

        for b, q in res["brands"].items():
            brand_stats[b] = brand_stats.get(b, 0) + q

    programs.sort(key=lambda x: x["received_date"], reverse=True)

    db_data = {
        "metadata": {
            "last_updated": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "total_programs": len(programs),
            "total_dispatched_units": sum(p["total_units"] for p in programs),
            "total_unique_stores": len(store_stats),
            "total_drivers": len(driver_stats)
        },
        "programs": programs,
        "stores": sorted(list(store_stats.values()), key=lambda x: x["total_units"], reverse=True),
        "drivers": driver_stats,
        "brands": brand_stats
    }

    if os.path.exists(DATA_DIR):
        try:
            with open(DB_PATH, "w", encoding="utf-8") as fp:
                json.dump(db_data, fp, ensure_ascii=False, indent=2)
        except Exception:
            pass

    print(f"[+] Indexing complete: {len(programs)} program(s) indexed ({db_data['metadata']['total_dispatched_units']:,} total units).")
    return db_data


# -------------------------------------------------------------------------
# CALENDAR APP (index.html) SYNCHRONIZATION MODULE
# -------------------------------------------------------------------------
def update_calendar_app(db_data, index_path=None):
    """
    Synchronizes index.html by injecting the live DELIVERY_DATA extracted
    from the parsed delivery programs.
    """
    if index_path is None:
        index_path = INDEX_PATH

    if not os.path.exists(index_path):
        print(f"[-] {index_path} not found. Skipping calendar sync.")
        return None

    today_str = datetime.date.today().strftime("%Y-%m-%d")
    all_routes = extract_routes_from_programs(db_data.get("programs", []), index_path=index_path)

    with open(index_path, "r", encoding="utf-8", newline="") as f:
        html_content = f.read()

    active_date = today_str
    y_sel, m_sel, d_sel = [int(x) for x in active_date.split("-")]
    json_routes = json.dumps(all_routes, ensure_ascii=False, indent=2)

    var_loc_start = html_content.find("const LOCAL_DELIVERY_DATA =")
    if var_loc_start != -1:
        bracket_start = html_content.find("[", var_loc_start)
        next_var = html_content.find("let DELIVERY_DATA =", bracket_start)
        if next_var != -1:
            semicolon_pos = html_content.rfind("];", bracket_start, next_var)
            if bracket_start != -1 and semicolon_pos != -1:
                html_content = html_content[:bracket_start] + json_routes + html_content[semicolon_pos + 1:]

    pattern_state = r'let currentViewDate\s*=\s*new Date\([^)]+\);(\s*)let selectedDateStr\s*=\s*"[^"]+";'
    replacement_state = f'let currentViewDate = new Date({y_sel}, {m_sel - 1}, {d_sel});\\1let selectedDateStr = "{active_date}";'
    new_html = re.sub(pattern_state, replacement_state, html_content)

    with open(index_path, "w", encoding="utf-8", newline="") as f:
        f.write(new_html)

    print(f"[+] Application (index.html) synchronized with {len(all_routes)} live delivery routes.")
    return index_path


# -------------------------------------------------------------------------
# CLI ENTRY POINT
# -------------------------------------------------------------------------
def main():
    import argparse
    parser = argparse.ArgumentParser(description="Tradifoot Delivery Tracker & Intelligence Engine")
    parser.add_argument("--sync", action="store_true", help="Sync latest programs from Outlook")
    parser.add_argument("--full", action="store_true", help="Perform full historical Outlook sync")
    parser.add_argument("--open", action="store_true", help="Open application in default browser")
    parser.add_argument("--stats", action="store_true", help="Display summary terminal stats")

    args = parser.parse_args()

    if args.sync or args.full:
        sync_from_outlook(days_back=2, full_scan=args.full)

    db_data = build_database()
    cal_path = update_calendar_app(db_data)

    if args.stats:
        m = db_data["metadata"]
        print("\n" + "="*50)
        print(" TRADIFOOT LOGISTICS INTELLIGENCE SUMMARY")
        print("="*50)
        print(f" Total Programs Tracked : {m['total_programs']}")
        print(f" Total Units Dispatched : {m['total_dispatched_units']:,}")
        print(f" Stores in Network      : {m['total_unique_stores']}")
        print(f" Active Drivers         : {m['total_drivers']}")
        print("="*50 + "\n")

    if args.open:
        target = cal_path or INDEX_PATH
        print(f"[*] Opening application in default web browser: {target}")
        webbrowser.open(target)


if __name__ == "__main__":
    main()
