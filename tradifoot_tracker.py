#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tradifoot Delivery Tracker & Intelligence Engine
Automated ingestion, parsing, tracking, and dashboard generator for
Tradifoot delivery programs (Programme de livraison.xlsx & Partages).
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
DASHBOARD_PATH = os.path.join(WORKSPACE_DIR, "dashboard.html")

os.makedirs(PROGRAMS_DIR, exist_ok=True)

# -------------------------------------------------------------------------
# OUTLOOK SYNC MODULE
# -------------------------------------------------------------------------
def sync_from_outlook(days_back=90, full_scan=False):
    """
    Connects to Outlook Classic via COM and searches for delivery programs.
    Downloads new Excel attachments into data/programs/.
    """
    print("[1/3] Connecting to Outlook Classic...")
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
    except Exception as e:
        print(f"[-] Could not connect to Outlook via COM: {e}")
        print("[!] Falling back to existing cached files in workspace...")
        return 0

    print(f"[+] Connected to Outlook as: {namespace.CurrentUser.Name}")
    root = namespace.Stores[0].GetRootFolder()

    target_folders = ["GDS", "COMMERCIALS", "Inbox", "Deleted Items"]
    downloaded_count = 0
    now = datetime.datetime.now()

    def scan_folder(folder, path):
        nonlocal downloaded_count
        try:
            items = folder.Items
            for item in items:
                try:
                    rec_time = item.ReceivedTime
                    # Convert pywintypes datetime to python datetime if needed
                    dt = datetime.datetime(rec_time.year, rec_time.month, rec_time.day, rec_time.hour, rec_time.minute)
                    if not full_scan and (now - dt).days > days_back:
                        continue

                    subj = str(item.Subject or "")
                    if item.Attachments and item.Attachments.Count > 0:
                        for i in range(1, item.Attachments.Count + 1):
                            att = item.Attachments.Item(i)
                            fn = str(att.FileName or "")
                            
                            is_prog_file = bool(re.search(r'prog|poro|livr', fn, re.I)) and fn.lower().endswith(('.xlsx', '.xls'))
                            is_prog_subj = bool(re.search(r'programme|programe|livraison', subj, re.I)) and fn.lower().endswith(('.xlsx', '.xls'))
                            
                            if is_prog_file or is_prog_subj:
                                date_str = dt.strftime("%Y%m%d_%H%M")
                                safe_fn = f"{date_str}_{fn}".replace("/", "_").replace("\\", "_").replace(":", "_")
                                save_path = os.path.join(PROGRAMS_DIR, safe_fn)
                                if not os.path.exists(save_path):
                                    att.SaveAsFile(save_path)
                                    downloaded_count += 1
                                    print(f"  [+] Downloaded: {safe_fn}")
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
            print(f"[*] Scanning {f.Name}...")
            scan_folder(f, f.Name)

    print(f"[+] Outlook sync complete: {downloaded_count} new file(s) retrieved.")
    return downloaded_count


# -------------------------------------------------------------------------
# EXCEL PARSER MODULE
# -------------------------------------------------------------------------
def parse_delivery_file(fpath):
    """
    Parses a single delivery program Excel file and extracts structured records.
    """
    fname = os.path.basename(fpath)
    
    # Extract date from filename if prefixed with YYYYMMDD_HHMM
    m_date = re.match(r'^(\d{4})(\d{2})(\d{2})_(\d{2})(\d{2})_', fname)
    if m_date:
        rec_date = f"{m_date.group(1)}-{m_date.group(2)}-{m_date.group(3)} {m_date.group(4)}:{m_date.group(5)}"
        year_month = f"{m_date.group(1)}-{m_date.group(2)}"
    else:
        rec_date = datetime.datetime.fromtimestamp(os.path.getmtime(fpath)).strftime("%Y-%m-%d %H:%M")
        year_month = rec_date[:7]

    try:
        wb = openpyxl.load_workbook(fpath, data_only=True, read_only=False)
    except Exception as e:
        return None

    # Detect program sheet
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
        wb.close()
        return None

    # Find table header
    h_idx = -1
    for i, r in enumerate(rows[:7]):
        line = ' '.join([str(x) for x in r if x is not None])
        if re.search(r'magasin|boutique|destinataire|wilaya|chauffeur|priorit|p\b|lignes', line, re.I):
            h_idx = i
            break

    if h_idx == -1:
        wb.close()
        return None

    headers = [str(x).strip().lower() if x is not None else '' for x in rows[h_idx]]

    # Map column positions
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

    # Subcategory columns
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

        # Compute row quantity
        row_qty = 0
        if col_total is not None and r[col_total] is not None:
            try:
                row_qty = int(float(r[col_total]))
            except Exception:
                row_qty = 0
        
        # If total column was 0 or not found, try summing category columns
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
        vehicle = str(r[col_vehicle]).strip().upper() if col_vehicle is not None and r[col_vehicle] else ''
        if driver and 'TOTAL' not in driver: drivers_in_prog.add(driver)
        if vehicle and 'TOTAL' not in vehicle: vehicles_in_prog.add(vehicle)

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

    # Discrepancies removed sheet
    discrepancies_count = 0
    rem_sheet = next((s for s in wb.sheetnames if re.search(r'enlever', s, re.I)), None)
    if rem_sheet:
        ws_r = wb[rem_sheet]
        discrepancies_count = max(0, ws_r.max_row - 1)

    wb.close()

    is_modified = bool(re.search(r'modif|copy|\(2\)', fname, re.I))

    # Form friendly display title
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
        "is_modified": is_modified,
        "total_units": tot_units,
        "destinations_count": len(destinations),
        "drivers": sorted(list(drivers_in_prog)),
        "vehicles": sorted(list(vehicles_in_prog)),
        "brands": brands,
        "categories": categories_totals,
        "discrepancies_count": discrepancies_count,
        "destinations": destinations
    }


def build_database():
    """
    Scans data/programs and any workspace folders for delivery program Excel files,
    parses them, and writes the complete structured JSON database.
    """
    print("[2/3] Building distribution database...")
    all_files = glob.glob(os.path.join(PROGRAMS_DIR, "*.xls*"))
    
    # Also include existing files in all_delivery_programs or delivery_september if not yet copied
    legacy_folders = [
        os.path.join(WORKSPACE_DIR, "all_delivery_programs"),
        os.path.join(WORKSPACE_DIR, "delivery_september")
    ]
    for leg in legacy_folders:
        if os.path.exists(leg):
            for lf in glob.glob(os.path.join(leg, "*.xls*")):
                base = os.path.basename(lf)
                target = os.path.join(PROGRAMS_DIR, base)
                if not os.path.exists(target):
                    import shutil
                    shutil.copy2(lf, target)
                    all_files.append(target)

    all_files = sorted(list(set(glob.glob(os.path.join(PROGRAMS_DIR, "*.xls*")))))

    programs = []
    store_stats = {}
    driver_stats = {}
    brand_stats = {}

    for fpath in all_files:
        fn = os.path.basename(fpath)
        # Filter out obvious non-program files
        if re.search(r'manque|sachet|evenement|notation|inventaire', fn, re.I):
            continue

        # Check filename or sheets for delivery program content
        is_prog = bool(re.search(r'prog|poro|livr', fn, re.I))
        if not is_prog:
            try:
                wb_chk = openpyxl.load_workbook(fpath, read_only=True)
                is_prog = any(re.search(r'prog|poro|livr', s, re.I) for s in wb_chk.sheetnames)
                wb_chk.close()
            except Exception:
                is_prog = False

        if not is_prog:
            continue

        res = parse_delivery_file(fpath)
        if res and res["destinations_count"] > 0:
            programs.append(res)
            
            # Aggregate stats
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

    # Sort programs chronologically descending
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

    with open(DB_PATH, "w", encoding="utf-8") as fp:
        json.dump(db_data, fp, ensure_ascii=False, indent=2)

    print(f"[+] Database updated: {len(programs)} programs indexed ({db_data['metadata']['total_dispatched_units']:,} total units).")
    return db_data


# -------------------------------------------------------------------------
# INTERACTIVE DASHBOARD HTML GENERATOR
# -------------------------------------------------------------------------
def generate_dashboard(db_data):
    """
    Generates a standalone, beautiful HTML dashboard with interactive charts,
    filters, search, and detailed delivery program modal inspectors.
    """
    print("[3/3] Generating Interactive Distribution Dashboard...")

    db_json_str = json.dumps(db_data, ensure_ascii=False)

    html_template = f"""<!DOCTYPE html>
<html lang="fr" class="h-full bg-slate-900 text-slate-100">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Tradifoot Delivery Tracker & Intelligence Dashboard</title>
  <!-- Tailwind CSS & Lucide Icons -->
  <script src="https://cdn.tailwindcss.com"></script>
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <script src="https://unpkg.com/lucide@latest"></script>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    body {{ font-family: 'Inter', sans-serif; }}
    ::-webkit-scrollbar {{ width: 8px; height: 8px; }}
    ::-webkit-scrollbar-track {{ background: #0f172a; }}
    ::-webkit-scrollbar-thumb {{ background: #334155; border-radius: 4px; }}
    ::-webkit-scrollbar-thumb:hover {{ background: #475569; }}
  </style>
</head>
<body class="h-full flex flex-col overflow-hidden bg-slate-950">

  <!-- TOP APP HEADER -->
  <header class="bg-slate-900 border-b border-slate-800 px-6 py-3.5 flex items-center justify-between shrink-0 shadow-lg">
    <div class="flex items-center space-x-3">
      <div class="h-10 w-10 rounded-xl bg-gradient-to-tr from-indigo-600 to-cyan-500 flex items-center justify-center text-white shadow-md shadow-indigo-500/20">
        <i data-lucide="truck" class="w-5 h-5"></i>
      </div>
      <div>
        <h1 class="text-lg font-bold tracking-tight text-white flex items-center gap-2">
          TRADIFOOT <span class="text-xs px-2 py-0.5 rounded-full bg-indigo-500/20 text-indigo-400 font-semibold border border-indigo-500/30">Delivery Tracker</span>
        </h1>
        <p class="text-xs text-slate-400">Suivi & Intelligence des Programmes de Livraison (Dépôt Hamiz)</p>
      </div>
    </div>

    <!-- Live Metadata & Controls -->
    <div class="flex items-center space-x-4">
      <div class="hidden md:flex items-center space-x-2 text-xs text-slate-400 bg-slate-800/60 px-3 py-1.5 rounded-lg border border-slate-700/50">
        <span class="inline-block w-2 h-2 rounded-full bg-emerald-400 animate-pulse"></span>
        <span>MAJ : <strong id="last-update" class="text-slate-200"></strong></span>
      </div>
      <button onclick="refreshDashboard()" class="flex items-center space-x-1.5 bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs px-3 py-1.5 rounded-lg border border-slate-700 transition font-medium">
        <i data-lucide="refresh-cw" class="w-3.5 h-3.5"></i>
        <span>Actualiser</span>
      </button>
    </div>
  </header>

  <!-- MAIN SCROLLABLE CONTENT -->
  <main class="flex-1 overflow-y-auto p-6 space-y-6">

    <!-- EXECUTIVE KPI CARDS -->
    <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
      <div class="bg-slate-900/90 border border-slate-800 rounded-xl p-4 shadow-sm hover:border-slate-700 transition">
        <div class="flex items-center justify-between">
          <span class="text-xs font-semibold uppercase tracking-wider text-slate-400">Vagues de Livraison</span>
          <div class="p-2 rounded-lg bg-indigo-500/10 text-indigo-400"><i data-lucide="calendar" class="w-5 h-5"></i></div>
        </div>
        <div class="mt-2 flex items-baseline gap-2">
          <span id="kpi-programs" class="text-2xl font-bold text-white tracking-tight">--</span>
          <span class="text-xs text-slate-400">programmes analysés</span>
        </div>
      </div>

      <div class="bg-slate-900/90 border border-slate-800 rounded-xl p-4 shadow-sm hover:border-slate-700 transition">
        <div class="flex items-center justify-between">
          <span class="text-xs font-semibold uppercase tracking-wider text-slate-400">Articles Expédiés</span>
          <div class="p-2 rounded-lg bg-emerald-500/10 text-emerald-400"><i data-lucide="package-check" class="w-5 h-5"></i></div>
        </div>
        <div class="mt-2 flex items-baseline gap-2">
          <span id="kpi-units" class="text-2xl font-bold text-white tracking-tight">--</span>
          <span class="text-xs text-slate-400">paires / pièces</span>
        </div>
      </div>

      <div class="bg-slate-900/90 border border-slate-800 rounded-xl p-4 shadow-sm hover:border-slate-700 transition">
        <div class="flex items-center justify-between">
          <span class="text-xs font-semibold uppercase tracking-wider text-slate-400">Points de Vente Ciblés</span>
          <div class="p-2 rounded-lg bg-cyan-500/10 text-cyan-400"><i data-lucide="store" class="w-5 h-5"></i></div>
        </div>
        <div class="mt-2 flex items-baseline gap-2">
          <span id="kpi-stores" class="text-2xl font-bold text-white tracking-tight">--</span>
          <span class="text-xs text-slate-400">boutiques & dépôts</span>
        </div>
      </div>

      <div class="bg-slate-900/90 border border-slate-800 rounded-xl p-4 shadow-sm hover:border-slate-700 transition">
        <div class="flex items-center justify-between">
          <span class="text-xs font-semibold uppercase tracking-wider text-slate-400">Flotte & Chauffeurs</span>
          <div class="p-2 rounded-lg bg-amber-500/10 text-amber-400"><i data-lucide="users" class="w-5 h-5"></i></div>
        </div>
        <div class="mt-2 flex items-baseline gap-2">
          <span id="kpi-drivers" class="text-2xl font-bold text-white tracking-tight">--</span>
          <span class="text-xs text-slate-400">conducteurs réguliers</span>
        </div>
      </div>
    </div>

    <!-- CHARTS GRID -->
    <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
      <!-- Monthly Evolution Chart -->
      <div class="lg:col-span-2 bg-slate-900/90 border border-slate-800 rounded-xl p-5 shadow-sm">
        <div class="flex items-center justify-between mb-4">
          <h2 class="text-sm font-semibold text-slate-200 flex items-center gap-2">
            <i data-lucide="trending-up" class="w-4 h-4 text-indigo-400"></i> Évolution des Volumes par Vague & Mois
          </h2>
          <span class="text-xs text-slate-400">Total Unités Expédiées</span>
        </div>
        <div class="h-64 relative">
          <canvas id="chart-timeline"></canvas>
        </div>
      </div>

      <!-- Brands Share Chart -->
      <div class="bg-slate-900/90 border border-slate-800 rounded-xl p-5 shadow-sm">
        <div class="flex items-center justify-between mb-4">
          <h2 class="text-sm font-semibold text-slate-200 flex items-center gap-2">
            <i data-lucide="pie-chart" class="w-4 h-4 text-cyan-400"></i> Répartition des Marques Clés
          </h2>
          <span class="text-xs text-slate-400">En volume</span>
        </div>
        <div class="h-64 relative flex items-center justify-center">
          <canvas id="chart-brands"></canvas>
        </div>
      </div>
    </div>

    <!-- SECONDARY CHARTS GRID -->
    <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
      <!-- Top Stores Chart -->
      <div class="bg-slate-900/90 border border-slate-800 rounded-xl p-5 shadow-sm">
        <div class="flex items-center justify-between mb-4">
          <h2 class="text-sm font-semibold text-slate-200 flex items-center gap-2">
            <i data-lucide="award" class="w-4 h-4 text-amber-400"></i> Top 10 Magasins Récepteurs
          </h2>
          <span class="text-xs text-slate-400">Dotation cumulée</span>
        </div>
        <div class="h-60 relative">
          <canvas id="chart-stores"></canvas>
        </div>
      </div>

      <!-- Drivers Assignments Chart -->
      <div class="bg-slate-900/90 border border-slate-800 rounded-xl p-5 shadow-sm">
        <div class="flex items-center justify-between mb-4">
          <h2 class="text-sm font-semibold text-slate-200 flex items-center gap-2">
            <i data-lucide="truck" class="w-4 h-4 text-emerald-400"></i> Mobilisation des Chauffeurs
          </h2>
          <span class="text-xs text-slate-400">Assignations sur tournées</span>
        </div>
        <div class="h-60 relative">
          <canvas id="chart-drivers"></canvas>
        </div>
      </div>
    </div>

    <!-- FILTER BAR & CONTROLS -->
    <div class="bg-slate-900/90 border border-slate-800 rounded-xl p-4 shadow-sm flex flex-col md:flex-row items-center justify-between gap-4">
      <div class="flex items-center gap-3 w-full md:w-auto">
        <div class="relative w-full md:w-80">
          <i data-lucide="search" class="w-4 h-4 absolute left-3 top-3 text-slate-400"></i>
          <input type="text" id="search-input" oninput="applyFilters()" placeholder="Rechercher boutique, chauffeur, date..." class="w-full bg-slate-950 border border-slate-700/80 rounded-lg pl-9 pr-4 py-2 text-xs text-slate-200 placeholder-slate-500 focus:outline-none focus:border-indigo-500 transition">
        </div>
        <select id="season-filter" onchange="applyFilters()" class="bg-slate-950 border border-slate-700/80 rounded-lg px-3 py-2 text-xs text-slate-300 focus:outline-none focus:border-indigo-500">
          <option value="ALL">Toutes les périodes</option>
          <option value="2026">Année 2026</option>
          <option value="2025">Année 2025</option>
        </select>
      </div>

      <div class="text-xs text-slate-400">
        Affichage de <span id="filtered-count" class="font-bold text-slate-200">0</span> programme(s) sur <span id="total-count" class="font-bold text-slate-200">0</span>
      </div>
    </div>

    <!-- DELIVERY PROGRAMS TABLE -->
    <div class="bg-slate-900/90 border border-slate-800 rounded-xl shadow-sm overflow-hidden">
      <div class="px-5 py-4 border-b border-slate-800 flex items-center justify-between">
        <h2 class="text-sm font-semibold text-white flex items-center gap-2">
          <i data-lucide="list" class="w-4 h-4 text-indigo-400"></i> Registre des Programmes de Livraison
        </h2>
        <span class="text-xs text-slate-400">Cliquez sur un programme pour inspecter le détail des tournées</span>
      </div>

      <div class="overflow-x-auto">
        <table class="w-full text-left text-xs text-slate-300">
          <thead class="bg-slate-950/60 uppercase font-semibold text-slate-400 border-b border-slate-800">
            <tr>
              <th class="px-5 py-3">Date Réception</th>
              <th class="px-5 py-3">Programme / Vague</th>
              <th class="px-5 py-3 text-right">Volume (Unités)</th>
              <th class="px-5 py-3 text-center">Magasins</th>
              <th class="px-5 py-3">Chauffeurs Affectés</th>
              <th class="px-5 py-3">Statut / Remarque</th>
              <th class="px-5 py-3 text-center">Action</th>
            </tr>
          </thead>
          <tbody id="programs-table-body" class="divide-y divide-slate-800/60">
            <!-- Program rows injected via JS -->
          </tbody>
        </table>
      </div>
    </div>

  </main>

  <!-- PROGRAM DETAIL MODAL -->
  <div id="modal-detail" class="fixed inset-0 z-50 bg-slate-950/80 backdrop-blur-sm hidden flex items-center justify-center p-4">
    <div class="bg-slate-900 border border-slate-700/80 rounded-2xl w-full max-w-4xl max-h-[90vh] flex flex-col shadow-2xl overflow-hidden">
      <!-- Modal Header -->
      <div class="px-6 py-4 border-b border-slate-800 flex items-center justify-between bg-slate-950/50">
        <div class="flex items-center space-x-3">
          <div class="p-2 rounded-xl bg-indigo-500/10 text-indigo-400 border border-indigo-500/20">
            <i data-lucide="file-spreadsheet" class="w-5 h-5"></i>
          </div>
          <div>
            <h3 id="modal-title" class="text-base font-bold text-white">Programme de Livraison</h3>
            <p id="modal-sub" class="text-xs text-slate-400"></p>
          </div>
        </div>
        <button onclick="closeModal()" class="text-slate-400 hover:text-white p-1 rounded-lg hover:bg-slate-800 transition">
          <i data-lucide="x" class="w-5 h-5"></i>
        </button>
      </div>

      <!-- Modal Body -->
      <div class="flex-1 overflow-y-auto p-6 space-y-6">
        <!-- Quick KPI Badges -->
        <div class="grid grid-cols-2 sm:grid-cols-4 gap-3 text-xs">
          <div class="bg-slate-950/60 p-3 rounded-xl border border-slate-800/80">
            <span class="text-slate-400 block mb-0.5">Total Unités</span>
            <span id="modal-units" class="text-lg font-bold text-emerald-400">--</span>
          </div>
          <div class="bg-slate-950/60 p-3 rounded-xl border border-slate-800/80">
            <span class="text-slate-400 block mb-0.5">Destinations</span>
            <span id="modal-stores" class="text-lg font-bold text-cyan-400">--</span>
          </div>
          <div class="bg-slate-950/60 p-3 rounded-xl border border-slate-800/80">
            <span class="text-slate-400 block mb-0.5">Chauffeurs</span>
            <span id="modal-drivers" class="text-xs font-semibold text-slate-200 truncate block">--</span>
          </div>
          <div class="bg-slate-950/60 p-3 rounded-xl border border-slate-800/80">
            <span class="text-slate-400 block mb-0.5">Écarts / Retirés</span>
            <span id="modal-discrepancies" class="text-lg font-bold text-amber-400">0</span>
          </div>
        </div>

        <!-- Brands Pills -->
        <div>
          <h4 class="text-xs font-semibold uppercase tracking-wider text-slate-400 mb-2">Marques Incluses</h4>
          <div id="modal-brands" class="flex flex-wrap gap-2 text-xs"></div>
        </div>

        <!-- Store Routes Table -->
        <div>
          <h4 class="text-xs font-semibold uppercase tracking-wider text-slate-400 mb-2">Détail des Tournées & Magasins</h4>
          <div class="border border-slate-800 rounded-xl overflow-hidden">
            <table class="w-full text-left text-xs text-slate-300">
              <thead class="bg-slate-950 uppercase font-semibold text-slate-400 border-b border-slate-800">
                <tr>
                  <th class="px-4 py-2.5">Prio</th>
                  <th class="px-4 py-2.5">Magasin Destinataire</th>
                  <th class="px-4 py-2.5">Wilaya / Zone</th>
                  <th class="px-4 py-2.5 text-right">Quantité</th>
                  <th class="px-4 py-2.5">Chauffeur & Véhicule</th>
                  <th class="px-4 py-2.5">Date</th>
                  <th class="px-4 py-2.5">Retour</th>
                </tr>
              </thead>
              <tbody id="modal-routes-body" class="divide-y divide-slate-800/60"></tbody>
            </table>
          </div>
        </div>
      </div>

      <!-- Modal Footer -->
      <div class="px-6 py-3 border-t border-slate-800 bg-slate-950/50 flex items-center justify-between">
        <span id="modal-filename" class="text-xs text-slate-500 font-mono"></span>
        <button onclick="closeModal()" class="bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs px-4 py-2 rounded-lg transition font-medium">
          Fermer
        </button>
      </div>
    </div>
  </div>

  <script>
    // Embedded Live Data
    const DB = {db_json_str};

    let timelineChart = null;
    let brandsChart = null;
    let storesChart = null;
    let driversChart = null;

    function initDashboard() {{
      lucide.createIcons();
      document.getElementById('last-update').textContent = DB.metadata.last_updated;
      document.getElementById('kpi-programs').textContent = DB.metadata.total_programs;
      document.getElementById('kpi-units').textContent = Number(DB.metadata.total_dispatched_units).toLocaleString();
      document.getElementById('kpi-stores').textContent = DB.metadata.total_unique_stores;
      document.getElementById('kpi-drivers').textContent = Object.keys(DB.drivers).length;

      renderCharts();
      applyFilters();
    }}

    function renderCharts() {{
      // 1. Timeline Chart (Grouped by Month)
      const monthlyData = {{}};
      DB.programs.forEach(p => {{
        const ym = p.year_month || 'Autre';
        monthlyData[ym] = (monthlyData[ym] || 0) + p.total_units;
      }});
      const sortedMonths = Object.keys(monthlyData).sort();

      const ctxTimeline = document.getElementById('chart-timeline').getContext('2d');
      timelineChart = new Chart(ctxTimeline, {{
        type: 'bar',
        data: {{
          labels: sortedMonths,
          datasets: [{{
            label: 'Unités Expédiées',
            data: sortedMonths.map(m => monthlyData[m]),
            backgroundColor: 'rgba(99, 102, 241, 0.75)',
            borderColor: '#6366f1',
            borderRadius: 6,
            borderWidth: 1
          }}]
        }},
        options: {{
          responsive: true,
          maintainAspectRatio: false,
          plugins: {{ legend: {{ display: false }} }},
          scales: {{
            x: {{ grid: {{ display: false }}, ticks: {{ color: '#94a3b8', font: {{ size: 10 }} }} }},
            y: {{ grid: {{ color: '#1e293b' }}, ticks: {{ color: '#94a3b8', font: {{ size: 10 }} }} }}
          }}
        }}
      }});

      // 2. Brands Chart
      const topBrands = Object.entries(DB.brands)
        .filter(([b]) => !b.includes('SACHET') && !b.includes('PAPIER'))
        .sort((a,b) => b[1] - a[1])
        .slice(0, 6);
      
      const ctxBrands = document.getElementById('chart-brands').getContext('2d');
      brandsChart = new Chart(ctxBrands, {{
        type: 'doughnut',
        data: {{
          labels: topBrands.map(x => x[0]),
          datasets: [{{
            data: topBrands.map(x => x[1]),
            backgroundColor: ['#6366f1', '#06b6d4', '#10b981', '#f59e0b', '#ec4899', '#8b5cf6'],
            borderWidth: 0
          }}]
        }},
        options: {{
          responsive: true,
          maintainAspectRatio: false,
          plugins: {{
            legend: {{ position: 'right', labels: {{ color: '#cbd5e1', font: {{ size: 10 }} }} }}
          }},
          cutout: '70%'
        }}
      }});

      // 3. Top Stores Chart
      const topStores = DB.stores.slice(0, 8);
      const ctxStores = document.getElementById('chart-stores').getContext('2d');
      storesChart = new Chart(ctxStores, {{
        type: 'bar',
        indexAxis: 'y',
        data: {{
          labels: topStores.map(s => s.store.replace('FOOTLAND ', 'FL-').replace('SKECHERS ', 'SK-')),
          datasets: [{{
            label: 'Total Unités',
            data: topStores.map(s => s.total_units),
            backgroundColor: '#06b6d4',
            borderRadius: 4
          }}]
        }},
        options: {{
          responsive: true,
          maintainAspectRatio: false,
          plugins: {{ legend: {{ display: false }} }},
          scales: {{
            x: {{ grid: {{ color: '#1e293b' }}, ticks: {{ color: '#94a3b8', font: {{ size: 10 }} }} }},
            y: {{ grid: {{ display: false }}, ticks: {{ color: '#94a3b8', font: {{ size: 10 }} }} }}
          }}
        }}
      }});

      // 4. Drivers Chart
      const topDrivers = Object.entries(DB.drivers)
        .filter(([d]) => d !== 'CLIENT' && d.length > 2)
        .sort((a,b) => b[1] - a[1])
        .slice(0, 6);

      const ctxDrivers = document.getElementById('chart-drivers').getContext('2d');
      driversChart = new Chart(ctxDrivers, {{
        type: 'doughnut',
        data: {{
          labels: topDrivers.map(x => x[0]),
          datasets: [{{
            data: topDrivers.map(x => x[1]),
            backgroundColor: ['#10b981', '#6366f1', '#f59e0b', '#06b6d4', '#8b5cf6', '#64748b'],
            borderWidth: 0
          }}]
        }},
        options: {{
          responsive: true,
          maintainAspectRatio: false,
          plugins: {{
            legend: {{ position: 'right', labels: {{ color: '#cbd5e1', font: {{ size: 10 }} }} }}
          }},
          cutout: '70%'
        }}
      }});
    }}

    function applyFilters() {{
      const query = document.getElementById('search-input').value.toLowerCase();
      const season = document.getElementById('season-filter').value;

      const filtered = DB.programs.filter(p => {{
        const matchesSeason = (season === 'ALL') || (p.received_date && p.received_date.startsWith(season));
        const matchesQuery = !query || 
          p.display_title.toLowerCase().includes(query) ||
          p.filename.toLowerCase().includes(query) ||
          p.drivers.some(d => d.toLowerCase().includes(query)) ||
          p.destinations.some(d => d.store.toLowerCase().includes(query));
        return matchesSeason && matchesQuery;
      }});

      document.getElementById('filtered-count').textContent = filtered.length;
      document.getElementById('total-count').textContent = DB.programs.length;

      const tbody = document.getElementById('programs-table-body');
      tbody.innerHTML = '';

      filtered.forEach(p => {{
        const tr = document.createElement('tr');
        tr.className = 'hover:bg-slate-800/40 cursor-pointer transition';
        tr.onclick = () => openModal(p.id);

        let badgeStatus = '<span class="px-2 py-0.5 rounded-full text-[10px] font-semibold bg-emerald-500/10 text-emerald-400 border border-emerald-500/20">Programme Validé</span>';
        if (p.is_modified) {{
          badgeStatus = '<span class="px-2 py-0.5 rounded-full text-[10px] font-semibold bg-amber-500/10 text-amber-400 border border-amber-500/20">Version Modifiée</span>';
        }}
        if (p.discrepancies_count > 0) {{
          badgeStatus += ` <span class="px-2 py-0.5 rounded-full text-[10px] font-semibold bg-red-500/10 text-red-400 border border-red-500/20">${{p.discrepancies_count}} écarts</span>`;
        }}

        tr.innerHTML = `
          <td class="px-5 py-3.5 font-mono text-slate-400">${{p.received_date}}</td>
          <td class="px-5 py-3.5 font-medium text-white">${{p.display_title}}</td>
          <td class="px-5 py-3.5 text-right font-bold text-emerald-400">${{Number(p.total_units).toLocaleString()}}</td>
          <td class="px-5 py-3.5 text-center"><span class="px-2 py-0.5 rounded-md bg-slate-800 text-slate-200 font-semibold">${{p.destinations_count}}</span></td>
          <td class="px-5 py-3.5 text-slate-300 truncate max-w-xs">${{p.drivers.slice(0, 2).join(', ') || 'N/A'}}${{p.drivers.length > 2 ? '...' : ''}}</td>
          <td class="px-5 py-3.5">${{badgeStatus}}</td>
          <td class="px-5 py-3.5 text-center">
            <button class="p-1.5 rounded-lg bg-indigo-500/10 text-indigo-400 hover:bg-indigo-500/20 transition">
              <i data-lucide="eye" class="w-4 h-4"></i>
            </button>
          </td>
        `;
        tbody.appendChild(tr);
      }});

      lucide.createIcons();
    }}

    function openModal(progId) {{
      const p = DB.programs.find(x => x.id === progId);
      if (!p) return;

      document.getElementById('modal-title').textContent = p.display_title;
      document.getElementById('modal-sub').textContent = `Reçu le ${{p.received_date}} | Fichier : ${{p.filename}}`;
      document.getElementById('modal-units').textContent = Number(p.total_units).toLocaleString();
      document.getElementById('modal-stores').textContent = p.destinations_count;
      document.getElementById('modal-drivers').textContent = p.drivers.join(', ') || 'Non spécifié';
      document.getElementById('modal-discrepancies').textContent = p.discrepancies_count;
      document.getElementById('modal-filename').textContent = p.filename;

      // Brands pills
      const brandsContainer = document.getElementById('modal-brands');
      brandsContainer.innerHTML = '';
      if (Object.keys(p.brands).length > 0) {{
        Object.entries(p.brands).forEach(([b, q]) => {{
          const span = document.createElement('span');
          span.className = 'px-2.5 py-1 rounded-lg bg-indigo-500/10 text-indigo-300 border border-indigo-500/20 font-medium';
          span.textContent = `${{b}} : ${{Number(q).toLocaleString()}} u.`;
          brandsContainer.appendChild(span);
        }});
      }} else {{
        brandsContainer.innerHTML = '<span class="text-slate-500 italic">Aucun détail marque spécifié</span>';
      }}

      // Routes table
      const tbody = document.getElementById('modal-routes-body');
      tbody.innerHTML = '';
      p.destinations.forEach(d => {{
        const tr = document.createElement('tr');
        tr.className = 'hover:bg-slate-800/30';
        tr.innerHTML = `
          <td class="px-4 py-2 font-mono text-slate-400">${{d.priority || '-'}}</td>
          <td class="px-4 py-2 font-semibold text-white">${{d.store}}</td>
          <td class="px-4 py-2 text-slate-400">${{d.wilaya || ''}} <span class="text-[10px] text-slate-500">(${{d.region || 'Centre'}})</span></td>
          <td class="px-4 py-2 text-right font-bold text-emerald-400">${{Number(d.quantity).toLocaleString()}}</td>
          <td class="px-4 py-2 text-slate-300">${{d.driver || '-'}} <span class="text-[10px] text-slate-500">(${{d.vehicle || '-'}})</span></td>
          <td class="px-4 py-2 font-mono text-slate-400">${{d.delivery_date || '-'}}</td>
          <td class="px-4 py-2 text-slate-400">${{d.retour ? '<span class="text-amber-400 font-semibold">' + d.retour + '</span>' : '-'}}</td>
        `;
        tbody.appendChild(tr);
      }});

      document.getElementById('modal-detail').classList.remove('hidden');
      lucide.createIcons();
    }}

    function closeModal() {{
      document.getElementById('modal-detail').classList.add('hidden');
    }}

    function refreshDashboard() {{
      location.reload();
    }}

    window.onload = initDashboard;
  </script>
</body>
</html>
"""

    with open(DASHBOARD_PATH, "w", encoding="utf-8") as fp:
        fp.write(html_template)

    print(f"[+] Interactive Dashboard generated successfully at: {DASHBOARD_PATH}")
    return DASHBOARD_PATH


# -------------------------------------------------------------------------
# CALENDAR APP (index.html) SYNCHRONIZATION MODULE
# -------------------------------------------------------------------------
def update_calendar_app(db_data, index_path=None):
    """
    Synchronizes index.html by injecting the live DELIVERY_DATA extracted
    from the parsed delivery programs in delivery_db.json.
    """
    if index_path is None:
        index_path = os.path.join(WORKSPACE_DIR, "index.html")

    if not os.path.exists(index_path):
        print(f"[-] {index_path} not found. Skipping calendar sync.")
        return None

    today_str = datetime.date.today().strftime("%Y-%m-%d")
    all_routes = []
    seen_route_keys = set()
    route_counter = 1

    for p in db_data.get("programs", []):
        # Identify brand label
        b_list = [b for b in p.get("brands", {}).keys() if "SACHET" not in b and "PAPIER" not in b]
        brand_str = " / ".join(b_list[:3]) if b_list else "Multi-marques"

        # Determine target delivery date fallback
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
            if re.match(r'^\d{4}-\d{2}-\d{2}$', raw_deliv_date):
                deliv_date = raw_deliv_date
            else:
                deliv_date = fallback_date

            qty = d.get("quantity", 0)
            store_clean = (d.get("store") or "").strip().upper()
            if not store_clean or store_clean.startswith("TOTAL"):
                continue

            # Strict deduplication: same date, same destination, same quantity
            route_key = (deliv_date, store_clean, qty)
            if route_key in seen_route_keys:
                continue
            seen_route_keys.add(route_key)

            # Compute Status
            if deliv_date < today_str:
                status = "non_livre"
            elif deliv_date == today_str:
                status = "non_livre"
            else:
                status = "non_livre"

            driver = re.sub(r"\s+", " ", d.get("driver") or "Non spécifié").strip().upper()
            vehicle = d.get("vehicle") or "Camion Porteur"
            if not vehicle or vehicle == "/":
                vehicle = "Camion Porteur"

            prio = d.get("priority") or 1
            try:
                prio = int(prio)
            except Exception:
                prio = 1

            notes = d.get("remark") or d.get("retour") or f"Programme {p_title}"
            time_str = f"Tournée {p_title}"

            all_routes.append({
                "id": f"DEL-{route_counter:04d}",
                "date": deliv_date,
                "driver": driver,
                "vehicle": vehicle,
                "destination": d["store"],
                "wilaya": d.get("wilaya") or "Alger",
                "region": d.get("region") or "Centre",
                "quantity": qty,
                "colis": max(1, int(round(qty / 20))) if qty > 0 else 1,
                "status": "Non Livr\u00e9",
                "completed": False,
                "priority": prio,
                "time": time_str,
                "notes": notes,
                "brand": brand_str
            })
            route_counter += 1

    # Read existing index.html
    with open(index_path, "r", encoding="utf-8") as f:
        html_content = f.read()

    # Always default calendar to today (current month & day)
    active_date = today_str
    y_sel, m_sel, d_sel = [int(x) for x in active_date.split("-")]

    json_routes = json.dumps(all_routes, ensure_ascii=False, indent=6)

    # Replacement 1: DELIVERY_DATA array via safe string slicing
    var_loc_start = html_content.find("const LOCAL_DELIVERY_DATA =")
    if var_loc_start != -1:
        bracket_start = html_content.find("[", var_loc_start)
        next_var = html_content.find("let DELIVERY_DATA =", bracket_start)
        if next_var != -1:
            semicolon_pos = html_content.rfind("];", bracket_start, next_var)
            if bracket_start != -1 and semicolon_pos != -1:
                html_content = html_content[:bracket_start] + json_routes + html_content[semicolon_pos + 1:]

    var_start = html_content.find("const DELIVERY_DATA =")
    if var_start != -1:
        bracket_start = html_content.find("[", var_start)
        state_start = html_content.find("// 2. STATE MANAGEMENT")
        if state_start != -1:
            semicolon_pos = html_content.rfind("];", bracket_start, state_start)
            if bracket_start != -1 and semicolon_pos != -1:
                html_content = html_content[:bracket_start] + json_routes + html_content[semicolon_pos + 1:]

    # Replacement 2: Initial view date
    pattern_state = r'let currentViewDate\s*=\s*new Date\([^)]+\);(\s*)let selectedDateStr\s*=\s*"[^"]+";'
    replacement_state = f'let currentViewDate = new Date({y_sel}, {m_sel - 1}, {d_sel});\\1let selectedDateStr = "{active_date}";'
    new_html = re.sub(pattern_state, replacement_state, html_content)

    with open(index_path, "w", encoding="utf-8") as f:
        f.write(new_html)

    print(f"[+] Calendar App (index.html) synchronized with {len(all_routes)} live delivery routes (Active Date: {active_date}).")
    return index_path


# -------------------------------------------------------------------------
# CLI ENTRY POINT
# -------------------------------------------------------------------------
def main():
    import argparse
    parser = argparse.ArgumentParser(description="Tradifoot Delivery Tracker & Intelligence Engine")
    parser.add_argument("--sync", action="store_true", help="Sync latest programs from Outlook")
    parser.add_argument("--full", action="store_true", help="Perform full historical Outlook sync")
    parser.add_argument("--open", action="store_true", help="Open dashboard in default browser")
    parser.add_argument("--stats", action="store_true", help="Display summary terminal stats")

    args = parser.parse_args()

    # Step 1: Sync from Outlook if requested
    if args.sync or args.full:
        sync_from_outlook(full_scan=args.full)

    # Step 2: Build JSON DB
    db_data = build_database()

    # Step 3: Generate Interactive Dashboard
    dash_path = generate_dashboard(db_data)

    # Step 4: Synchronize Calendar App (index.html)
    cal_path = update_calendar_app(db_data)

    # Step 5: Display stats if requested
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

    # Step 6: Open in browser
    if args.open:
        print(f"[*] Opening dashboard in default web browser: {dash_path}")
        webbrowser.open(dash_path)


if __name__ == "__main__":
    main()
