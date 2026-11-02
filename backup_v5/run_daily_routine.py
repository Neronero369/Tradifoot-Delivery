#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tradifoot 19h Daily Delivery Routine
Runs automatically every day at 19:00 to check for new delivery programs in Outlook Classic,
download them, update the database and dashboard, and alert the user.
"""

import os
import sys
import datetime
import subprocess
import json

WORKSPACE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_DIR = os.path.join(WORKSPACE_DIR, "data", "logs")
os.makedirs(LOG_DIR, exist_ok=True)
LOG_FILE = os.path.join(LOG_DIR, "routine.log")

def log(msg):
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    entry = f"[{now_str}] {msg}"
    print(entry)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(entry + "\n")

def show_notification(title, message):
    """Triggers a Windows toast notification via PowerShell."""
    try:
        ps_cmd = f"""
        [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
        [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null
        $template = @"
<toast>
    <visual>
        <binding template="ToastGeneric">
            <text>{title}</text>
            <text>{message}</text>
        </binding>
    </visual>
</toast>
"@
        $xml = New-Object Windows.Data.Xml.Dom.XmlDocument
        $xml.LoadXml($template)
        $toast = [Windows.UI.Notifications.ToastNotification]::new($xml)
        [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("Tradifoot Logistics").Show($toast)
        """
        subprocess.run(["powershell", "-NoProfile", "-Command", ps_cmd], capture_output=True, timeout=5)
    except Exception as e:
        # Fallback to simple msg
        pass

def main():
    log("="*60)
    log("DÉBUT DE LA ROUTINE QUOTIDIENNE (19H00) - TRADIFOOT LOGISTICS")
    log("Vérification des nouveaux programmes de livraison dans Outlook...")

    import tradifoot_tracker

    # Step 1: Sync from Outlook (check last 2 days)
    new_files_count = tradifoot_tracker.sync_from_outlook(days_back=2, full_scan=False)

    if new_files_count > 0:
        log(f"[+] Succès : {new_files_count} nouveau(x) fichier(s) de livraison détecté(s) !")
        
        # Step 2: Build database
        db_data = tradifoot_tracker.build_database()
        
        # Step 3: Regenerate dashboard and update interactive Calendar App (index.html)
        dash_path = tradifoot_tracker.generate_dashboard(db_data)
        cal_path = tradifoot_tracker.update_calendar_app(db_data)
        
        latest_prog = db_data["programs"][0] if db_data["programs"] else None
        units_str = f"{latest_prog['total_units']:,}" if latest_prog else "N/A"
        stores_count = latest_prog['destinations_count'] if latest_prog else "N/A"

        log(f"[+] Base, Dashboard ({dash_path}) et Calendrier ({cal_path}) actualisés avec succès !")
        log(f"[+] Dernier programme : {latest_prog['display_title']} ({stores_count} magasins, {units_str} unités)")
        
        # Step 4: Alert User via Windows Toast
        show_notification(
            "Tradifoot Logistics - 19h",
            f"Nouveau programme détecté ! {stores_count} magasins, {units_str} unités expédiées."
        )
    else:
        log("[i] Aucun nouveau programme de livraison reçu aujourd'hui.")
        # Re-verify database, dashboard, and calendar app to ensure up to date
        db_data = tradifoot_tracker.build_database()
        tradifoot_tracker.generate_dashboard(db_data)
        cal_path = tradifoot_tracker.update_calendar_app(db_data)
        log("[i] Base de données, dashboard et calendrier (index.html) synchronisés.")

    log("FIN DE LA ROUTINE QUOTIDIENNE.")
    log("="*60 + "\n")

if __name__ == "__main__":
    main()
