@echo off
title Tradifoot Delivery Tracker Sync
echo ===================================================
echo   TRADIFOOT DELIVERY TRACKER - SYNCING OUTLOOK
echo ===================================================
echo [1/2] Connecting to Outlook Classic & syncing latest delivery programs...
python "%~dp0tradifoot_tracker.py" --sync --open
if %ERRORLEVEL% NEQ 0 (
    echo [!] Python sync encountered an error. Opening existing dashboard...
    start "" "%~dp0index.html"
)
exit
