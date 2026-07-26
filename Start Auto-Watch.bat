@echo off
cd /d "%~dp0"
echo Starting auto-watch mode. Keep this window open.
echo Drop or replace your Excel file in the "incoming" folder any time.
echo Press Ctrl+C to stop.
python watch_folder.py
pause
