@echo off
setlocal

REM Change to repo directory
cd /d "S:\Shared With Me\Water for Inhalation\BIAutomations\sales-lot-tool"

echo.
echo Starting Sales & Inventory Lot Tool web app...
echo Keep this window open while you use the tool.
echo.

REM Use python -m to avoid needing 'streamlit' on PATH
python -m streamlit run app.py

endlocal
