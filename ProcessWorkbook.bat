@echo off
setlocal

if "%~1"=="" (
    echo Drag and drop an Excel (.xlsx) file onto this script to process it.
    echo Or run it from a command prompt with the file path as an argument.
    pause
    goto :eof
)

set "FILE=%~1"

REM Change to repo directory
cd /d "S:\Shared With Me\Water for Inhalation\BIAutomations\sales-lot-tool"

echo Processing "%FILE%" ...
python -m sales_tool.processor "%FILE%"

echo.
echo Done. Press any key to close this window.
pause >nul

endlocal
