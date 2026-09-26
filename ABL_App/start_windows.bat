@echo off
cd /d "%~dp0"
title ABL Report Generator

echo.
echo  ╔══════════════════════════════════════╗
echo  ║    Radiometer ABL Report Generator  ║
echo  ║    Pathology Queensland              ║
echo  ╚══════════════════════════════════════╝
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo  [!] Python not found.
    echo      Please download Python from https://python.org
    echo      Make sure to tick "Add Python to PATH" during install.
    pause
    exit /b 1
)

echo  [1/3] Checking Python packages...
:: Only reach the internet on FIRST-TIME setup; normal runs are fully offline.
python -c "import flask, docx, openpyxl, pandas, matplotlib" >nul 2>&1
if errorlevel 1 (
    echo        First-time setup: installing packages ^(internet required once^)...
    pip install flask python-docx openpyxl pandas matplotlib --quiet
) else (
    echo        All packages present - no internet needed.
)

echo  [2/3] Starting app server...
echo  [3/3] Opening browser...
echo.
echo  ─────────────────────────────────────────
echo  App is running at: http://localhost:5758
echo  Close this window to stop the app.
echo  ─────────────────────────────────────────
echo.

python app.py
pause
