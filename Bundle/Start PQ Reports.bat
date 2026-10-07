@echo off
title Pathology Queensland - Operator Report Generator
cd /d "%~dp0"

if not exist "%~dp0python\python.exe" (
    echo  [!] Portable Python not found next to this launcher.
    echo      Keep this file, the "python", "portal", "iSTAT_App" and "ABL_App"
    echo      folders together.
    pause
    exit /b 1
)

rem Keep matplotlib's font cache on this drive, not on the host PC
set "MPLCONFIGDIR=%~dp0python\mplcache"
if not exist "%MPLCONFIGDIR%" mkdir "%MPLCONFIGDIR%"

echo.
echo   Pathology Queensland - Operator Report Generator (portable)
echo   Runs entirely from this drive: no installation, no internet.
echo.
echo   Selection page:  http://localhost:5750
echo   The FIRST launch on a new computer can take up to 30 seconds -
echo   the browser opens automatically as soon as it is ready.
echo   Close this window to stop everything.
echo.

rem Open the browser once the portal is accepting connections
start "" /min "%~dp0python\python.exe" "%~dp0open_browser_when_ready.py"

"%~dp0python\python.exe" "%~dp0portal\portal.py"
pause
