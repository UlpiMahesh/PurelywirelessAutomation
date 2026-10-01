@echo off
cd /d "%~dp0"

echo Starting T-Mobile Tracking Automation...
echo.

"..\.venv\Scripts\python.exe" main.py

if errorlevel 1 (
    echo.
    echo ========================================
    echo Automation failed.
    echo ========================================
)

pause