@echo off
echo ============================================
echo   Truely Wireless Automation - Setup
echo ============================================
echo.

:: Check Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH.
    echo Please install Python from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)

echo [1/4] Python found. Creating virtual environment...
python -m venv .venv

echo [2/4] Activating virtual environment...
call .venv\Scripts\activate.bat

echo [3/4] Installing Python dependencies...
pip install -r requirements.txt

echo [4/4] Installing Playwright browser (Chromium)...
playwright install chromium

echo.
echo ============================================
echo   Setup Complete! 
echo   Run run.bat to launch the app.
echo ============================================
pause
