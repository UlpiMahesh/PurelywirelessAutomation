@echo off
echo ============================================
echo   Truely Wireless Automation
echo ============================================
echo.

:: Check venv exists — if not, tell them to run install first
if not exist ".venv\Scripts\activate.bat" (
    echo ERROR: Setup not complete.
    echo Please run install.bat first.
    pause
    exit /b 1
)

echo Starting app...
call .venv\Scripts\activate.bat
start "" http://localhost:8501
streamlit run app.py

pause
