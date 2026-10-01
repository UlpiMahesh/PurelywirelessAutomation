@echo off
cd /d "%~dp0"
python src\main.py
if errorlevel 1 (echo.&echo Automation failed. Read the error above.&pause&exit /b 1)
echo.&echo Report created successfully in output folder.&pause
