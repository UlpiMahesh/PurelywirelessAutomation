@echo off
REM ---------------------------------------------------------------
REM IMEI loader - launched by Windows Task Scheduler
REM Put this file inside the imei_loader folder.
REM ---------------------------------------------------------------

REM Move to the PROJECT ROOT (one level above imei_loader)
cd /d "%~dp0.."

REM If you use a virtual environment, point at it here:
REM set PYTHON=.venv\Scripts\python.exe
set PYTHON=python

%PYTHON% -m imei_loader.load_imei
set RC=%ERRORLEVEL%

if %RC% NEQ 0 (
    echo Loader exited with code %RC%
)
exit /b %RC%
