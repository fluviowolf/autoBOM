@echo off
setlocal
cd /d "%~dp0"
title Implant BOM Generator

set "PY=.venv\Scripts\python.exe"
if not exist "%PY%" (
    where py >nul 2>nul && (set "PY=py") || (set "PY=python")
)

"%PY%" -c "import streamlit, pandas, openpyxl" >nul 2>nul
if errorlevel 1 (
    echo First-time setup: installing required packages...
    "%PY%" -m pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo Setup failed. Please contact support.
        pause
        exit /b 1
    )
)

echo Starting Implant BOM Generator. Your browser will open shortly.
echo Keep this window open while using the app. Close it to stop the app.
"%PY%" -m streamlit run app.py --server.address 127.0.0.1 --browser.gatherUsageStats false

pause
