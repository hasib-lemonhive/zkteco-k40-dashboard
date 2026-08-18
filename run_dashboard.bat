@echo off
title ZKTeco Attendance Dashboard & Live Listener

echo ===================================================
echo   ZKTeco Attendance Dashboard & Live Listener
echo ===================================================
echo.

cd /d "%~dp0"

IF NOT EXIST ".venv" (
    echo [1/3] Creating virtual environment...
    python -m venv .venv
    IF %ERRORLEVEL% NEQ 0 (
        echo.
        echo [ERROR] Python is not installed or not added to PATH!
        echo Please install Python from https://www.python.org/ and check "Add Python to PATH".
        echo.
        pause
        exit /b %ERRORLEVEL%
    )
)

echo [2/3] Checking requirements...
call .venv\Scripts\pip.exe install -r requirements.txt --quiet

echo [3/3] Starting Real-Time Live Listener...
start "ZKTeco Live Listener" /min .venv\Scripts\python.exe sync_k40.py

echo.
echo Starting Attendance Dashboard...
echo Local Access:    http://localhost:5000
echo.

start "" "http://localhost:5000"
call .venv\Scripts\python.exe dashboard\app.py

pause
