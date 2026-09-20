@echo off
:: ============================================================
::  Excel AI Assistant - Bridge Server Launcher
::  Starts the local bridge server (HTTP) + OpenCode serve for the Excel add-in.
::  Run this BEFORE opening Excel.
:: ============================================================
title Excel AI Assistant - Bridge Server

cd /d "%~dp0"

echo.
echo  ============================================================
echo   Excel AI Assistant - Bridge Server
echo  ============================================================
echo.

:: Check Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python not found. Please install Python 3.8+ and add to PATH.
    pause
    exit /b 1
)

:: Install required Python packages if not already installed
echo  Checking dependencies...
python -m pip install cryptography --quiet --disable-pip-version-check 2>nul

echo.
echo  Starting bridge server on http://localhost:3000 ...
echo  (OpenCode serve auto-starts on 127.0.0.1:4096 when needed.)
echo  Press Ctrl+C to stop.
echo.

python server\bridge_server.py

pause
