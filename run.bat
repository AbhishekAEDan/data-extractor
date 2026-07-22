@echo off
rem Brent & Co. Data Extractor -- author: AbhishekAEDan
title Brent ^& Co. Data Extractor
color 0B

rem No admin needed: pip, winget and the Ollama installer all work per-user.
rem (Self-elevation removed -- UAC + SmartScreen killed zip-downloaded copies.)

cd /d "%~dp0"

where python >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo   Python was not found on this computer.
    echo   Install it from https://www.python.org/downloads/
    echo   IMPORTANT: tick "Add python.exe to PATH" during install.
    echo.
    pause
    exit /b 1
)

python main.py %*
pause
