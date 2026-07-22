@echo off
rem Brent & Co. Data Extractor -- author: AbhishekAEDan
title Brent ^& Co. Data Extractor
color 0B

rem ---- self-elevate to admin (needed for auto-installs), keep dragged paths ----
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Requesting administrator rights...
    if "%~1"=="" (
        powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    ) else (
        powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -ArgumentList '%*' -Verb RunAs"
    )
    exit /b
)

cd /d "%~dp0"
python main.py %*
pause
