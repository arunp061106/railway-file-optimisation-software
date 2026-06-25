@echo off
title Railway File Organiser
cd /d "%~dp0"
python app.py
if %errorlevel% neq 0 (
    echo.
    echo  ERROR: Could not start the app.
    echo  Please run INSTALL.bat first to set up Python dependencies.
    pause
)
