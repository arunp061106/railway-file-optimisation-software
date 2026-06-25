@echo off
title Railway File Organiser — Setup
color 0B
echo.
echo  =========================================================
echo    Railway File Organiser — First Time Setup
echo  =========================================================
echo.
echo  Checking Python installation...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo  [ERROR] Python is not installed or not in PATH.
    echo  Please install Python 3.9+ from https://python.org
    echo  and make sure to check "Add Python to PATH" during install.
    pause
    exit /b 1
)
python --version
echo.
echo  Installing required libraries...
echo  (This only needs to be done once. No internet needed after this.)
echo.
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo.
    echo  [ERROR] Could not install some libraries.
    echo  Make sure you are connected to the internet for first-time setup.
    pause
    exit /b 1
)
echo.
echo  =========================================================
echo    Setup complete!
echo    Run START.bat to launch the Railway File Organiser.
echo  =========================================================
echo.
pause
