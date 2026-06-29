@echo off
title Railway File Organiser
cd /d "%~dp0"
echo [START] Launching Railway File Organiser...
python app.py
echo.
echo [INFO] App closed with code: %errorlevel%
echo.
echo If the app did not open, please read any errors above.
pause
