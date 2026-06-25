@echo off
title Building Railway File Organiser...
color 0A
echo.
echo  Building standalone Railway File Organiser.exe ...
echo  Please wait, this takes about 1-2 minutes.
echo.
pyinstaller ^
    --onefile ^
    --windowed ^
    --name "Railway File Organiser" ^
    --add-data "categories.json;." ^
    --add-data "config.json;." ^
    app.py
if %errorlevel% neq 0 (
    echo.
    echo  [ERROR] Build failed.
    pause
    exit /b 1
)
echo.
echo  =========================================================
echo    SUCCESS!
echo    Your app is ready: dist\Railway File Organiser.exe
echo    
echo    To deploy on any PC:
echo      Copy the .exe + categories.json + config.json
echo      to the same folder. Double-click to run!
echo  =========================================================
echo.
pause
