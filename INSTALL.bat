@echo off
title Railway File Organiser — Setup
color 0A
echo.
echo  ============================================================
echo    Railway File Organiser — Setup and Install
echo  ============================================================
echo.

REM ── Step 1: Check Python ─────────────────────────────────────────
echo  [1/4] Checking Python installation...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo  ERROR: Python is not installed or not in PATH.
    echo.
    echo  Please install Python 3.10 or newer from:
    echo  https://www.python.org/downloads/
    echo.
    echo  IMPORTANT: During install, tick the box:
    echo  "Add Python to PATH"
    echo.
    pause
    exit /b 1
)
for /f "tokens=*" %%v in ('python --version') do echo   Found: %%v
echo.

REM ── Step 2: Install dependencies ─────────────────────────────────
echo  [2/4] Installing required packages...
echo   (this may take a minute on first run)
echo.
python -m pip install --upgrade pip --quiet
python -m pip install -r "%~dp0requirements.txt" --quiet
if %errorlevel% neq 0 (
    echo.
    echo  ERROR: Failed to install packages.
    echo  Please check your internet connection and try again.
    pause
    exit /b 1
)
echo   All packages installed successfully.
echo.

REM ── Step 3: Create Desktop shortcut ──────────────────────────────
echo  [3/4] Creating Desktop shortcut...
powershell -ExecutionPolicy Bypass -Command ^
  "$s=(New-Object -ComObject WScript.Shell).CreateShortcut([Environment]::GetFolderPath('Desktop')+'\Railway File Organiser.lnk');$s.TargetPath='%~dp0START.bat';$s.WorkingDirectory='%~dp0';$s.Description='Railway File Organiser';$s.Save()"
echo   Desktop shortcut created.
echo.

REM ── Step 4: Reset config to dynamic defaults ──────────────────────
echo  [4/4] Resetting configuration for this machine...
python -c "
import json, os
from pathlib import Path
cfg_path = r'%~dp0config.json'
cfg = {
    'watch_folder': str(Path.home() / 'Downloads'),
    'base_folder': str(Path.home() / 'Documents' / 'Railway Files'),
    'excel_log_filename': 'Railway_Files_Log.xlsx',
    'auto_confirm_timeout_seconds': 60,
    'open_excel_after_update': True,
    'stable_wait_seconds': 2,
    'auto_start_with_windows': False,
    'handle_duplicates': 'rename',
    'show_notifications': True,
    'min_file_size_bytes': 100,
    'theme': 'dark'
}
with open(cfg_path, 'w') as f:
    json.dump(cfg, f, indent=4)
print('   Config reset for this machine.')
"
echo.

echo  ============================================================
echo   Installation Complete!
echo  ============================================================
echo.
echo   A shortcut has been placed on your Desktop.
echo   Double-click "Railway File Organiser" to launch the app.
echo.
echo   Launching app now...
echo.
timeout /t 2 /nobreak >nul
start "" "%~dp0START.bat"
pause
