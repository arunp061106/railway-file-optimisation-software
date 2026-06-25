$AppName    = "Railway File Organiser"
$InstallDir = "$env:LOCALAPPDATA\RailwayFileOrganiser"
$ExeName    = "Railway File Organiser.exe"
$ExeSrc     = "$PSScriptRoot\dist\$ExeName"
$RegKey     = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\RailwayFileOrganiser"
$StartupKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"

Write-Host ""
Write-Host "Installing $AppName..." -ForegroundColor Cyan
Write-Host ""

if (-not (Test-Path $ExeSrc)) {
    Write-Host "ERROR: App file not found at: $ExeSrc" -ForegroundColor Red
    Write-Host "Make sure dist\Railway File Organiser.exe exists." -ForegroundColor Yellow
    Start-Sleep -Seconds 5
    exit 1
}

# Step 1 - Create install folder and copy files
Write-Host "  [1/5] Copying files to $InstallDir ..." -ForegroundColor Yellow
New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
Copy-Item $ExeSrc "$InstallDir\$ExeName" -Force
if (Test-Path "$PSScriptRoot\categories.json") { Copy-Item "$PSScriptRoot\categories.json" "$InstallDir\categories.json" -Force }
if (Test-Path "$PSScriptRoot\config.json")     { Copy-Item "$PSScriptRoot\config.json"     "$InstallDir\config.json"     -Force }
Write-Host "       Done." -ForegroundColor Green

# Step 2 - Desktop shortcut
Write-Host "  [2/5] Creating Desktop shortcut ..." -ForegroundColor Yellow
$shell     = New-Object -ComObject WScript.Shell
$deskLnk   = [System.IO.Path]::Combine([Environment]::GetFolderPath("Desktop"), "$AppName.lnk")
$sc        = $shell.CreateShortcut($deskLnk)
$sc.TargetPath       = "$InstallDir\$ExeName"
$sc.WorkingDirectory = $InstallDir
$sc.Description      = "Railway File Organiser"
$sc.Save()
Write-Host "       Done. Icon on Desktop." -ForegroundColor Green

# Step 3 - Start Menu shortcut
Write-Host "  [3/5] Adding to Start Menu ..." -ForegroundColor Yellow
$smDir  = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Railway"
New-Item -ItemType Directory -Path $smDir -Force | Out-Null
$smLnk  = "$smDir\$AppName.lnk"
$sm     = $shell.CreateShortcut($smLnk)
$sm.TargetPath       = "$InstallDir\$ExeName"
$sm.WorkingDirectory = $InstallDir
$sm.Description      = "Railway File Organiser"
$sm.Save()
Write-Host "       Done. Start Menu > Railway > $AppName" -ForegroundColor Green

# Step 4 - Auto-start with Windows
Write-Host "  [4/5] Setting up auto-start on Windows login ..." -ForegroundColor Yellow
Set-ItemProperty -Path $StartupKey -Name $AppName -Value "`"$InstallDir\$ExeName`""
Write-Host "       Done. App will start automatically with Windows." -ForegroundColor Green

# Step 5 - Register in Apps and Features
Write-Host "  [5/5] Registering in Apps and Features ..." -ForegroundColor Yellow
New-Item -Path $RegKey -Force | Out-Null
Set-ItemProperty -Path $RegKey -Name "DisplayName"     -Value $AppName
Set-ItemProperty -Path $RegKey -Name "Publisher"       -Value "Indian Railways"
Set-ItemProperty -Path $RegKey -Name "DisplayVersion"  -Value "1.0.0"
Set-ItemProperty -Path $RegKey -Name "InstallLocation" -Value $InstallDir
Set-ItemProperty -Path $RegKey -Name "UninstallString" -Value "powershell -ExecutionPolicy Bypass -File `"$PSScriptRoot\UNINSTALL.ps1`""
Set-ItemProperty -Path $RegKey -Name "NoModify"        -Value 1 -Type DWord
Write-Host "       Done. Visible in Settings > Apps." -ForegroundColor Green

Write-Host ""
Write-Host "=============================================" -ForegroundColor Green
Write-Host "  Installation Complete!" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Launching Railway File Organiser now..." -ForegroundColor Cyan
Write-Host ""
Start-Sleep -Seconds 2
Start-Process "$InstallDir\$ExeName"
Start-Sleep -Seconds 3
