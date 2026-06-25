$AppName    = "Railway File Organiser"
$InstallDir = "$env:LOCALAPPDATA\RailwayFileOrganiser"
$RegKey     = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\RailwayFileOrganiser"
$StartupKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"

Write-Host ""
Write-Host "Uninstalling $AppName..." -ForegroundColor Yellow

Get-Process -Name "Railway File Organiser" -ErrorAction SilentlyContinue | Stop-Process -Force

if (Test-Path $InstallDir)  { Remove-Item $InstallDir -Recurse -Force }
if (Test-Path $RegKey)      { Remove-Item $RegKey -Recurse -Force }

$deskLnk = [System.IO.Path]::Combine([Environment]::GetFolderPath("Desktop"), "$AppName.lnk")
if (Test-Path $deskLnk) { Remove-Item $deskLnk -Force }

$smLnk = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Railway\$AppName.lnk"
if (Test-Path $smLnk)   { Remove-Item $smLnk -Force }

Remove-ItemProperty -Path $StartupKey -Name $AppName -ErrorAction SilentlyContinue

Write-Host "$AppName has been removed." -ForegroundColor Green
Write-Host ""
Start-Sleep -Seconds 3
