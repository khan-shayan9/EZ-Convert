# EZConvert uninstaller.
# Usage (PowerShell):  irm https://raw.githubusercontent.com/khan-shayan9/EZ-Convert/main/uninstall.ps1 | iex
# Removes the right-click menu and the %LOCALAPPDATA%\EZConvert folder. Python packages are left installed.

$dest = Join-Path $env:LOCALAPPDATA 'EZConvert'
$base = 'HKCU:\Software\Classes\SystemFileAssociations'
Get-ChildItem $base -ErrorAction SilentlyContinue | ForEach-Object {
    $key = Join-Path $_.PSPath 'shell\EZConvert'
    if (Test-Path $key) { Remove-Item $key -Recurse -Force }
}
if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
Write-Host "EZConvert removed." -ForegroundColor Green
