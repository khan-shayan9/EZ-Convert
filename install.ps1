# EZConvert installer / updater.
# Usage (PowerShell):  irm https://raw.githubusercontent.com/khan-shayan9/EZ-Convert/main/install.ps1 | iex
# Downloads the latest release to %LOCALAPPDATA%\EZConvert, installs Python packages
# and adds the right-click "Convert to" menu (current user only, no admin needed).

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$repo = 'khan-shayan9/EZ-Convert'
$dest = Join-Path $env:LOCALAPPDATA 'EZConvert'

function Fail($msg) { Write-Host $msg -ForegroundColor Red; return }

# 1. Python 3.10+ (skip the Microsoft Store "python" alias, which isn't a real install)
$python = $null
foreach ($cand in 'python', 'py') {
    $cmd = Get-Command $cand -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.Source -notlike '*WindowsApps*') {
        $ver = & $cand -c "import sys; print('%d.%d' % sys.version_info[:2])" 2>$null
        if ($ver -and [version]$ver -ge [version]'3.10') { $python = $cand; break }
    }
}
if (-not $python) {
    Fail "Python 3.10+ not found. Install it from https://www.python.org/downloads/ (tick 'Add to PATH'), or run:`n  winget install Python.Python.3.12`nthen run this command again."
    return
}

# 2. Download the latest release (falls back to main if there is no release yet)
try {
    $rel = Invoke-RestMethod "https://api.github.com/repos/$repo/releases/latest" -Headers @{ 'User-Agent' = 'EZConvert-installer' }
    $url = $rel.zipball_url; $version = $rel.tag_name
} catch {
    $url = "https://github.com/$repo/archive/refs/heads/main.zip"; $version = 'main'
}
Write-Host "Downloading EZConvert $version..."
$tmp = Join-Path ([IO.Path]::GetTempPath()) ("ezconvert_" + [guid]::NewGuid())
New-Item -ItemType Directory $tmp | Out-Null
Invoke-WebRequest $url -OutFile "$tmp\ez.zip" -UseBasicParsing
Expand-Archive "$tmp\ez.zip" "$tmp\x"
$src = Get-ChildItem "$tmp\x" -Directory | Select-Object -First 1

# 3. Copy into place (replaces an older install)
New-Item -ItemType Directory -Force $dest | Out-Null
Copy-Item "$($src.FullName)\*" $dest -Recurse -Force
Remove-Item $tmp -Recurse -Force

# 4. Packages + right-click menu
Write-Host "Installing Python packages (first time takes a minute)..."
& $python -m pip install --quiet --disable-pip-version-check -r "$dest\requirements.txt"
if ($LASTEXITCODE -ne 0) { Fail "Package install failed."; return }
& $python "$dest\install_menu.py"
if ($LASTEXITCODE -ne 0) { Fail "Menu setup failed."; return }

Write-Host "`nEZConvert $version installed to $dest" -ForegroundColor Green
Write-Host "Right-click a file and choose 'Convert to' (Windows 11: under 'Show more options')."
