# Headway Windows installer — per-user (installs under %LOCALAPPDATA%), no admin required.
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/smo-key/headway/main/install.ps1 | iex"
$ErrorActionPreference = 'Stop'

$repo = 'smo-key/headway'

Write-Host 'Fetching the latest Headway release…'
$release = Invoke-RestMethod "https://api.github.com/repos/$repo/releases/latest"
$asset = $release.assets | Where-Object { $_.name -like '*-setup.exe' } | Select-Object -First 1
if (-not $asset) { throw "Could not find a Windows installer in the latest release of $repo." }

$exe = Join-Path $env:TEMP $asset.name
Write-Host "Downloading $($asset.name)…"
Invoke-WebRequest $asset.browser_download_url -OutFile $exe

Write-Host 'Installing (silent, per-user)…'
Start-Process -FilePath $exe -ArgumentList '/S' -Wait
Remove-Item $exe -ErrorAction SilentlyContinue

Write-Host 'Headway installed. Find it in the Start menu.'
