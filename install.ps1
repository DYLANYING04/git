param(
  [switch]$SkipPackages
)

$ErrorActionPreference = 'Stop'
$skillRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $skillRoot

if (-not $SkipPackages) {
  python -m pip install -r requirements.txt
}

if (-not (Get-Command soffice -ErrorAction SilentlyContinue)) {
  Write-Warning 'LibreOffice was not found. Install LibreOffice before rendering previews.'
}
if (-not (Get-Command pdftoppm -ErrorAction SilentlyContinue)) {
  Write-Warning 'pdftoppm was not found. Install Poppler before rendering previews.'
}
Write-Host 'Nanyan proposal skill is ready.'
