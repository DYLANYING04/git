param(
  [Parameter(Mandatory = $true)]
  [string]$Xlsx,
  [switch]$SkipRender
)

$ErrorActionPreference = 'Stop'
$skillRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $skillRoot
$argsList = @('scripts/monthly_report.py', '--xlsx', $Xlsx)
if ($SkipRender) { $argsList += '--skip-render' }
python @argsList
