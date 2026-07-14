param(
  [string]$Server = 'WINPRDAF3350',
  [string]$Database = 'AutomateMetrics',
  [string]$MidasStartDate = '2023-10-01',
  [int]$MaxMidasRows = 300000,
  [string]$NeStartDate = '2025-01-01',
  [switch]$SkipSharpload,
  [switch]$SkipDcMaster
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host 'Refreshing MIDAS submissions export...'
& (Join-Path $scriptDir 'refresh_midas_submissions.ps1') -Server $Server -Database $Database -StartDate $MidasStartDate -MaxRows $MaxMidasRows

Write-Host 'Refreshing New Entrant submissions export...'
& (Join-Path $scriptDir 'refresh_ne_submissions.ps1') -Server $Server -Database $Database -StartDate $NeStartDate

if (-not $SkipDcMaster) {
  Write-Host 'Refreshing DC master data export...'
  & (Join-Path $scriptDir 'refresh_dc_master_list.ps1')
}

if (-not $SkipSharpload) {
  Write-Host 'Refreshing sharpload bulk export from workbook...'
  & (Join-Path $scriptDir 'export_sharpload.ps1')
}

Write-Host 'Operational Reporting data refresh complete.'
