param(
  [string]$Server = 'WINPRDAF3350',
  [string]$Database = 'AutomateMetrics',
  [string]$MidasStartDate = '2023-10-01',
  [int]$MaxMidasRows = 300000,
  [string]$NeStartDate = '2025-01-01',
  [double]$MaxSourceAgeHours = 24,
  [switch]$SkipDcMaster,
  [switch]$IncludeSharpload,
  [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

function Get-SourceFreshness {
  param(
    [string]$Server,
    [string]$Database
  )

  $sql = @"
SELECT
  GETDATE() AS ServerNow,
  MAX(CASE WHEN TP2V3_DATE_CREATED <= GETDATE() THEN TP2V3_DATE_CREATED END) AS LatestCreatedAt,
  SUM(CASE WHEN CAST(TP2V3_DATE_CREATED AS date) = CAST(GETDATE() AS date) THEN 1 ELSE 0 END) AS RowsCreatedToday
FROM dbo.vw_ILCB_CUSTOM_DATA_G360;
"@

  $conn = New-Object System.Data.SqlClient.SqlConnection("Server=$Server;Database=$Database;Integrated Security=True;TrustServerCertificate=True;")
  $conn.Open()

  try {
    $cmd = $conn.CreateCommand()
    $cmd.CommandTimeout = 0
    $cmd.CommandText = $sql

    $reader = $cmd.ExecuteReader()
    $row = $null

    if ($reader.Read()) {
      $row = [PSCustomObject]@{
        ServerNow = [datetime]$reader['ServerNow']
        LatestCreatedAt = if ($reader['LatestCreatedAt'] -eq [DBNull]::Value) { $null } else { [datetime]$reader['LatestCreatedAt'] }
        RowsCreatedToday = if ($reader['RowsCreatedToday'] -eq [DBNull]::Value) { 0 } else { [int]$reader['RowsCreatedToday'] }
      }
    }

    $reader.Close()
    return $row
  }
  finally {
    $conn.Close()
  }
}

function Get-OutputFilesSummary {
  param([string]$DataDir)

  $files = @('midas_submissions_data.js', 'ne_submissions_data.js', 'dc_master_list_data.js', 'sharpload_bulk_data.js')
  $summary = foreach ($name in $files) {
    $path = Join-Path $DataDir $name
    if (Test-Path $path) {
      $item = Get-Item $path
      [PSCustomObject]@{
        File = $item.Name
        LastWriteTime = $item.LastWriteTime
        SizeKB = [math]::Round($item.Length / 1kb, 2)
      }
    }
    else {
      [PSCustomObject]@{
        File = $name
        LastWriteTime = $null
        SizeKB = $null
      }
    }
  }

  return $summary
}

Write-Host "=== Operational Reporting Source Pre-check ===" -ForegroundColor Cyan
$source = Get-SourceFreshness -Server $Server -Database $Database

if ($null -eq $source -or $null -eq $source.LatestCreatedAt) {
  throw 'Unable to determine source freshness from dbo.vw_ILCB_CUSTOM_DATA_G360.'
}

$ageHours = [math]::Round(($source.ServerNow - $source.LatestCreatedAt).TotalHours, 2)

Write-Host ("ServerNow        : {0}" -f $source.ServerNow.ToString('yyyy-MM-dd HH:mm:ss'))
Write-Host ("LatestCreatedAt  : {0}" -f $source.LatestCreatedAt.ToString('yyyy-MM-dd HH:mm:ss'))
Write-Host ("RowsCreatedToday : {0}" -f $source.RowsCreatedToday)
Write-Host ("SourceAgeHours   : {0}" -f $ageHours)

if ($ageHours -gt $MaxSourceAgeHours) {
  throw ("Source appears stale. LatestCreatedAt is {0} hours old, over MaxSourceAgeHours={1}." -f $ageHours, $MaxSourceAgeHours)
}

if ($DryRun) {
  Write-Host 'DryRun enabled: source pre-check passed, no refresh scripts executed.' -ForegroundColor Yellow
  exit 0
}

Write-Host "=== Running Refresh Scripts ===" -ForegroundColor Cyan

Write-Host 'Refreshing MIDAS submissions...'
& (Join-Path $scriptDir 'refresh_midas_submissions.ps1') -Server $Server -Database $Database -StartDate $MidasStartDate -MaxRows $MaxMidasRows

Write-Host 'Refreshing NE submissions...'
& (Join-Path $scriptDir 'refresh_ne_submissions.ps1') -Server $Server -Database $Database -StartDate $NeStartDate

if (-not $SkipDcMaster) {
  Write-Host 'Refreshing DC master...'
  & (Join-Path $scriptDir 'refresh_dc_master_list.ps1') -Server $Server -Database $Database
}

if ($IncludeSharpload) {
  Write-Host 'Refreshing Sharpload export...'
  & (Join-Path $scriptDir 'export_sharpload.ps1')
}

Write-Host "=== Post-refresh Output File Check ===" -ForegroundColor Cyan
Get-OutputFilesSummary -DataDir $scriptDir | Format-Table -AutoSize

Write-Host 'Operational Reporting refresh with checks completed successfully.' -ForegroundColor Green
