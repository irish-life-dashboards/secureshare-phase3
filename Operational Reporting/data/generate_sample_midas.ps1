param(
  [int]$RowsPerMonth = 1000,
  [int]$MonthsBack = 7,
  [string]$OutFile = 'c:/Users/m347/Documents/Power BI Desktop/SecureShare Report/secureshare-phase3/Operational Reporting/data/midas_submissions_data.js'
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

Write-Output "Generating sample MIDAS data: $MonthsBack months x $RowsPerMonth rows/month"

$brokers = @("Broker A", "Broker B", "Broker C", "Broker D", "Broker E")
$employers = @("Employer 1", "Employer 2", "Employer 3", "Employer 4")
$schemes = @("Scheme001", "Scheme002", "Scheme003", "Scheme004", "Scheme005", "Scheme006", "Scheme007", "Scheme008")
$workTypes = @("Auto Quote", "New Entrant Bulk Auto", "New Entrant Single Auto", "Cash / Direct Debits", "Scheme Update")
$statuses = @("Completed", "Completed", "Completed", "Completed", "Completed", "Completed", "Pending")

$rows = @()
$today = Get-Date
$totalRows = 0

for ($m = $MonthsBack; $m -ge 0; $m--) {
  $monthDate = $today.AddMonths(-$m)
  $monthStr = $monthDate.ToString("yyyy-MM")
  $daysInMonth = [DateTime]::DaysInMonth($monthDate.Year, $monthDate.Month)
  
  for ($i = 0; $i -lt $RowsPerMonth; $i++) {
    $day = Get-Random -Min 1 -Max ($daysInMonth + 1)
    $dateCreated = "{0:yyyy-MM-dd}" -f [DateTime]::new($monthDate.Year, $monthDate.Month, $day)
    
    $completed = Get-Random -Min 0 -Max 2
    $dateCompleted = if ($completed -eq 1) {
      $completedDay = Get-Random -Min $day -Max ($daysInMonth + 1)
      "{0:yyyy-MM-dd}" -f [DateTime]::new($monthDate.Year, $monthDate.Month, $completedDay)
    } else {
      $null
    }
    
    $schemeNum = $schemes | Get-Random
    $rowObj = [PSCustomObject]@{
      '[DateCreated]' = $dateCreated
      '[DateCompleted]' = $dateCompleted
      '[SchemeNumber]' = $schemeNum
      '[SchemeName]' = "Scheme Name for $schemeNum"
      '[Employer_Org]' = $employers | Get-Random
      '[Broker_Org]' = $brokers | Get-Random
      '[SubmitterEmail]' = "user$(Get-Random -Min 1 -Max 100)@example.com"
      '[WorkContext]' = "SecureShare"
      '[WorkTypeCode]' = "AUTO"
      '[WorkTypeDescription]' = $workTypes | Get-Random
      '[WorkStatus]' = $statuses | Get-Random
      '[GoldScheme]' = if ((Get-Random -Min 0 -Max 2) -eq 1) { "Yes" } else { "No" }
    }
    $rows += $rowObj
    $totalRows++
  }
  Write-Output "  Generated $($RowsPerMonth) rows for $monthStr"
}

Write-Output "Converting $totalRows rows to JSON..."
$json = $rows | ConvertTo-Json -Compress -Depth 1

$payload = "window.MIDAS_SUBMISSIONS_DATA = $json;"
Write-Output "Writing to $OutFile..."

[System.IO.File]::WriteAllText($OutFile, $payload, (New-Object System.Text.UTF8Encoding($false)))

$fileSizeMB = [math]::Round((Get-Item $OutFile).Length / 1MB, 2)
Write-Output "[SUCCESS] Generated $totalRows rows in $fileSizeMB MB"
