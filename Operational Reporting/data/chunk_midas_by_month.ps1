param(
  [string]$InputFile = 'c:/Users/m347/Documents/Power BI Desktop/SecureShare Report/secureshare-phase3/Operational Reporting/data/midas_submissions_data.js',
  [string]$OutputDir = 'c:/Users/m347/Documents/Power BI Desktop/SecureShare Report/secureshare-phase3/Operational Reporting/data/chunks'
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# Create output directory
if (-not (Test-Path $OutputDir)) {
  New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

Write-Output "Reading $InputFile and extracting data array..."

# Read the file and extract the JSON array more efficiently
$reader = New-Object System.IO.StreamReader($InputFile)
$content = $reader.ReadToEnd()
$reader.Close()

# Find the start and end of the JSON array
$startIdx = $content.IndexOf('window.MIDAS_SUBMISSIONS_DATA = [')
if ($startIdx -eq -1) {
  throw "Could not find 'window.MIDAS_SUBMISSIONS_DATA = [' in file"
}
$startIdx += 33  # length of 'window.MIDAS_SUBMISSIONS_DATA = '

# Find the end of the array (the closing ])
$endIdx = $content.LastIndexOf('];')
if ($endIdx -eq -1) {
  throw "Could not find '];' in file"
}

$jsonStr = $content.Substring($startIdx, $endIdx - $startIdx)

Write-Output "Parsing JSON array..."
$data = $jsonStr | ConvertFrom-Json

Write-Output "Loaded $($data.Count) rows. Grouping by month..."

# Group by month (YYYY-MM) - store as PSCustomObject arrays
$byMonth = @{}
$monthOrder = @()
foreach ($row in $data) {
  $dateStr = $row.('[DateCreated]')
  if ($dateStr -match '^\d{4}-\d{2}') {
    $month = $dateStr.Substring(0, 7)
    if (-not $byMonth.ContainsKey($month)) {
      $byMonth[$month] = @()
      $monthOrder += $month
    }
    $byMonth[$month] += $row
  }
}

Write-Output "Generating monthly chunk files in $OutputDir..."

foreach ($month in ($monthOrder | Sort-Object)) {
  $rowsForMonth = $byMonth[$month]
  $json = $rowsForMonth | ConvertTo-Json -Depth 6 -Compress
  $payload = "window.MIDAS_SUBMISSIONS_DATA_$($month -replace '-', '_') = $json;`r`n"
  
  $outputFile = Join-Path $OutputDir "midas_$month.js"
  [System.IO.File]::WriteAllText($outputFile, $payload, (New-Object System.Text.UTF8Encoding($false)))
  
  $sizeMB = [math]::Round((Get-Item $outputFile).Length / 1MB, 2)
  Write-Output "  [OK] $month : $($rowsForMonth.Count) rows -> $outputFile ($sizeMB MB)"
}

Write-Output "Done! Generated $($byMonth.Count) monthly chunk files."
