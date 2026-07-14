$ErrorActionPreference = 'Stop'

$masterPath = 'c:/Users/m347/Documents/Power BI Desktop/SecureShare Report/secureshare-phase3/Operational Reporting/dc_master_list_data.js'

# Read existing May-26 baseline payload.
$raw = [System.IO.File]::ReadAllText($masterPath)
$prefix = 'window.DC_MASTER_DATA = '
if (-not $raw.StartsWith($prefix)) {
  throw 'Unexpected dc_master_list_data.js format.'
}

$jsonText = $raw.Substring($prefix.Length).Trim()
if ($jsonText.EndsWith(';')) {
  $jsonText = $jsonText.Substring(0, $jsonText.Length - 1)
}

$existing = $jsonText | ConvertFrom-Json
$existingSchemes = @($existing.schemes)
$existingMap = @{}

foreach ($s in $existingSchemes) {
  if ($null -ne $s.SchemeNumber) {
    $existingMap[[string]$s.SchemeNumber] = [PSCustomObject]@{
      SchemeNumber = [string]$s.SchemeNumber
      SchemeName   = [string]$s.SchemeName
      Employer_Org = [string]$s.Employer_Org
      Broker_Org   = [string]$s.Broker_Org
      GoldScheme   = [string]$s.GoldScheme
    }
  }
}

# Pull latest scheme attributes from AutomateMetrics and pick best values by frequency.
$sql = @"
WITH mapped AS (
  SELECT
    SchemeNumber = COALESCE(NULLIF(LTRIM(RTRIM(TP2V3_SCHEME_NO)), ''), NULLIF(LTRIM(RTRIM(TP2V3_WI_LINK_SCHEME_NO)), '')),
    SchemeName = NULLIF(LTRIM(RTRIM(TP2V3_WI_SCHEME_NAME)), ''),
    Employer_Org = NULLIF(LTRIM(RTRIM(TP2V3_WI_SCHEME_OWNER)), ''),
    Broker_Org = NULLIF(LTRIM(RTRIM(TP2V3_WI_BROKER_NAME)), ''),
    TopScheme = NULLIF(LTRIM(RTRIM(TP2V3_WI_TOP_SCHEME)), '')
  FROM dbo.vw_ILCB_CUSTOM_DATA_G360
  WHERE TP2V3_DATE_CREATED >= '2025-05-01'
),
clean AS (
  SELECT *
  FROM mapped
  WHERE SchemeNumber IS NOT NULL
    AND SchemeNumber NOT IN ('.', '-', '0')
),
name_rank AS (
  SELECT SchemeNumber, SchemeName, Cnt = COUNT(*)
  FROM clean
  WHERE SchemeName IS NOT NULL
  GROUP BY SchemeNumber, SchemeName
),
name_pick AS (
  SELECT SchemeNumber, SchemeName,
         ROW_NUMBER() OVER (PARTITION BY SchemeNumber ORDER BY Cnt DESC, LEN(SchemeName) DESC, SchemeName) AS rn
  FROM name_rank
),
emp_rank AS (
  SELECT SchemeNumber, Employer_Org, Cnt = COUNT(*)
  FROM clean
  WHERE Employer_Org IS NOT NULL
  GROUP BY SchemeNumber, Employer_Org
),
emp_pick AS (
  SELECT SchemeNumber, Employer_Org,
         ROW_NUMBER() OVER (PARTITION BY SchemeNumber ORDER BY Cnt DESC, LEN(Employer_Org) DESC, Employer_Org) AS rn
  FROM emp_rank
),
broker_rank AS (
  SELECT SchemeNumber, Broker_Org, Cnt = COUNT(*)
  FROM clean
  WHERE Broker_Org IS NOT NULL
  GROUP BY SchemeNumber, Broker_Org
),
broker_pick AS (
  SELECT SchemeNumber, Broker_Org,
         ROW_NUMBER() OVER (PARTITION BY SchemeNumber ORDER BY Cnt DESC, LEN(Broker_Org) DESC, Broker_Org) AS rn
  FROM broker_rank
),
top_rank AS (
  SELECT SchemeNumber, TopScheme, Cnt = COUNT(*)
  FROM clean
  WHERE TopScheme IS NOT NULL
  GROUP BY SchemeNumber, TopScheme
),
top_pick AS (
  SELECT SchemeNumber, TopScheme,
         ROW_NUMBER() OVER (PARTITION BY SchemeNumber ORDER BY Cnt DESC, TopScheme) AS rn
  FROM top_rank
)
SELECT
  s.SchemeNumber,
  n.SchemeName,
  e.Employer_Org,
  b.Broker_Org,
  t.TopScheme
FROM (SELECT DISTINCT SchemeNumber FROM clean) s
LEFT JOIN name_pick n ON n.SchemeNumber = s.SchemeNumber AND n.rn = 1
LEFT JOIN emp_pick e ON e.SchemeNumber = s.SchemeNumber AND e.rn = 1
LEFT JOIN broker_pick b ON b.SchemeNumber = s.SchemeNumber AND b.rn = 1
LEFT JOIN top_pick t ON t.SchemeNumber = s.SchemeNumber AND t.rn = 1
ORDER BY s.SchemeNumber;
"@

$conn = New-Object System.Data.SqlClient.SqlConnection('Server=WINPRDAF3350;Database=AutomateMetrics;Integrated Security=True;TrustServerCertificate=True;')
$conn.Open()
$cmd = $conn.CreateCommand()
$cmd.CommandTimeout = 0
$cmd.CommandText = $sql
$da = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
$dt = New-Object System.Data.DataTable
[void]$da.Fill($dt)
$conn.Close()

$updated = @{}
foreach ($key in $existingMap.Keys) {
  $updated[$key] = $existingMap[$key]
}

foreach ($row in $dt.Rows) {
  $schemeNo = [string]$row['SchemeNumber']
  if ([string]::IsNullOrWhiteSpace($schemeNo)) { continue }

  $schemeNo = $schemeNo.Trim()
  $schemeName = if ($row['SchemeName'] -ne [DBNull]::Value) { [string]$row['SchemeName'] } else { '' }
  $employer = if ($row['Employer_Org'] -ne [DBNull]::Value) { [string]$row['Employer_Org'] } else { '' }
  $broker = if ($row['Broker_Org'] -ne [DBNull]::Value) { [string]$row['Broker_Org'] } else { '' }
  $topScheme = if ($row['TopScheme'] -ne [DBNull]::Value) { [string]$row['TopScheme'] } else { '' }

  if ($updated.ContainsKey($schemeNo)) {
    $curr = $updated[$schemeNo]
    if (-not [string]::IsNullOrWhiteSpace($schemeName)) { $curr.SchemeName = $schemeName }
    if (-not [string]::IsNullOrWhiteSpace($employer)) { $curr.Employer_Org = $employer }
    if (-not [string]::IsNullOrWhiteSpace($broker)) { $curr.Broker_Org = $broker }
    if ([string]::IsNullOrWhiteSpace($curr.GoldScheme) -and ($topScheme -match '^(?i:y|yes|true|1|gold)$')) {
      $curr.GoldScheme = 'Gold'
    }
  } else {
    $gold = if ($topScheme -match '^(?i:y|yes|true|1|gold)$') { 'Gold' } else { '' }
    $updated[$schemeNo] = [PSCustomObject]@{
      SchemeNumber = $schemeNo
      SchemeName   = $schemeName
      Employer_Org = $employer
      Broker_Org   = $broker
      GoldScheme   = $gold
    }
  }
}

$schemes = @($updated.Values) | Sort-Object SchemeNumber
$employers = $schemes | ForEach-Object { $_.Employer_Org } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique
$brokers = $schemes | ForEach-Object { $_.Broker_Org } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique

$payload = [ordered]@{
  generatedAt   = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
  source        = 'DC Schemes Master List - May26 + AutomateMetrics (WINPRDAF3350)'
  schemeCount   = $schemes.Count
  employerCount = $employers.Count
  brokerCount   = $brokers.Count
  schemes       = $schemes
  employers     = $employers
  brokers       = $brokers
}

$out = 'window.DC_MASTER_DATA = ' + ($payload | ConvertTo-Json -Depth 6 -Compress) + ';'
[System.IO.File]::WriteAllText($masterPath, $out, (New-Object System.Text.UTF8Encoding($false)))

Write-Output ('Updated dc_master_list_data.js -> schemes: {0}, employers: {1}, brokers: {2}' -f $schemes.Count, $employers.Count, $brokers.Count)
