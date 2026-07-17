param(
  [string]$Server = 'WINPRDAF3350',
  [string]$Database = 'AutomateMetrics',
  [string]$OutFile = 'c:/Users/m347/Documents/Power BI Desktop/SecureShare Report/secureshare-phase3/Operational Reporting/data/dc_master_list_data.js',
  [string]$StartDate = '2025-05-01'
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$sql = @"
WITH mapped AS (
  SELECT
    SchemeNumber = COALESCE(NULLIF(LTRIM(RTRIM(TP2V3_SCHEME_NO)), ''), NULLIF(LTRIM(RTRIM(TP2V3_WI_LINK_SCHEME_NO)), '')),
    SchemeName = NULLIF(LTRIM(RTRIM(TP2V3_WI_SCHEME_NAME)), ''),
    Employer_Org = NULLIF(LTRIM(RTRIM(TP2V3_WI_SCHEME_OWNER)), ''),
    Broker_Org = NULLIF(LTRIM(RTRIM(TP2V3_WI_BROKER_NAME)), ''),
    TopScheme = NULLIF(LTRIM(RTRIM(TP2V3_WI_TOP_SCHEME)), '')
  FROM dbo.vw_ILCB_CUSTOM_DATA_G360
  WHERE TP2V3_DATE_CREATED >= TRY_CONVERT(date, @StartDate)
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

$conn = New-Object System.Data.SqlClient.SqlConnection("Server=$Server;Database=$Database;Integrated Security=True;TrustServerCertificate=True;")
$conn.Open()
$cmd = $conn.CreateCommand()
$cmd.CommandTimeout = 0
$cmd.CommandText = $sql
$null = $cmd.Parameters.Add('@StartDate', [System.Data.SqlDbType]::VarChar, 10)
$cmd.Parameters['@StartDate'].Value = $StartDate
$da = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
$dt = New-Object System.Data.DataTable
[void]$da.Fill($dt)
$conn.Close()

$schemes = foreach ($row in $dt.Rows) {
  $schemeNo = [string]$row['SchemeNumber']
  if ([string]::IsNullOrWhiteSpace($schemeNo)) { continue }

  $topScheme = if ($row['TopScheme'] -ne [DBNull]::Value) { [string]$row['TopScheme'] } else { '' }
  [PSCustomObject]@{
    SchemeNumber = $schemeNo.Trim()
    SchemeName   = if ($row['SchemeName'] -ne [DBNull]::Value) { [string]$row['SchemeName'] } else { '' }
    Employer_Org = if ($row['Employer_Org'] -ne [DBNull]::Value) { [string]$row['Employer_Org'] } else { '' }
    Broker_Org   = if ($row['Broker_Org'] -ne [DBNull]::Value) { [string]$row['Broker_Org'] } else { '' }
    GoldScheme   = if ($topScheme -match '^(?i:y|yes|true|1|gold)$') { 'Gold' } else { '' }
  }
}

$schemes = @($schemes) | Sort-Object SchemeNumber -Unique
$employers = $schemes | ForEach-Object { $_.Employer_Org } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique
$brokers = $schemes | ForEach-Object { $_.Broker_Org } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique

$payload = [ordered]@{
  generatedAt   = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
  source        = "AutomateMetrics ($Server)"
  schemeCount   = $schemes.Count
  employerCount = $employers.Count
  brokerCount   = $brokers.Count
  schemes       = $schemes
  employers     = $employers
  brokers       = $brokers
}

$out = 'window.DC_MASTER_DATA = ' + ($payload | ConvertTo-Json -Depth 6 -Compress) + ';'
[System.IO.File]::WriteAllText($OutFile, $out, (New-Object System.Text.UTF8Encoding($false)))

Write-Output ('Updated dc_master_list_data.js -> schemes: {0}, employers: {1}, brokers: {2}, startDate: {3}, server: {4}' -f $schemes.Count, $employers.Count, $brokers.Count, $StartDate, $Server)
