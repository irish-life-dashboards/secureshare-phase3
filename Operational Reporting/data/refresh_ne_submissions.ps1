param(
  [string]$Server = 'WINPRDAF3350',
  [string]$Database = 'AutomateMetrics',
  [string]$OutFile = 'c:/Users/m347/Documents/Power BI Desktop/SecureShare Report/secureshare-phase3/Operational Reporting/data/ne_submissions_data.js',
  [string]$StartDate = '2025-01-01'
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$sql = @"
WITH src AS (
  SELECT
    DateCreated = TRY_CONVERT(date, TP2V3_DATE_CREATED),
    Via = COALESCE(NULLIF(LTRIM(RTRIM(TP2V3_WI_SOURCE)), ''), NULLIF(LTRIM(RTRIM(TP2V3_SOURCE_INTEXT)), ''), 'Null'),
    WorkTypeCode = UPPER(LTRIM(RTRIM(TP2V3_WI_WORKTYPE))),
    WorkTypeDesc = NULLIF(LTRIM(RTRIM(TP2V3_WI_WORKTYPE_DESC)), ''),
    BrokerCode = COALESCE(NULLIF(LTRIM(RTRIM(TP2V3_BROKER_CD)), ''), 'NULL'),
    BrokerName = COALESCE(NULLIF(LTRIM(RTRIM(TP2V3_WI_BROKER_NAME)), ''), 'NULL'),
    SchemeNo = COALESCE(NULLIF(LTRIM(RTRIM(TP2V3_SCHEME_NO)), ''), NULLIF(LTRIM(RTRIM(TP2V3_WI_LINK_SCHEME_NO)), ''), 'NULL')
  FROM dbo.vw_ILCB_CUSTOM_DATA_G360
), filtered AS (
  SELECT *
  FROM src
  WHERE DateCreated >= TRY_CONVERT(date, @StartDate)
    AND WorkTypeCode IN ('NEBAUTO', 'NESAUTO', 'NEUW', 'NEVCM', 'NEVR', 'NES', 'RNEBAUTO', 'RNESAUTO')
)
SELECT
  YearMonth = CONVERT(varchar(7), DateCreated, 23),
  Via,
  WorkTypeCode,
  WorkTypeDesc = ISNULL(WorkTypeDesc, WorkTypeCode),
  BrokerCode,
  BrokerName,
  SchemeNo,
  SubmissionCount = COUNT_BIG(1)
FROM filtered
GROUP BY
  CONVERT(varchar(7), DateCreated, 23),
  Via,
  WorkTypeCode,
  ISNULL(WorkTypeDesc, WorkTypeCode),
  BrokerCode,
  BrokerName,
  SchemeNo
ORDER BY YearMonth, Via, WorkTypeCode, BrokerName, SchemeNo;
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

$rows = foreach ($r in $dt.Rows) {
  [ordered]@{
    YearMonth = [string]$r['YearMonth']
    Via = [string]$r['Via']
    WorkTypeCode = [string]$r['WorkTypeCode']
    WorkTypeDesc = [string]$r['WorkTypeDesc']
    BrokerCode = [string]$r['BrokerCode']
    BrokerName = [string]$r['BrokerName']
    SchemeNo = [string]$r['SchemeNo']
    SubmissionCount = [int64]$r['SubmissionCount']
  }
}

$header = @(
  '// New Entrant Submissions Data - Source: WINPRDAF3350 / AutomateMetrics',
  '// Columns: YearMonth, Via, WorkTypeCode, WorkTypeDesc, BrokerCode, BrokerName, SchemeNo, SubmissionCount'
) -join "`r`n"

$yearMonths = @($rows | ForEach-Object { $_.YearMonth } | Where-Object { $_ -match '^\d{4}-\d{2}$' } | Sort-Object -Unique)
$maxCurrentYm = (Get-Date).ToString('yyyy-MM')
$yearMonths = @($yearMonths | Where-Object { $_ -le $maxCurrentYm })
$dateRange = if ($yearMonths.Count -gt 0) {
  '{0} to {1}' -f $yearMonths[0], $yearMonths[$yearMonths.Count - 1]
} else {
  'n/a'
}

$payloadObj = [ordered]@{
  metadata = [ordered]@{
    generatedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    source = "AutomateMetrics ($Server)"
    dateRange = $dateRange
    startDate = $StartDate
  }
  data = $rows
}

$json = $payloadObj | ConvertTo-Json -Depth 6 -Compress
$payload = "$header`r`nwindow.NE_SUBMISSIONS_PAYLOAD = $json;`r`nwindow.NE_SUBMISSIONS_META = window.NE_SUBMISSIONS_PAYLOAD.metadata;`r`nvar ne_submissions_data = window.NE_SUBMISSIONS_PAYLOAD.data;"
[System.IO.File]::WriteAllText($OutFile, $payload, (New-Object System.Text.UTF8Encoding($false)))

Write-Output ("Updated ne_submissions_data.js -> rows: {0}, startDate: {1}, server: {2}" -f $rows.Count, $StartDate, $Server)
