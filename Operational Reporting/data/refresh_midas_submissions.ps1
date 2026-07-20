param(
  [string]$Server = 'WINPRDAF3350',
  [string]$Database = 'AutomateMetrics',
  [string]$OutFile = 'c:/Users/m347/Documents/Power BI Desktop/SecureShare Report/secureshare-phase3/Operational Reporting/data/midas_submissions_data.js',
  [string]$StartDate = '2023-10-01',
  [int]$MaxFutureDateDays = 0,
  [int]$MaxRows = 0
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$today = [DateTime]::Today
$maxAllowedDate = $today.AddDays($MaxFutureDateDays)

$sql = @"
WITH src AS (
  SELECT
    DateCreated = TRY_CONVERT(date, TP2V3_DATE_CREATED),
    DateCompleted = TRY_CONVERT(date, TP2V3_WI_DATE_COMPLETED),
    SchemeNumber = COALESCE(NULLIF(LTRIM(RTRIM(TP2V3_SCHEME_NO)), ''), NULLIF(LTRIM(RTRIM(TP2V3_WI_LINK_SCHEME_NO)), '')),
    SchemeName = NULLIF(LTRIM(RTRIM(TP2V3_WI_SCHEME_NAME)), ''),
    EmployerOrg = NULLIF(LTRIM(RTRIM(TP2V3_WI_SCHEME_OWNER)), ''),
    BrokerOrg = NULLIF(LTRIM(RTRIM(TP2V3_WI_BROKER_NAME)), ''),
    SubmitterEmail = COALESCE(NULLIF(LTRIM(RTRIM(TP2V3_WI_REQUEST_USER_ID)), ''), NULLIF(LTRIM(RTRIM(ER_FROM_ADDRESS)), ''), NULLIF(LTRIM(RTRIM(TP2V3_WI_ALLOCATE_USER)), '')),
    WorkContext = NULLIF(LTRIM(RTRIM(TP2V3_WI_CONTEXT)), ''),
    WorkTypeCode = NULLIF(LTRIM(RTRIM(TP2V3_WI_WORKTYPE)), ''),
    WorkTypeDescription = NULLIF(LTRIM(RTRIM(TP2V3_WI_WORKTYPE_DESC)), ''),
    WorkStatus = NULLIF(LTRIM(RTRIM(TP2V3_WI_STATUS)), ''),
    TopScheme = NULLIF(LTRIM(RTRIM(TP2V3_WI_TOP_SCHEME)), '')
  FROM dbo.vw_ILCB_CUSTOM_DATA_G360
), filtered AS (
  SELECT
    DateCreated,
    DateCompleted,
    SchemeNumber,
    SchemeName,
    EmployerOrg,
    BrokerOrg,
    SubmitterEmail,
    WorkContext,
    WorkTypeCode,
    WorkTypeDescription,
    WorkStatus,
    TopScheme,
    RowNum = ROW_NUMBER() OVER (ORDER BY DateCreated DESC, SchemeNumber DESC)
  FROM src
  WHERE DateCreated >= TRY_CONVERT(date, @StartDate)
    AND SchemeNumber IS NOT NULL
), limited AS (
  SELECT
    DateCreated,
    DateCompleted,
    SchemeNumber,
    SchemeName,
    EmployerOrg,
    BrokerOrg,
    SubmitterEmail,
    WorkContext,
    WorkTypeCode,
    WorkTypeDescription,
    WorkStatus,
    TopScheme
  FROM filtered
  WHERE @MaxRows <= 0 OR RowNum <= @MaxRows
)
SELECT
  [DateCreated] = CONVERT(varchar(10), DateCreated, 23),
  [DateCompleted] = CASE WHEN DateCompleted IS NULL THEN NULL ELSE CONVERT(varchar(10), DateCompleted, 23) END,
  [SchemeNumber] = ISNULL(SchemeNumber, ''),
  [SchemeName] = ISNULL(SchemeName, ''),
  [Employer_Org] = ISNULL(EmployerOrg, ''),
  [Broker_Org] = ISNULL(BrokerOrg, ''),
  [SubmitterEmail] = ISNULL(SubmitterEmail, ''),
  [WorkContext] = ISNULL(WorkContext, ''),
  [WorkTypeCode] = ISNULL(WorkTypeCode, ''),
  [WorkTypeDescription] = ISNULL(WorkTypeDescription, ''),
  [WorkStatus] = ISNULL(WorkStatus, ''),
  [GoldScheme] = CASE WHEN LOWER(TopScheme) IN ('y','yes','true','1','gold') THEN 'Yes' ELSE 'No' END
FROM limited
ORDER BY DateCreated, SchemeNumber;
"@

$conn = New-Object System.Data.SqlClient.SqlConnection("Server=$Server;Database=$Database;Integrated Security=True;TrustServerCertificate=True;")
$conn.Open()

$cmd = $conn.CreateCommand()
$cmd.CommandTimeout = 0
$cmd.CommandText = $sql
$null = $cmd.Parameters.Add('@StartDate', [System.Data.SqlDbType]::VarChar, 10)
$cmd.Parameters['@StartDate'].Value = $StartDate
$null = $cmd.Parameters.Add('@MaxRows', [System.Data.SqlDbType]::Int)
$cmd.Parameters['@MaxRows'].Value = $MaxRows

$da = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
$dt = New-Object System.Data.DataTable
[void]$da.Fill($dt)
$conn.Close()

$rows = foreach ($r in $dt.Rows) {
  [ordered]@{
    '[DateCreated]' = [string]$r['DateCreated']
    '[DateCompleted]' = if ($r['DateCompleted'] -eq [DBNull]::Value) { $null } else { [string]$r['DateCompleted'] }
    '[SchemeNumber]' = [string]$r['SchemeNumber']
    '[SchemeName]' = [string]$r['SchemeName']
    '[Employer_Org]' = [string]$r['Employer_Org']
    '[Broker_Org]' = [string]$r['Broker_Org']
    '[SubmitterEmail]' = [string]$r['SubmitterEmail']
    '[WorkContext]' = [string]$r['WorkContext']
    '[WorkTypeCode]' = [string]$r['WorkTypeCode']
    '[WorkTypeDescription]' = [string]$r['WorkTypeDescription']
    '[WorkStatus]' = [string]$r['WorkStatus']
    '[GoldScheme]' = [string]$r['GoldScheme']
  }
}

$futureRows = @($rows | Where-Object {
  $_['[DateCreated]'] -and ([DateTime]::ParseExact($_['[DateCreated]'], 'yyyy-MM-dd', $null)) -gt $maxAllowedDate
})

if ($futureRows.Count -gt 0) {
  Write-Warning ("Excluded {0} MIDAS row(s) with DateCreated after {1}. Review source data for these schemes: {2}" -f $futureRows.Count, $maxAllowedDate.ToString('yyyy-MM-dd'), (($futureRows | ForEach-Object { '{0} ({1})' -f $_['[SchemeNumber]'], $_['[DateCreated]'] }) -join ', '))
}

$rows = @($rows | Where-Object {
  -not $_['[DateCreated]'] -or ([DateTime]::ParseExact($_['[DateCreated]'], 'yyyy-MM-dd', $null)) -le $maxAllowedDate
})

$minDate = $null
$maxDate = $null
if ($rows.Count -gt 0) {
  $orderedDates = @($rows | ForEach-Object { $_['[DateCreated]'] } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object)
  if ($orderedDates.Count -gt 0) {
    $minDate = [DateTime]::ParseExact($orderedDates[0], 'yyyy-MM-dd', $null)
    $maxDate = [DateTime]::ParseExact($orderedDates[$orderedDates.Count - 1], 'yyyy-MM-dd', $null)
  }
}

if ($maxDate -and $maxDate -gt $today) {
  $maxDate = $today
}
if ($minDate -and $maxDate -and $minDate -gt $maxDate) {
  $minDate = $maxDate
}

$dateRange = if ($minDate -and $maxDate) {
  '{0} to {1}' -f $minDate.ToString('yyyy-MM-dd'), $maxDate.ToString('yyyy-MM-dd')
} else {
  'n/a'
}

$payloadObj = [ordered]@{
  metadata = [ordered]@{
    generatedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    source = "AutomateMetrics ($Server)"
    dateRange = $dateRange
    startDate = $StartDate
    maxRows = $MaxRows
    maxFutureDateDays = $MaxFutureDateDays
    excludedFutureRows = $futureRows.Count
  }
  data = $rows
}

$json = $payloadObj | ConvertTo-Json -Depth 6 -Compress
$payload = "window.MIDAS_SUBMISSIONS_PAYLOAD = $json;`r`nwindow.MIDAS_SUBMISSIONS_META = window.MIDAS_SUBMISSIONS_PAYLOAD.metadata;`r`nwindow.MIDAS_SUBMISSIONS_DATA = window.MIDAS_SUBMISSIONS_PAYLOAD.data;"
[System.IO.File]::WriteAllText($OutFile, $payload, (New-Object System.Text.UTF8Encoding($false)))

Write-Output ("Updated midas_submissions_data.js -> rows: {0}, excludedFutureRows: {1}, startDate: {2}, maxRows: {3}, maxFutureDateDays: {4}, server: {5}" -f $rows.Count, $futureRows.Count, $StartDate, $MaxRows, $MaxFutureDateDays, $Server)
