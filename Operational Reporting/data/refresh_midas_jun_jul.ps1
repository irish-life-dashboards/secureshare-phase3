param(
  [string]$Server = 'WINPRDAF3350',
  [string]$Database = 'AutomateMetrics',
  [string]$OutFile = 'c:/Users/m347/Documents/Power BI Desktop/SecureShare Report/secureshare-phase3/Operational Reporting/data/midas_submissions_data.js',
  [int]$MaxRowsPerMonth = 10000,
  [int]$MaxFutureDateDays = 0
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
    RowNumInMonth = ROW_NUMBER() OVER (
      PARTITION BY CONVERT(char(7), DateCreated, 23)
      ORDER BY DateCreated DESC, SchemeNumber DESC
    )
  FROM src
  WHERE DateCreated >= '2026-01-01'
    AND DateCreated < '2026-08-01'
    AND SchemeNumber IS NOT NULL
    AND (
      UPPER(ISNULL(WorkTypeCode, '')) LIKE '%AUTO%'
      OR UPPER(ISNULL(WorkTypeDescription, '')) LIKE '%SECURESHARE%'
      OR UPPER(ISNULL(WorkTypeDescription, '')) LIKE '%CASH%'
      OR UPPER(ISNULL(WorkTypeDescription, '')) LIKE '%DIRECT DEBIT%'
      OR UPPER(ISNULL(WorkTypeDescription, '')) LIKE '%NEW ENTRANT%'
      OR UPPER(ISNULL(WorkTypeDescription, '')) LIKE '%NE BULK%'
      OR UPPER(ISNULL(WorkTypeDescription, '')) LIKE '%NE SINGLE%'
      OR UPPER(ISNULL(WorkTypeCode, '')) LIKE '%NE%'
    )
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
FROM filtered
WHERE @MaxRowsPerMonth <= 0 OR RowNumInMonth <= @MaxRowsPerMonth
ORDER BY DateCreated, SchemeNumber;
"@

$conn = New-Object System.Data.SqlClient.SqlConnection("Server=$Server;Database=$Database;Integrated Security=True;TrustServerCertificate=True;")
$conn.Open()

$cmd = $conn.CreateCommand()
$cmd.CommandTimeout = 0
$cmd.CommandText = $sql
$null = $cmd.Parameters.Add('@MaxRowsPerMonth', [System.Data.SqlDbType]::Int)
$cmd.Parameters['@MaxRowsPerMonth'].Value = $MaxRowsPerMonth

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

$rows = @($rows | Where-Object {
  -not $_['[DateCreated]'] -or ([DateTime]::ParseExact($_['[DateCreated]'], 'yyyy-MM-dd', $null)) -le $maxAllowedDate
})

$payloadObj = [ordered]@{
  metadata = [ordered]@{
    generatedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    source = "AutomateMetrics ($Server)"
    dateRange = '2026-01-01 to 2026-07-31'
    extraction = 'Jan-Jul SecureShare/Cash rows'
    maxRowsPerMonth = $MaxRowsPerMonth
    excludedFutureRows = $futureRows.Count
  }
  data = $rows
}

$json = $payloadObj | ConvertTo-Json -Depth 6 -Compress
$payload = "window.MIDAS_SUBMISSIONS_PAYLOAD = $json;`r`nwindow.MIDAS_SUBMISSIONS_META = window.MIDAS_SUBMISSIONS_PAYLOAD.metadata;`r`nwindow.MIDAS_SUBMISSIONS_DATA = window.MIDAS_SUBMISSIONS_PAYLOAD.data;"
[System.IO.File]::WriteAllText($OutFile, $payload, (New-Object System.Text.UTF8Encoding($false)))

$fileSizeMB = [math]::Round((Get-Item $OutFile).Length / 1MB, 2)
Write-Output ("[SUCCESS] Updated midas_submissions_data.js -> rows: {0}, excludedFutureRows: {1}, fileSizeMB: {2}" -f $rows.Count, $futureRows.Count, $fileSizeMB)
