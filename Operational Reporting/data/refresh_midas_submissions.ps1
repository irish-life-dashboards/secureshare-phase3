param(
  [string]$Server = 'WINPRDAF3350',
  [string]$Database = 'AutomateMetrics',
  [string]$OutFile = 'c:/Users/m347/Documents/Power BI Desktop/SecureShare Report/secureshare-phase3/Operational Reporting/midas_submissions_data.js',
  [string]$StartDate = '2023-10-01',
  [int]$MaxRows = 300000
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

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
), limited AS (
  SELECT TOP (@MaxRows)
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
  FROM src
  WHERE DateCreated >= TRY_CONVERT(date, @StartDate)
    AND SchemeNumber IS NOT NULL
  ORDER BY DateCreated DESC, SchemeNumber DESC
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

$json = $rows | ConvertTo-Json -Depth 4 -Compress
$payload = "window.MIDAS_SUBMISSIONS_DATA = $json;"
[System.IO.File]::WriteAllText($OutFile, $payload, (New-Object System.Text.UTF8Encoding($false)))

Write-Output ("Updated midas_submissions_data.js -> rows: {0}, startDate: {1}, maxRows: {2}, server: {3}" -f $rows.Count, $StartDate, $MaxRows, $Server)
