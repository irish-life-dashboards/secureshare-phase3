param(
  [string]$Server = 'WINPRDAF3350',
  [string]$Database = 'AutomateMetrics',
  [string]$OutFile = 'c:/Users/m347/Documents/Power BI Desktop/SecureShare Report/secureshare-phase3/VRDD Appropriations/data/vrdd_appropriations_data.js',
  [string]$StartDate = '2023-10-01',
  [int]$MaxRows = 5000
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# SQL query to extract VRDD appropriations data from AutomateMetrics
# Using same view/columns as MIDAS refresh but filtering for VRDD work types
$sql = @"
WITH src AS (
  SELECT
    Month = CONVERT(varchar(7), TP2V3_DATE_CREATED, 23),
    DateCreated = TRY_CONVERT(date, TP2V3_DATE_CREATED),
    DateCompleted = TRY_CONVERT(date, TP2V3_WI_DATE_COMPLETED),
    SchemeNumber = COALESCE(NULLIF(LTRIM(RTRIM(TP2V3_SCHEME_NO)), ''), NULLIF(LTRIM(RTRIM(TP2V3_WI_LINK_SCHEME_NO)), '')),
    SchemeName = NULLIF(LTRIM(RTRIM(TP2V3_WI_SCHEME_NAME)), ''),
    Broker = NULLIF(LTRIM(RTRIM(TP2V3_WI_BROKER_NAME)), ''),
    WorkContext = NULLIF(LTRIM(RTRIM(TP2V3_WI_CONTEXT)), ''),
    WorkTypeCode = NULLIF(LTRIM(RTRIM(TP2V3_WI_WORKTYPE)), ''),
    WorkTypeDesc = NULLIF(LTRIM(RTRIM(TP2V3_WI_WORKTYPE_DESC)), ''),
    WorkStatus = NULLIF(LTRIM(RTRIM(TP2V3_WI_STATUS)), '')
  FROM dbo.vw_ILCB_CUSTOM_DATA_G360
  WHERE TP2V3_WI_WORKTYPE_DESC LIKE '%DD%'
    OR TP2V3_WI_WORKTYPE LIKE '%VARIABLEDD%'
    OR TP2V3_WI_WORKTYPE LIKE '%CADD%'
), limited AS (
  SELECT TOP (@MaxRows)
    Month,
    DateCreated,
    DateCompleted,
    SchemeNumber,
    SchemeName,
    Broker,
    WorkContext,
    WorkTypeCode,
    WorkTypeDesc,
    WorkStatus
  FROM src
  WHERE DateCreated >= TRY_CONVERT(date, @StartDate)
    AND SchemeNumber IS NOT NULL
  ORDER BY DateCreated DESC, SchemeNumber DESC
)
SELECT
  [Month] = ISNULL(Month, ''),
  [DateCreated] = CONVERT(varchar(10), DateCreated, 103),
  [Reference] = SchemeNumber + '-' + CONVERT(varchar(10), DateCreated, 112),
  [SchemeName] = ISNULL(SchemeName, ''),
  [SchemeNumber] = ISNULL(SchemeNumber, ''),
  [Context] = ISNULL(WorkContext, 'DC'),
  [Amount] = 0,
  [WorkType] = ISNULL(WorkTypeDesc, 'Cash / Direct Debits'),
  [Broker] = ISNULL(Broker, ''),
  [Team] = 'Finance',
  [Source] = 'SecureShare',
  [Status] = CASE 
    WHEN WorkStatus LIKE '%Complete%' THEN 'Complete Work'
    WHEN WorkStatus LIKE '%Pending%' THEN 'Pending'
    WHEN WorkStatus LIKE '%Exception%' THEN 'Exception'
    WHEN WorkStatus LIKE '%Progress%' THEN 'Work In Progress'
    ELSE ISNULL(WorkStatus, 'Pending')
  END,
  [DateCompleted] = CASE WHEN DateCompleted IS NULL THEN 'WIP' ELSE CONVERT(varchar(10), DateCompleted, 103) END,
  [CompletedBy] = '',
  [PendReason] = 'n/a',
  [SystemDate] = CONVERT(varchar(10), GETDATE(), 103),
  [SpVisible] = 'Yes',
  [CbInformed] = 'Yes',
  [ExceptionType] = NULL
FROM limited
ORDER BY DateCreated DESC, SchemeNumber DESC;
"@

# Open connection to AutomateMetrics
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

# Convert rows to VRDD_DATA compatible objects
$rows = foreach ($r in $dt.Rows) {
  [ordered]@{
    'month' = [string]$r['Month']
    'dateCreated' = [string]$r['DateCreated']
    'ref' = [string]$r['Reference']
    'schemeName' = [string]$r['SchemeName']
    'schemeNo' = [string]$r['SchemeNumber']
    'context' = [string]$r['Context']
    'amount' = ([decimal]$r['Amount'])
    'workType' = [string]$r['WorkType']
    'broker' = [string]$r['Broker']
    'team' = [string]$r['Team']
    'source' = [string]$r['Source']
    'status' = [string]$r['Status']
    'dateCompleted' = [string]$r['DateCompleted']
    'completedBy' = [string]$r['CompletedBy']
    'pendReason' = [string]$r['PendReason']
    'effectiveDate' = [string]$r['DateCreated']
    'systemDate' = [string]$r['SystemDate']
    'preappropDate' = [string]$r['DateCreated']
    'spVisible' = [string]$r['SpVisible']
    'cbInformed' = [string]$r['CbInformed']
    'exType' = if ($r['ExceptionType'] -eq [DBNull]::Value) { $null } else { [string]$r['ExceptionType'] }
  }
}

# Convert to JSON and write to file as JavaScript module
$json = $rows | ConvertTo-Json -Depth 4 -Compress
if ($rows.Count -eq 0) {
  $payload = "window.VRDD_DATA_SOURCE='AutomateMetrics';const VRDD_DATA=[];"
} else {
  $payload = "window.VRDD_DATA_SOURCE='AutomateMetrics';const VRDD_DATA=$json;"
}

[System.IO.File]::WriteAllText($OutFile, $payload, (New-Object System.Text.UTF8Encoding($false)))

Write-Output ("Updated vrdd_appropriations_data.js -> rows: {0}, startDate: {1}, maxRows: {2}, server: {3}" -f $rows.Count, $StartDate, $MaxRows, $Server)
