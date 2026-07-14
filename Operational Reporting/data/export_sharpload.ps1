$p = 'C:\Users\m347\Documents\Power BI Desktop\SecureShare Report\secureshare-phase3\Bulk New Entrant Report- look ups.xlsx'
$out = 'C:\Users\m347\Documents\Power BI Desktop\SecureShare Report\secureshare-phase3\sharpload_bulk_data.js'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($p)
$ws = $wb.Worksheets.Item('Export Worksheet')
$ur = $ws.UsedRange
$headers = @()
for ($c = 1; $c -le $ur.Columns.Count; $c++) {
  $headers += ([string]$ur.Cells.Item(1,$c).Text).Trim()
}
$rows = New-Object System.Collections.Generic.List[object]
for ($r = 2; $r -le $ur.Rows.Count; $r++) {
  $midas = ([string]$ur.Cells.Item($r,1).Text).Trim()
  if ([string]::IsNullOrWhiteSpace($midas)) { continue }
  $row = [ordered]@{}
  for ($c = 1; $c -le $ur.Columns.Count; $c++) {
    $header = $headers[$c - 1]
    if ([string]::IsNullOrWhiteSpace($header)) { continue }
    if (-not $row.Contains($header)) {
      $row[$header] = ([string]$ur.Cells.Item($r,$c).Text).Trim()
    }
  }
  $rows.Add($row)
}
$json = $rows | ConvertTo-Json -Depth 4 -Compress
Set-Content -Path $out -Value ("window.SHARPLOAD_BULK_DATA = " + $json + ";") -Encoding UTF8
Write-Host "Wrote $($rows.Count) rows to $out"
$wb.Close($false)
$xl.Quit()
