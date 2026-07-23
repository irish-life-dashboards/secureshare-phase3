param()
$ErrorActionPreference = 'Stop'
$noBom   = New-Object System.Text.UTF8Encoding $false
$base    = "C:\Users\m347\Documents\Power BI Desktop\SecureShare Report\secureshare-phase3\Mothly Engagement\PBIP"
$tmdl    = "$base\Monthly Engagement.SemanticModel\definition\tables\UshurHub.tmdl"
$pgRoot  = "$base\Monthly Engagement.Report\definition\pages"

if (Get-Process -Name "PBIDesktop" -EA SilentlyContinue) {
    Write-Error "Close PBI Desktop first."; exit 1
}

# ---------------------------------------------------------
# CSS string (no double-quotes inside — safe as DAX literal)
# ---------------------------------------------------------
$CSS = 'body{margin:0;padding:10px 12px;font-family:Segoe UI,sans-serif;background:#f4f7fb;color:#25323b;font-size:13px}.topbar{background:linear-gradient(96deg,#3f468e,#5c61ac 45%,#4f88ba);color:#fff;border-radius:12px;padding:10px 14px;margin-bottom:10px;display:flex;justify-content:space-between;align-items:center}.ttl{font-size:.92rem;font-weight:800}.sub{font-size:.76rem;opacity:.9}.g4{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:10px}.g2{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:10px}.card{background:#fff;border:1px solid #dce3ee;border-radius:12px;padding:12px;box-shadow:0 4px 16px rgba(37,50,59,.07)}.kl{font-size:.69rem;text-transform:uppercase;color:#67747f;font-weight:700}.kv{font-size:1.5rem;font-weight:800;margin-top:3px}.km{font-size:.73rem;color:#5f6b75;margin-top:3px}h3{margin:0 0 8px;font-size:.85rem;font-weight:800;color:#2f3f49}table{width:100%;border-collapse:collapse;font-size:.74rem}th{background:linear-gradient(90deg,#5c61ac,#4f86b8);color:#fff;padding:6px 8px;text-align:left;position:sticky;top:0;font-size:.69rem;text-transform:uppercase}td{padding:6px 8px;border-bottom:1px solid #ebeff6;vertical-align:top}.wrap{border:1px solid #dce3ee;border-radius:8px;overflow:auto;background:#fff}.flag{color:#c33d4b;font-weight:700}.br{display:flex;align-items:center;gap:6px;margin:3px 0}.bl{width:70px;font-size:.69rem;text-align:right;color:#67747f;flex-shrink:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.bt{flex:1;background:#eef1f8;border-radius:3px;height:13px}.bf{height:13px;border-radius:3px;min-width:2px}.bv{width:52px;font-size:.7rem;font-weight:700;padding-left:4px}.fstep{border:1px solid #dce3ee;border-radius:8px;padding:8px 10px;background:#fff;margin:3px 0}.fstep strong{font-size:.95rem;color:#2f3f49}.fmeta{font-size:.73rem;color:#5f6b75}'

# ---------------------------------------------------------
# DAX expressions (single-quoted PS here-strings: no $ expansion)
# ---------------------------------------------------------
$DAX1 = @'
VAR _css = "CSSPLACEHOLDER"
VAR _tot = [Total Submissions]
VAR _cmp = [Completed Submissions]
VAR _rt  = [Completion Rate]
VAR _org = [Active Organizations]
VAR _err = [Total Errors]
VAR _ert = [Error Rate]
VAR _cc  = IF(_rt >= 0.8, "#14763d", IF(_rt >= 0.6, "#5c61ac", "#e03e52"))
VAR _ec  = IF(_ert > 0.05, "#e03e52", IF(_ert > 0.02, "#fab812", "#14763d"))
VAR _all = CALCULATETABLE(SUMMARIZE(UshurHub, UshurHub[MonthKey], UshurHub[MonthYear]), ALL(UshurHub[MonthKey]), ALL(UshurHub[MonthYear]))
VAR _mxK = MAXX(_all, UshurHub[MonthKey])
VAR _pvK = MAXX(FILTER(_all, UshurHub[MonthKey] < _mxK), UshurHub[MonthKey])
VAR _mxY = MAXX(FILTER(_all, UshurHub[MonthKey] = _mxK), UshurHub[MonthYear])
VAR _pvY = IF(ISBLANK(_pvK), "(no prior month)", MAXX(FILTER(_all, UshurHub[MonthKey] = _pvK), UshurHub[MonthYear]))
VAR _wBase = ADDCOLUMNS(CALCULATETABLE(SUMMARIZE(UshurHub, UshurHub[Organization]), ALL(UshurHub[MonthKey]), ALL(UshurHub[MonthYear])), "_pv", CALCULATE([Total Submissions], ALL(UshurHub[MonthKey]), ALL(UshurHub[MonthYear]), UshurHub[MonthKey] = _pvK), "_cv", CALCULATE([Total Submissions], ALL(UshurHub[MonthKey]), ALL(UshurHub[MonthYear]), UshurHub[MonthKey] = _mxK), "_pc", CALCULATE([Completion Rate], ALL(UshurHub[MonthKey]), ALL(UshurHub[MonthYear]), UshurHub[MonthKey] = _pvK), "_cc", CALCULATE([Completion Rate], ALL(UshurHub[MonthKey]), ALL(UshurHub[MonthYear]), UshurHub[MonthKey] = _mxK))
VAR _wf  = FILTER(_wBase, [_pv] >= 3 && (DIVIDE([_cv] - [_pv], [_pv]) <= -0.5 || [_cc] - [_pc] <= -0.15))
VAR _wr  = CONCATENATEX(TOPN(15, _wf, [_pv], DESC), "<tr><td class='flag'>" & UshurHub[Organization] & "</td><td>" & FORMAT([_pv], "#,0") & "</td><td>" & FORMAT([_cv], "#,0") & "</td><td>" & FORMAT([_pc], "0.0%") & "</td><td>" & FORMAT([_cc], "0.0%") & "</td><td>" & IF(DIVIDE([_cv]-[_pv],[_pv]) <= -0.5 && [_cc]-[_pc] <= -0.15, "Vol+Comp", IF(DIVIDE([_cv]-[_pv],[_pv]) <= -0.5, "Vol drop", "Comp drop")) & "</td></tr>", "")
VAR _wh  = IF(COUNTROWS(_wf) = 0, "<tr><td colspan='6' style='text-align:center;color:#67747f;padding:12px'>No flagged organisations</td></tr>", _wr)
VAR _tBase = ADDCOLUMNS(SUMMARIZE(UshurHub, UshurHub[Organization]), "_ts", [Total Submissions], "_cs", [Completed Submissions], "_cr", [Completion Rate])
VAR _tr  = CONCATENATEX(TOPN(20, _tBase, [_ts], DESC), "<tr><td>" & UshurHub[Organization] & "</td><td>" & FORMAT([_ts], "#,0") & "</td><td>" & FORMAT([_cs], "#,0") & "</td><td>" & FORMAT([_cr], "0.0%") & "</td></tr>", "", [_ts], DESC)
RETURN "<!DOCTYPE html><html><head><meta charset='utf-8'/><style>" & _css & "</style></head><body><div class='topbar'><div><div class='ttl'>SecureShare - Ushur Hub</div><div class='sub'>Monthly Engagement - Summary</div></div></div><div class='g4'><div class='card'><div class='kl'>Total Submissions</div><div class='kv'>" & FORMAT(_tot, "#,0") & "</div><div class='km'>Unique by reference</div></div><div class='card'><div class='kl'>Completion Rate</div><div class='kv' style='color:" & _cc & "'>" & FORMAT(_rt, "0.0%") & "</div><div class='km'>" & FORMAT(_cmp, "#,0") & " completed</div></div><div class='card'><div class='kl'>Active Organizations</div><div class='kv'>" & FORMAT(_org, "#,0") & "</div><div class='km'>Submitted at least once</div></div><div class='card'><div class='kl'>Error Rate</div><div class='kv' style='color:" & _ec & "'>" & FORMAT(_ert, "0.0%") & "</div><div class='km'>" & FORMAT(_err, "#,0") & " errors</div></div></div><div class='card' style='margin-bottom:10px'><h3>Watch List - " & _pvY & " to " & _mxY & "</h3><div class='wrap' style='max-height:200px'><table><thead><tr><th>Organization</th><th>Prior Vol</th><th>Curr Vol</th><th>Prior%</th><th>Curr%</th><th>Flag</th></tr></thead><tbody>" & _wh & "</tbody></table></div></div><div class='card'><h3>Top 20 Organizations</h3><div class='wrap' style='max-height:260px'><table><thead><tr><th>Organization</th><th>Submissions</th><th>Completed</th><th>Comp Rate</th></tr></thead><tbody>" & _tr & "</tbody></table></div></div></body></html>"
'@

$DAX2 = @'
VAR _css = "CSSPLACEHOLDER"
VAR _mb  = ADDCOLUMNS(SUMMARIZE(UshurHub, UshurHub[MonthKey], UshurHub[MonthYear]), "_ts", [Total Submissions], "_cs", [Completed Submissions], "_cr", [Completion Rate], "_ao", [Active Organizations], "_te", [Total Errors], "_er", [Error Rate])
VAR _mxTS = MAXX(_mb, [_ts])
VAR _mxAO = MAXX(_mb, [_ao])
VAR _bTS = CONCATENATEX(_mb, "<div class='br'><div class='bl'>" & UshurHub[MonthYear] & "</div><div class='bt'><div class='bf' style='width:" & FORMAT(MIN(DIVIDE([_ts],_mxTS)*100,100),"0") & "%;background:#5c61ac'></div></div><div class='bv'>" & FORMAT([_ts],"#,0") & "</div></div>", "", UshurHub[MonthKey], ASC)
VAR _bCR = CONCATENATEX(_mb, "<div class='br'><div class='bl'>" & UshurHub[MonthYear] & "</div><div class='bt'><div class='bf' style='width:" & FORMAT([_cr]*100,"0") & "%;background:#55c2b6'></div></div><div class='bv'>" & FORMAT([_cr],"0.0%") & "</div></div>", "", UshurHub[MonthKey], ASC)
VAR _bAO = CONCATENATEX(_mb, "<div class='br'><div class='bl'>" & UshurHub[MonthYear] & "</div><div class='bt'><div class='bf' style='width:" & FORMAT(MIN(DIVIDE([_ao],_mxAO)*100,100),"0") & "%;background:#00b1d9'></div></div><div class='bv'>" & FORMAT([_ao],"#,0") & "</div></div>", "", UshurHub[MonthKey], ASC)
VAR _bER = CONCATENATEX(_mb, "<div class='br'><div class='bl'>" & UshurHub[MonthYear] & "</div><div class='bt'><div class='bf' style='width:" & FORMAT([_er]*100,"0") & "%;background:#e03e52'></div></div><div class='bv'>" & FORMAT([_er],"0.0%") & "</div></div>", "", UshurHub[MonthKey], ASC)
VAR _tr  = CONCATENATEX(_mb, "<tr><td>" & UshurHub[MonthYear] & "</td><td>" & FORMAT([_ts],"#,0") & "</td><td>" & FORMAT([_cs],"#,0") & "</td><td>" & FORMAT([_cr],"0.0%") & "</td><td>" & FORMAT([_ao],"#,0") & "</td><td>" & FORMAT([_te],"#,0") & "</td><td>" & FORMAT([_er],"0.0%") & "</td></tr>", "", UshurHub[MonthKey], ASC)
RETURN "<!DOCTYPE html><html><head><meta charset='utf-8'/><style>" & _css & "</style></head><body><div class='topbar'><div><div class='ttl'>SecureShare - Ushur Hub</div><div class='sub'>Monthly Engagement - Monthly Trends</div></div></div><div class='g2'><div class='card'><h3>Monthly Total Submissions</h3>" & _bTS & "</div><div class='card'><h3>Monthly Completion Rate</h3>" & _bCR & "</div><div class='card'><h3>Monthly Active Organizations</h3>" & _bAO & "</div><div class='card'><h3>Monthly Error Rate</h3>" & _bER & "</div><div class='card' style='grid-column:span 2'><h3>Monthly Trend Data</h3><div class='wrap'><table><thead><tr><th>Month</th><th>Submissions</th><th>Completed</th><th>Comp Rate</th><th>Active Orgs</th><th>Errors</th><th>Error Rate</th></tr></thead><tbody>" & _tr & "</tbody></table></div></div></div></body></html>"
'@

$DAX3 = @'
VAR _css = "CSSPLACEHOLDER"
VAR _qb  = ADDCOLUMNS(SUMMARIZE(UshurHub, UshurHub[YearQuarter]), "_ts", [Total Submissions], "_cs", [Completed Submissions], "_cr", [Completion Rate], "_ao", [Active Organizations], "_te", [Total Errors], "_er", [Error Rate])
VAR _mxQTS = MAXX(_qb, [_ts])
VAR _bQ  = CONCATENATEX(_qb, "<div class='br'><div class='bl'>" & UshurHub[YearQuarter] & "</div><div class='bt'><div class='bf' style='width:" & FORMAT(MIN(DIVIDE([_ts],_mxQTS)*100,100),"0") & "%;background:#5c61ac'></div></div><div class='bv'>" & FORMAT([_ts],"#,0") & "</div></div>", "", UshurHub[YearQuarter], ASC)
VAR _bCR = CONCATENATEX(_qb, "<div class='br'><div class='bl'>" & UshurHub[YearQuarter] & "</div><div class='bt'><div class='bf' style='width:" & FORMAT([_cr]*100,"0") & "%;background:#55c2b6'></div></div><div class='bv'>" & FORMAT([_cr],"0.0%") & "</div></div>", "", UshurHub[YearQuarter], ASC)
VAR _tr  = CONCATENATEX(_qb, "<tr><td>" & UshurHub[YearQuarter] & "</td><td>" & FORMAT([_ts],"#,0") & "</td><td>" & FORMAT([_cs],"#,0") & "</td><td>" & FORMAT([_cr],"0.0%") & "</td><td>" & FORMAT([_ao],"#,0") & "</td><td>" & FORMAT([_te],"#,0") & "</td><td>" & FORMAT([_er],"0.0%") & "</td></tr>", "", UshurHub[YearQuarter], ASC)
RETURN "<!DOCTYPE html><html><head><meta charset='utf-8'/><style>" & _css & "</style></head><body><div class='topbar'><div><div class='ttl'>SecureShare - Ushur Hub</div><div class='sub'>Monthly Engagement - Quarterly Trends</div></div></div><div class='g2'><div class='card'><h3>Quarterly Submissions</h3>" & _bQ & "</div><div class='card'><h3>Quarterly Completion Rate</h3>" & _bCR & "</div><div class='card' style='grid-column:span 2'><h3>Quarterly Trend Data</h3><div class='wrap'><table><thead><tr><th>Quarter</th><th>Submissions</th><th>Completed</th><th>Comp Rate</th><th>Active Orgs</th><th>Errors</th><th>Error Rate</th></tr></thead><tbody>" & _tr & "</tbody></table></div></div></div></body></html>"
'@

$DAX4 = @'
VAR _css = "CSSPLACEHOLDER"
VAR _ucB = ADDCOLUMNS(SUMMARIZE(UshurHub, UshurHub[UseCase]), "_ts", [Total Submissions])
VAR _mxUC = MAXX(_ucB, [_ts])
VAR _bUC = CONCATENATEX(TOPN(15,_ucB,[_ts],DESC), "<div class='br'><div class='bl'>" & UshurHub[UseCase] & "</div><div class='bt'><div class='bf' style='width:" & FORMAT(MIN(DIVIDE([_ts],_mxUC)*100,100),"0") & "%;background:#5c61ac'></div></div><div class='bv'>" & FORMAT([_ts],"#,0") & "</div></div>", "", [_ts], DESC)
VAR _stB = ADDCOLUMNS(SUMMARIZE(UshurHub, UshurHub[Status]), "_ts", [Total Submissions])
VAR _mxST = MAXX(_stB, [_ts])
VAR _bST = CONCATENATEX(_stB, "<div class='br'><div class='bl'>" & UshurHub[Status] & "</div><div class='bt'><div class='bf' style='width:" & FORMAT(MIN(DIVIDE([_ts],_mxST)*100,100),"0") & "%;background:#55c2b6'></div></div><div class='bv'>" & FORMAT([_ts],"#,0") & "</div></div>", "", [_ts], DESC)
VAR _ttl = [Total Submissions]
VAR _cmp = [Completed Submissions]
VAR _can = [Cancelled Submissions]
VAR _err = [Total Errors]
VAR _fun = "<div class='fstep'><strong>" & FORMAT(_ttl,"#,0") & "</strong><div class='fmeta'>Total Received</div></div><div class='fstep'><strong style='color:#5c61ac'>" & FORMAT(_cmp+_err+_can,"#,0") & " (" & FORMAT(DIVIDE(_cmp+_err+_can,_ttl),"0.0%") & ")</strong><div class='fmeta'>Processed</div></div><div class='fstep'><strong style='color:#14763d'>" & FORMAT(_cmp,"#,0") & " (" & FORMAT(DIVIDE(_cmp,_ttl),"0.0%") & ")</strong><div class='fmeta'>Completed</div></div><div class='fstep'><strong style='color:#e03e52'>" & FORMAT(_err,"#,0") & " (" & FORMAT(DIVIDE(_err,_ttl),"0.0%") & ")</strong><div class='fmeta'>Errors</div></div>"
VAR _mBase = ADDCOLUMNS(SUMMARIZE(UshurHub, UshurHub[UseCase]), "_tot", [Total Submissions], "_can", CALCULATE([Cancelled Submissions]), "_cmp", CALCULATE([Completed Submissions]), "_err", CALCULATE([Total Errors]), "_rec", CALCULATE([Received Submissions]), "_sub", CALCULATE([Submitted Submissions]))
VAR _mr  = CONCATENATEX(TOPN(15,_mBase,[_tot],DESC), "<tr><td>" & UshurHub[UseCase] & "</td><td>" & FORMAT([_can],"#,0") & "</td><td>" & FORMAT([_cmp],"#,0") & "</td><td>" & FORMAT([_err],"#,0") & "</td><td>" & FORMAT([_rec],"#,0") & "</td><td>" & FORMAT([_sub],"#,0") & "</td><td>" & FORMAT([_tot],"#,0") & "</td></tr>", "", [_tot], DESC)
VAR _eBase = ADDCOLUMNS(SUMMARIZE(UshurHub, UshurHub[UseCase]), "_te", CALCULATE([Total Errors]))
VAR _er  = CONCATENATEX(TOPN(10,FILTER(_eBase,[_te]>0),[_te],DESC), "<tr><td>" & UshurHub[UseCase] & "</td><td>" & FORMAT([_te],"#,0") & "</td><td>" & FORMAT(DIVIDE([_te],_err),"0.0%") & " of all</td></tr>", "", [_te], DESC)
VAR _eH  = IF(_err=0,"<tr><td colspan='3' style='text-align:center;color:#67747f;padding:10px'>No errors in selection</td></tr>",_er)
RETURN "<!DOCTYPE html><html><head><meta charset='utf-8'/><style>" & _css & "</style></head><body><div class='topbar'><div><div class='ttl'>SecureShare - Ushur Hub</div><div class='sub'>Monthly Engagement - Detail</div></div></div><div class='g2'><div class='card'><h3>Submissions by Use Case</h3>" & _bUC & "</div><div class='card'><h3>Status Breakdown</h3>" & _bST & "</div><div class='card'><h3>Completion Funnel</h3>" & _fun & "</div><div class='card'><h3>Error Concentration</h3><div class='wrap' style='max-height:220px'><table><thead><tr><th>Use Case</th><th>Errors</th><th>Share</th></tr></thead><tbody>" & _eH & "</tbody></table></div></div><div class='card' style='grid-column:span 2'><h3>Use Case x Status Matrix</h3><div class='wrap'><table><thead><tr><th>Use Case</th><th>Cancelled</th><th>Completed</th><th>Error</th><th>Received</th><th>Submitted</th><th>Total</th></tr></thead><tbody>" & _mr & "</tbody></table></div></div></div></body></html>"
'@

# Substitute the CSS placeholder
$DAX1 = $DAX1 -replace 'CSSPLACEHOLDER', $CSS
$DAX2 = $DAX2 -replace 'CSSPLACEHOLDER', $CSS
$DAX3 = $DAX3 -replace 'CSSPLACEHOLDER', $CSS
$DAX4 = $DAX4 -replace 'CSSPLACEHOLDER', $CSS

# ---------------------------------------------------------
# Build TMDL measure block: 3-tab indent for expression, 2-tab for props
# ---------------------------------------------------------
function Mk-Block([string]$name, [string]$dax, [string]$guid) {
    $lines = ($dax -replace "`r`n","`n" -replace "`r","`n") -split "`n" `
             | Where-Object { $_.Trim() -ne "" } `
             | ForEach-Object { "`t`t`t$_" }
    $body = $lines -join "`n"
    return "`n`tmeasure '$name' =`n$body`n`t`tformatString: `n`t`tdisplayFolder: HTML Pages`n`t`tlineageTag: $guid"
}

$blocks = (Mk-Block "HTML Summary"          $DAX1 "3f8e7a6d-1c2b-4a59-8f0e-9b7d6c5a4e3f") +
          (Mk-Block "HTML Monthly Trends"   $DAX2 "7c4b9d2e-5a1f-4862-b390-e1d8c6a5f2b4") +
          (Mk-Block "HTML Quarterly Trends" $DAX3 "9e6d4b8a-3f5c-4271-a0b9-d7e2c4f8a1e6") +
          (Mk-Block "HTML Detail"           $DAX4 "1a5f8c3e-9b7d-4e62-8a4f-c2b6d0e9f3a7")

# ---------------------------------------------------------
# 1. Patch TMDL
# ---------------------------------------------------------
$tmdlText = [System.IO.File]::ReadAllText($tmdl, [System.Text.Encoding]::UTF8) -replace "`r`n","`n"

if ($tmdlText -like "*measure 'HTML Summary'*") {
    Write-Host "[SKIP] HTML measures already in TMDL" -ForegroundColor Yellow
} else {
    $anchor = "`n`tcolumn Organization"
    $idx    = $tmdlText.IndexOf($anchor)
    if ($idx -lt 0) { Write-Error "Cannot find anchor in TMDL"; exit 1 }
    $tmdlText = $tmdlText.Insert($idx, $blocks)
    [System.IO.File]::WriteAllText($tmdl, $tmdlText, $noBom)
    Write-Host "[OK] TMDL: 4 HTML measures inserted" -ForegroundColor Green
}

# ---------------------------------------------------------
# 2. Create HTML Content visual on each page
# (replaces any non-slicer visuals)
# ---------------------------------------------------------
$vType = "htmlContent443BE3AD55E043BF878BED274D3A6855"
$schema = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.3.0/schema.json"

$pages = @(
    @{ guid="965f4a0e8701b6becab8"; measure="HTML Summary";          vn="html_p1_summary"   }
    @{ guid="a2b3c4d5e6f7a8b9c0d1"; measure="HTML Monthly Trends";   vn="html_p2_monthly"   }
    @{ guid="b3c4d5e6f7a8b9c0d1e2"; measure="HTML Quarterly Trends"; vn="html_p3_quarterly" }
    @{ guid="c4d5e6f7a8b9c0d1e2f3"; measure="HTML Detail";           vn="html_p4_detail"    }
)

foreach ($p in $pages) {
    $vRoot = Join-Path $pgRoot "$($p.guid)\visuals"

    # Remove non-slicer folders
    Get-ChildItem $vRoot -Directory -EA SilentlyContinue |
        Where-Object { $_.Name -notlike "slicer_*" -and $_.Name -ne $p.vn } |
        ForEach-Object { Remove-Item $_.FullName -Recurse -Force; Write-Host "  Removed $($_.Name)" }

    # Write visual.json
    $vDir  = Join-Path $vRoot $p.vn
    $vFile = Join-Path $vDir "visual.json"
    if (-not (Test-Path $vDir)) { New-Item -ItemType Directory -Path $vDir -Force | Out-Null }

    $vJson = @"
{
  "`$schema": "$schema",
  "name": "$($p.vn)",
  "position": { "x": 10, "y": 100, "z": 1000, "width": 1260, "height": 612, "tabOrder": 10000 },
  "visual": {
    "visualType": "$vType",
    "drillFilterOtherVisuals": true,
    "query": {
      "queryState": {
        "Values": {
          "projections": [
            {
              "field": {
                "Measure": {
                  "Expression": { "SourceRef": { "Entity": "UshurHub" } },
                  "Property": "$($p.measure)"
                }
              },
              "queryRef": "UshurHub.$($p.measure)"
            }
          ]
        }
      }
    }
  }
}
"@
    [System.IO.File]::WriteAllText($vFile, $vJson, $noBom)
    Write-Host "[OK] $($p.vn)  ->  '$($p.measure)'" -ForegroundColor Green
}

Write-Host ""
Write-Host "=== DONE ===" -ForegroundColor Cyan
Write-Host "Open Monthly Engagement.pbip in PBI Desktop." -ForegroundColor Cyan
