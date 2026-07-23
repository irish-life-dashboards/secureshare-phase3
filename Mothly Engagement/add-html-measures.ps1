# ============================================================
# add-html-measures.ps1
# Run AFTER saving (Ctrl+S) and closing Power BI Desktop.
# Adds 4 HTML Page measures to UshurHub.tmdl, adds
# publicCustomVisuals to report.json, and replaces all
# non-slicer visuals on each page with an HTML Content visual.
# ============================================================
$ErrorActionPreference = 'Stop'
$noBom = New-Object System.Text.UTF8Encoding $false

$base      = "C:\Users\m347\Documents\Power BI Desktop\SecureShare Report\secureshare-phase3\Mothly Engagement\PBIP"
$tmdlPath  = "$base\Monthly Engagement.SemanticModel\definition\tables\UshurHub.tmdl"
$rptJson   = "$base\Monthly Engagement.Report\definition\report.json"
$pagesRoot = "$base\Monthly Engagement.Report\definition\pages"

# Safety check
$pbi = Get-Process -Name "PBIDesktop" -ErrorAction SilentlyContinue
if ($pbi) {
    Write-Host "ERROR: Power BI Desktop is still running. Save (Ctrl+S) and close it first." -ForegroundColor Red
    exit 1
}

# GUIDs for new measures
$g1 = [System.Guid]::NewGuid().ToString()
$g2 = [System.Guid]::NewGuid().ToString()
$g3 = [System.Guid]::NewGuid().ToString()
$g4 = [System.Guid]::NewGuid().ToString()

# Shared CSS (single-line, no double quotes inside, all HTML attrs use single quotes)
$css = "body{margin:0;padding:10px 12px;font-family:Segoe UI,sans-serif;background:#f4f7fb;color:#25323b;font-size:13px}.topbar{background:linear-gradient(96deg,#3f468e,#5c61ac 45%,#4f88ba);color:#fff;border-radius:12px;padding:10px 14px;margin-bottom:10px;display:flex;justify-content:space-between;align-items:center}.ttl{font-size:.92rem;font-weight:800}.sub{font-size:.76rem;opacity:.9}.g4{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:10px}.g2{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:10px}.card{background:#fff;border:1px solid #dce3ee;border-radius:12px;padding:12px;box-shadow:0 4px 16px rgba(37,50,59,.07)}.kl{font-size:.69rem;text-transform:uppercase;color:#67747f;font-weight:700;letter-spacing:.3px}.kv{font-size:1.5rem;font-weight:800;margin-top:3px}.km{font-size:.73rem;color:#5f6b75;margin-top:3px}h3{margin:0 0 8px;font-size:.85rem;font-weight:800;color:#2f3f49}table{width:100%;border-collapse:collapse;font-size:.74rem}th{background:linear-gradient(90deg,#5c61ac,#4f86b8);color:#fff;padding:6px 8px;text-align:left;position:sticky;top:0;font-size:.69rem;text-transform:uppercase}td{padding:6px 8px;border-bottom:1px solid #ebeff6;vertical-align:top}tr:last-child td{border-bottom:none}.wrap{border:1px solid #dce3ee;border-radius:8px;overflow:auto;background:#fff}.flag{color:#c33d4b;font-weight:700}.br{display:flex;align-items:center;gap:6px;margin:3px 0}.bl{width:70px;font-size:.69rem;text-align:right;color:#67747f;flex-shrink:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.bt{flex:1;background:#eef1f8;border-radius:3px;height:13px}.bf{height:13px;border-radius:3px;min-width:2px}.bv{width:52px;font-size:.7rem;font-weight:700;padding-left:4px}.fstep{border:1px solid #dce3ee;border-radius:8px;padding:8px 10px;background:#fff;margin:3px 0}.fstep strong{font-size:.95rem;color:#2f3f49}.fmeta{font-size:.73rem;color:#5f6b75}"

# ============================================================
# DAX expressions — multiline here-strings
# Note: $css is substituted by PowerShell (no $ in CSS content)
# All HTML attribute values use single quotes
# ============================================================

$dax1 = @"
VAR _css = "$css"
VAR _total = [Total Submissions]
VAR _comp = [Completed Submissions]
VAR _rate = [Completion Rate]
VAR _orgs = [Active Organizations]
VAR _errors = [Total Errors]
VAR _errrate = [Error Rate]
VAR _compColor = IF(_rate >= 0.8, "#14763d", IF(_rate >= 0.6, "#5c61ac", "#e03e52"))
VAR _errColor = IF(_errrate > 0.05, "#e03e52", IF(_errrate > 0.02, "#fab812", "#14763d"))
VAR _mkAll = CALCULATETABLE(SUMMARIZE(UshurHub, UshurHub[MonthKey], UshurHub[MonthYear]), ALL(UshurHub[MonthKey]), ALL(UshurHub[MonthYear]))
VAR _maxMK = MAXX(_mkAll, UshurHub[MonthKey])
VAR _prevMK = MAXX(FILTER(_mkAll, UshurHub[MonthKey] < _maxMK), UshurHub[MonthKey])
VAR _maxMY = MAXX(FILTER(_mkAll, UshurHub[MonthKey] = _maxMK), UshurHub[MonthYear])
VAR _prevMY = COALESCE(MAXX(FILTER(_mkAll, UshurHub[MonthKey] = _prevMK), UshurHub[MonthYear]), "(no prior month)")
VAR _watchBase = ADDCOLUMNS(CALCULATETABLE(SUMMARIZE(UshurHub, UshurHub[Organization]), ALL(UshurHub[MonthKey]), ALL(UshurHub[MonthYear])), "_pv", CALCULATE([Total Submissions], ALL(UshurHub[MonthKey]), ALL(UshurHub[MonthYear]), UshurHub[MonthKey] = _prevMK), "_cv", CALCULATE([Total Submissions], ALL(UshurHub[MonthKey]), ALL(UshurHub[MonthYear]), UshurHub[MonthKey] = _maxMK), "_pc", CALCULATE([Completion Rate], ALL(UshurHub[MonthKey]), ALL(UshurHub[MonthYear]), UshurHub[MonthKey] = _prevMK), "_cc", CALCULATE([Completion Rate], ALL(UshurHub[MonthKey]), ALL(UshurHub[MonthYear]), UshurHub[MonthKey] = _maxMK))
VAR _wf = FILTER(_watchBase, [_pv] >= 3 && (DIVIDE([_cv] - [_pv], [_pv]) <= -0.5 || [_cc] - [_pc] <= -0.15))
VAR _wr = CONCATENATEX(TOPN(15, _wf, [_pv], DESC), "<tr><td class='flag'>" & UshurHub[Organization] & "</td><td>" & FORMAT([_pv], "#,0") & "</td><td>" & FORMAT([_cv], "#,0") & "</td><td>" & FORMAT([_pc], "0.0%") & "</td><td>" & FORMAT([_cc], "0.0%") & "</td><td>" & IF(DIVIDE([_cv]-[_pv],[_pv]) <= -0.5 && [_cc]-[_pc] <= -0.15, "Vol+Comp", IF(DIVIDE([_cv]-[_pv],[_pv]) <= -0.5, "Vol drop", "Comp drop")) & "</td></tr>", "")
VAR _wb = IF(COUNTROWS(_wf) = 0, "<tr><td colspan='6' style='text-align:center;color:#67747f;padding:12px'>No flagged organisations</td></tr>", _wr)
VAR _tb = ADDCOLUMNS(SUMMARIZE(UshurHub, UshurHub[Organization]), "_ts", [Total Submissions], "_cs", [Completed Submissions], "_cr", [Completion Rate])
VAR _tr = CONCATENATEX(TOPN(20, _tb, [_ts], DESC), "<tr><td>" & UshurHub[Organization] & "</td><td>" & FORMAT([_ts], "#,0") & "</td><td>" & FORMAT([_cs], "#,0") & "</td><td>" & FORMAT([_cr], "0.0%") & "</td></tr>", "", [_ts], DESC)
RETURN "<!DOCTYPE html><html><head><meta charset='utf-8'/><style>" & _css & "</style></head><body><div class='topbar'><div><div class='ttl'>SecureShare - Ushur Hub</div><div class='sub'>Monthly Engagement - Summary</div></div></div><div class='g4'><div class='card'><div class='kl'>Total Submissions</div><div class='kv'>" & FORMAT(_total, "#,0") & "</div><div class='km'>Unique by reference</div></div><div class='card'><div class='kl'>Completion Rate</div><div class='kv' style='color:" & _compColor & "'>" & FORMAT(_rate, "0.0%") & "</div><div class='km'>" & FORMAT(_comp, "#,0") & " completed</div></div><div class='card'><div class='kl'>Active Organizations</div><div class='kv'>" & FORMAT(_orgs, "#,0") & "</div><div class='km'>Submitted at least once</div></div><div class='card'><div class='kl'>Error Rate</div><div class='kv' style='color:" & _errColor & "'>" & FORMAT(_errrate, "0.0%") & "</div><div class='km'>" & FORMAT(_errors, "#,0") & " errors</div></div></div><div class='card' style='margin-bottom:10px'><h3>Watch List - " & _prevMY & " to " & _maxMY & "</h3><div class='wrap' style='max-height:200px'><table><thead><tr><th>Organization</th><th>Prior Vol</th><th>Curr Vol</th><th>Prior Comp%</th><th>Curr Comp%</th><th>Flag</th></tr></thead><tbody>" & _wb & "</tbody></table></div></div><div class='card'><h3>Top 20 Organizations</h3><div class='wrap' style='max-height:260px'><table><thead><tr><th>Organization</th><th>Submissions</th><th>Completed</th><th>Completion Rate</th></tr></thead><tbody>" & _tr & "</tbody></table></div></div></body></html>"
"@

$dax2 = @"
VAR _css = "$css"
VAR _mb = ADDCOLUMNS(SUMMARIZE(UshurHub, UshurHub[MonthKey], UshurHub[MonthYear]), "_ts", [Total Submissions], "_cs", [Completed Submissions], "_cr", [Completion Rate], "_ao", [Active Organizations], "_te", [Total Errors], "_er", [Error Rate])
VAR _maxTS = MAXX(_mb, [_ts])
VAR _maxAO = MAXX(_mb, [_ao])
VAR _bTS = CONCATENATEX(_mb, "<div class='br'><div class='bl'>" & UshurHub[MonthYear] & "</div><div class='bt'><div class='bf' style='width:" & FORMAT(MIN(DIVIDE([_ts], _maxTS) * 100, 100), "0") & "%;background:#5c61ac'></div></div><div class='bv'>" & FORMAT([_ts], "#,0") & "</div></div>", "", UshurHub[MonthKey], ASC)
VAR _bCR = CONCATENATEX(_mb, "<div class='br'><div class='bl'>" & UshurHub[MonthYear] & "</div><div class='bt'><div class='bf' style='width:" & FORMAT([_cr] * 100, "0") & "%;background:#55c2b6'></div></div><div class='bv'>" & FORMAT([_cr], "0.0%") & "</div></div>", "", UshurHub[MonthKey], ASC)
VAR _bAO = CONCATENATEX(_mb, "<div class='br'><div class='bl'>" & UshurHub[MonthYear] & "</div><div class='bt'><div class='bf' style='width:" & FORMAT(MIN(DIVIDE([_ao], _maxAO) * 100, 100), "0") & "%;background:#00b1d9'></div></div><div class='bv'>" & FORMAT([_ao], "#,0") & "</div></div>", "", UshurHub[MonthKey], ASC)
VAR _bER = CONCATENATEX(_mb, "<div class='br'><div class='bl'>" & UshurHub[MonthYear] & "</div><div class='bt'><div class='bf' style='width:" & FORMAT([_er] * 100, "0") & "%;background:#e03e52'></div></div><div class='bv'>" & FORMAT([_er], "0.0%") & "</div></div>", "", UshurHub[MonthKey], ASC)
VAR _trows = CONCATENATEX(_mb, "<tr><td>" & UshurHub[MonthYear] & "</td><td>" & FORMAT([_ts], "#,0") & "</td><td>" & FORMAT([_cs], "#,0") & "</td><td>" & FORMAT([_cr], "0.0%") & "</td><td>" & FORMAT([_ao], "#,0") & "</td><td>" & FORMAT([_te], "#,0") & "</td><td>" & FORMAT([_er], "0.0%") & "</td></tr>", "", UshurHub[MonthKey], ASC)
RETURN "<!DOCTYPE html><html><head><meta charset='utf-8'/><style>" & _css & "</style></head><body><div class='topbar'><div><div class='ttl'>SecureShare - Ushur Hub</div><div class='sub'>Monthly Engagement - Monthly Trends</div></div></div><div class='g2'><div class='card'><h3>Monthly Total Submissions</h3>" & _bTS & "</div><div class='card'><h3>Monthly Completion Rate</h3>" & _bCR & "</div><div class='card'><h3>Monthly Active Organizations</h3>" & _bAO & "</div><div class='card'><h3>Monthly Error Rate</h3>" & _bER & "</div><div class='card' style='grid-column:span 2'><h3>Monthly Trend Data</h3><div class='wrap'><table><thead><tr><th>Month</th><th>Submissions</th><th>Completed</th><th>Completion Rate</th><th>Active Orgs</th><th>Errors</th><th>Error Rate</th></tr></thead><tbody>" & _trows & "</tbody></table></div></div></div></body></html>"
"@

$dax3 = @"
VAR _css = "$css"
VAR _qb = ADDCOLUMNS(SUMMARIZE(UshurHub, UshurHub[YearQuarter]), "_ts", [Total Submissions], "_cs", [Completed Submissions], "_cr", [Completion Rate], "_ao", [Active Organizations], "_te", [Total Errors], "_er", [Error Rate])
VAR _maxQTS = MAXX(_qb, [_ts])
VAR _bQ = CONCATENATEX(_qb, "<div class='br'><div class='bl'>" & UshurHub[YearQuarter] & "</div><div class='bt'><div class='bf' style='width:" & FORMAT(MIN(DIVIDE([_ts], _maxQTS) * 100, 100), "0") & "%;background:#5c61ac'></div></div><div class='bv'>" & FORMAT([_ts], "#,0") & "</div></div>", "", UshurHub[YearQuarter], ASC)
VAR _bQCR = CONCATENATEX(_qb, "<div class='br'><div class='bl'>" & UshurHub[YearQuarter] & "</div><div class='bt'><div class='bf' style='width:" & FORMAT([_cr] * 100, "0") & "%;background:#55c2b6'></div></div><div class='bv'>" & FORMAT([_cr], "0.0%") & "</div></div>", "", UshurHub[YearQuarter], ASC)
VAR _qrows = CONCATENATEX(_qb, "<tr><td>" & UshurHub[YearQuarter] & "</td><td>" & FORMAT([_ts], "#,0") & "</td><td>" & FORMAT([_cs], "#,0") & "</td><td>" & FORMAT([_cr], "0.0%") & "</td><td>" & FORMAT([_ao], "#,0") & "</td><td>" & FORMAT([_te], "#,0") & "</td><td>" & FORMAT([_er], "0.0%") & "</td></tr>", "", UshurHub[YearQuarter], ASC)
RETURN "<!DOCTYPE html><html><head><meta charset='utf-8'/><style>" & _css & "</style></head><body><div class='topbar'><div><div class='ttl'>SecureShare - Ushur Hub</div><div class='sub'>Monthly Engagement - Quarterly Trends</div></div></div><div class='g2'><div class='card'><h3>Quarterly Submissions</h3>" & _bQ & "</div><div class='card'><h3>Quarterly Completion Rate</h3>" & _bQCR & "</div><div class='card' style='grid-column:span 2'><h3>Quarterly Trend Data</h3><div class='wrap'><table><thead><tr><th>Quarter</th><th>Submissions</th><th>Completed</th><th>Completion Rate</th><th>Active Orgs</th><th>Errors</th><th>Error Rate</th></tr></thead><tbody>" & _qrows & "</tbody></table></div></div></div></body></html>"
"@

$dax4 = @"
VAR _css = "$css"
VAR _ucBase = ADDCOLUMNS(SUMMARIZE(UshurHub, UshurHub[UseCase]), "_ts", [Total Submissions])
VAR _maxUC = MAXX(_ucBase, [_ts])
VAR _bUC = CONCATENATEX(TOPN(15, _ucBase, [_ts], DESC), "<div class='br'><div class='bl'>" & UshurHub[UseCase] & "</div><div class='bt'><div class='bf' style='width:" & FORMAT(MIN(DIVIDE([_ts], _maxUC) * 100, 100), "0") & "%;background:#5c61ac'></div></div><div class='bv'>" & FORMAT([_ts], "#,0") & "</div></div>", "", [_ts], DESC)
VAR _stBase = ADDCOLUMNS(SUMMARIZE(UshurHub, UshurHub[Status]), "_ts", [Total Submissions])
VAR _maxST = MAXX(_stBase, [_ts])
VAR _bST = CONCATENATEX(_stBase, "<div class='br'><div class='bl'>" & UshurHub[Status] & "</div><div class='bt'><div class='bf' style='width:" & FORMAT(MIN(DIVIDE([_ts], _maxST) * 100, 100), "0") & "%;background:#55c2b6'></div></div><div class='bv'>" & FORMAT([_ts], "#,0") & "</div></div>", "", [_ts], DESC)
VAR _ttl = [Total Submissions]
VAR _cmp = [Completed Submissions]
VAR _cncl = [Cancelled Submissions]
VAR _err = [Total Errors]
VAR _funnel = "<div class='fstep'><strong>" & FORMAT(_ttl, "#,0") & "</strong><div class='fmeta'>Total Received</div></div><div class='fstep'><strong style='color:#5c61ac'>" & FORMAT(_cmp + _err + _cncl, "#,0") & " (" & FORMAT(DIVIDE(_cmp + _err + _cncl, _ttl), "0.0%") & ")</strong><div class='fmeta'>Processed (completed, errors, cancelled)</div></div><div class='fstep'><strong style='color:#14763d'>" & FORMAT(_cmp, "#,0") & " (" & FORMAT(DIVIDE(_cmp, _ttl), "0.0%") & ")</strong><div class='fmeta'>Successfully Completed</div></div><div class='fstep'><strong style='color:#e03e52'>" & FORMAT(_err, "#,0") & " (" & FORMAT(DIVIDE(_err, _ttl), "0.0%") & ")</strong><div class='fmeta'>Errors</div></div>"
VAR _matBase = ADDCOLUMNS(SUMMARIZE(UshurHub, UshurHub[UseCase]), "_tot", [Total Submissions], "_can", CALCULATE([Cancelled Submissions]), "_cmp2", CALCULATE([Completed Submissions]), "_err2", CALCULATE([Total Errors]), "_rec", CALCULATE([Received Submissions]), "_sub2", CALCULATE([Submitted Submissions]))
VAR _mrows = CONCATENATEX(TOPN(15, _matBase, [_tot], DESC), "<tr><td>" & UshurHub[UseCase] & "</td><td>" & FORMAT([_can], "#,0") & "</td><td>" & FORMAT([_cmp2], "#,0") & "</td><td>" & FORMAT([_err2], "#,0") & "</td><td>" & FORMAT([_rec], "#,0") & "</td><td>" & FORMAT([_sub2], "#,0") & "</td><td>" & FORMAT([_tot], "#,0") & "</td></tr>", "", [_tot], DESC)
VAR _errBase = ADDCOLUMNS(SUMMARIZE(UshurHub, UshurHub[UseCase]), "_te", CALCULATE([Total Errors]))
VAR _erows = CONCATENATEX(TOPN(10, FILTER(_errBase, [_te] > 0), [_te], DESC), "<tr><td>" & UshurHub[UseCase] & "</td><td>" & FORMAT([_te], "#,0") & "</td><td>" & FORMAT(DIVIDE([_te], _err), "0.0%") & " of all errors</td></tr>", "", [_te], DESC)
VAR _noErr = IF(_err = 0, "<tr><td colspan='3' style='text-align:center;color:#67747f;padding:10px'>No errors in selection</td></tr>", _erows)
RETURN "<!DOCTYPE html><html><head><meta charset='utf-8'/><style>" & _css & "</style></head><body><div class='topbar'><div><div class='ttl'>SecureShare - Ushur Hub</div><div class='sub'>Monthly Engagement - Detail</div></div></div><div class='g2'><div class='card'><h3>Submissions by Use Case</h3>" & _bUC & "</div><div class='card'><h3>Status Breakdown</h3>" & _bST & "</div><div class='card'><h3>Completion Funnel</h3>" & _funnel & "</div><div class='card'><h3>Error Concentration</h3><div class='wrap' style='max-height:220px'><table><thead><tr><th>Use Case</th><th>Errors</th><th>Share</th></tr></thead><tbody>" & _noErr & "</tbody></table></div></div><div class='card' style='grid-column:span 2'><h3>Use Case x Status Matrix</h3><div class='wrap'><table><thead><tr><th>Use Case</th><th>Cancelled</th><th>Completed</th><th>Error</th><th>Received</th><th>Submitted</th><th>Total</th></tr></thead><tbody>" & _mrows & "</tbody></table></div></div></div></body></html>"
"@

# ============================================================
# Build TMDL measure blocks
# expression lines: 3 tabs | properties: 2 tabs
# ============================================================
function Build-TmdlMeasure([string]$name, [string]$dax, [string]$guid) {
    $lines = ($dax -replace "`r`n","`n" -replace "`r","`n") -split "`n" |
             Where-Object { $_.Trim() -ne "" } |
             ForEach-Object { "`t`t`t$_" }
    $body = $lines -join "`n"
    "`n`tmeasure '$name' =`n$body`n`t`tformatString: `n`t`tdisplayFolder: HTML Pages`n`t`tlineageTag: $guid"
}

$newMeasures = (Build-TmdlMeasure "HTML Summary"          $dax1 $g1) +
               (Build-TmdlMeasure "HTML Monthly Trends"   $dax2 $g2) +
               (Build-TmdlMeasure "HTML Quarterly Trends" $dax3 $g3) +
               (Build-TmdlMeasure "HTML Detail"           $dax4 $g4)

# ============================================================
# 1. Patch UshurHub.tmdl
# ============================================================
$tmdl = [System.IO.File]::ReadAllText($tmdlPath, [System.Text.Encoding]::UTF8) -replace "`r`n","`n"

if ($tmdl -like "*measure 'HTML Summary'*") {
    Write-Host "[SKIP] HTML measures already present in TMDL" -ForegroundColor Yellow
} else {
    $anchor = "`n`tcolumn "
    $idx    = $tmdl.IndexOf($anchor)
    if ($idx -lt 0) { Write-Error "Could not find column anchor in TMDL"; exit 1 }
    $tmdl   = $tmdl.Insert($idx, $newMeasures)
    [System.IO.File]::WriteAllText($tmdlPath, $tmdl, $noBom)
    Write-Host "[OK] UshurHub.tmdl — 4 HTML measures inserted" -ForegroundColor Green
}

# ============================================================
# 2. Patch report.json — publicCustomVisuals
# ============================================================
$htmlGuid = "htmlContent443BE3AD55E043BF878BED274D3A6855"
$rptRaw   = [System.IO.File]::ReadAllText($rptJson, [System.Text.Encoding]::UTF8)

if ($rptRaw -like "*$htmlGuid*") {
    Write-Host "[SKIP] publicCustomVisuals already in report.json" -ForegroundColor Yellow
} else {
    $rpt = $rptRaw | ConvertFrom-Json
    if (-not $rpt.publicCustomVisuals) {
        $rpt | Add-Member -MemberType NoteProperty -Name 'publicCustomVisuals' -Value @($htmlGuid)
    } else {
        $arr = [System.Collections.ArrayList]@($rpt.publicCustomVisuals)
        $arr.Add($htmlGuid) | Out-Null
        $rpt.publicCustomVisuals = $arr.ToArray()
    }
    [System.IO.File]::WriteAllText($rptJson, ($rpt | ConvertTo-Json -Depth 20), $noBom)
    Write-Host "[OK] report.json — publicCustomVisuals added" -ForegroundColor Green
}

# ============================================================
# 3. Replace non-slicer visuals with HTML Content visuals
# ============================================================
$pages = @(
    [pscustomobject]@{ guid="965f4a0e8701b6becab8"; measure="HTML Summary";          visName="html_p1_summary"   }
    [pscustomobject]@{ guid="a2b3c4d5e6f7a8b9c0d1"; measure="HTML Monthly Trends";   visName="html_p2_monthly"   }
    [pscustomobject]@{ guid="b3c4d5e6f7a8b9c0d1e2"; measure="HTML Quarterly Trends"; visName="html_p3_quarterly" }
    [pscustomobject]@{ guid="c4d5e6f7a8b9c0d1e2f3"; measure="HTML Detail";           visName="html_p4_detail"    }
)

foreach ($p in $pages) {
    $visuRoot = Join-Path $pagesRoot "$($p.guid)\visuals"
    # Remove all non-slicer visual folders
    Get-ChildItem $visuRoot -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notlike "slicer_*" } |
        ForEach-Object { Remove-Item $_.FullName -Recurse -Force; Write-Host "  Removed: $($_.Name)" }

    # Create HTML Content visual
    $visFolder = Join-Path $visuRoot $p.visName
    New-Item -ItemType Directory -Path $visFolder -Force | Out-Null

    $visualJson = [ordered]@{
        '$schema' = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.3.0/schema.json"
        name      = $p.visName
        position  = [ordered]@{ x=10; y=100; z=1000; width=1260; height=612; tabOrder=10000 }
        visual    = [ordered]@{
            visualType = "htmlContent443BE3AD55E043BF878BED274D3A6855"
            query      = [ordered]@{
                queryState = [ordered]@{
                    Values = [ordered]@{
                        projections = @(
                            [ordered]@{
                                field    = [ordered]@{
                                    Measure = [ordered]@{
                                        Expression = [ordered]@{
                                            SourceRef = [ordered]@{ Entity = "UshurHub" }
                                        }
                                        Property = $p.measure
                                    }
                                }
                                queryRef = "UshurHub.$($p.measure)"
                            }
                        )
                    }
                }
            }
        }
    }
    [System.IO.File]::WriteAllText("$visFolder\visual.json", ($visualJson | ConvertTo-Json -Depth 20), $noBom)
    Write-Host "[OK] $($p.visName)  ->  '$($p.measure)'" -ForegroundColor Green
}

Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host " Done! Open Monthly Engagement.pbip in PBI Desktop." -ForegroundColor Cyan
Write-Host " Each page has one HTML Content visual bound to" -ForegroundColor Cyan
Write-Host " its measure. Slicers remain for filtering." -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
