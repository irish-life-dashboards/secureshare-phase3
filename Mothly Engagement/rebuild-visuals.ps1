$noBom = New-Object System.Text.UTF8Encoding $false
$base = "C:\Users\m347\Documents\Power BI Desktop\SecureShare Report\secureshare-phase3\Mothly Engagement\PBIP\Monthly Engagement.Report\definition\pages"
$sv = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.3.0/schema.json"
$sp = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.0.0/schema.json"
$sm = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.0.0/schema.json"
$e  = "UshurHub"

function Write-Json($path, $json) {
    $dir = Split-Path $path -Parent
    if (!(Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    [System.IO.File]::WriteAllText($path, $json, $noBom)
}
function wv($pg, $name, $json) { Write-Json "$base\$pg\visuals\$name\visual.json" $json }

# Build position JSON
function pos($x,$y,$z,$h,$w,$tab) { '"position":{"x":' + $x + ',"y":' + $y + ',"z":' + $z + ',"height":' + $h + ',"width":' + $w + ',"tabOrder":' + $tab + '}' }

# Build title visualContainerObjects
function title($t) { '"visualContainerObjects":{"title":[{"properties":{"show":{"expr":{"Literal":{"Value":"true"}}},"text":{"expr":{"Literal":{"Value":"' + "'" + $t + "'" + '"}}}}}]}' }

# Build a column reference projection
function colproj($prop) { '{"field":{"Column":{"Expression":{"SourceRef":{"Entity":"' + $e + '"}},"Property":"' + $prop + '"}},"queryRef":"' + $e + '.' + $prop + '"}' }

# Build a measure reference projection
function msproj($measure) { '{"field":{"Measure":{"Expression":{"SourceRef":{"Entity":"' + $e + '"}},"Property":"' + $measure + '"}},"queryRef":"' + $e + '.' + $measure + '"}' }

# Slicer visual
function mksl($name,$x,$y,$z,$h,$w,$tab,$ftype,$prop,$sg,$titleText) {
    '{"$schema":"' + $sv + '","name":"' + $name + '",' + (pos $x $y $z $h $w $tab) + ',"visual":{"visualType":"slicer","query":{"queryState":{"Field":{"projections":[{"field":{"' + $ftype + '":{"Expression":{"SourceRef":{"Entity":"' + $e + '"}},"Property":"' + $prop + '"}},"queryRef":"' + $e + '.' + $prop + '"}]}}},"syncGroup":{"groupName":"' + $sg + '","fieldChanges":false,"filterChanges":true},' + (title $titleText) + '}}'
}

# Card visual
function mkcard($name,$x,$y,$z,$h,$w,$tab,$measure,$titleText) {
    '{"$schema":"' + $sv + '","name":"' + $name + '",' + (pos $x $y $z $h $w $tab) + ',"visual":{"visualType":"card","query":{"queryState":{"Values":{"projections":[' + (msproj $measure) + ']}}},' + (title $titleText) + '}}'
}

# Single-measure chart (column / line / bar)
function mkchart($name,$vtype,$x,$y,$z,$h,$w,$tab,$catProp,$catFtype,$measure,$titleText) {
    '{"$schema":"' + $sv + '","name":"' + $name + '",' + (pos $x $y $z $h $w $tab) + ',"visual":{"visualType":"' + $vtype + '","query":{"queryState":{"Category":{"projections":[{"field":{"' + $catFtype + '":{"Expression":{"SourceRef":{"Entity":"' + $e + '"}},"Property":"' + $catProp + '"}},"queryRef":"' + $e + '.' + $catProp + '"}]},"Y":{"projections":[' + (msproj $measure) + ']}}},' + (title $titleText) + '}}'
}

# Table visual
function mktable($name,$x,$y,$z,$h,$w,$tab,$fieldsJson,$titleText) {
    '{"$schema":"' + $sv + '","name":"' + $name + '",' + (pos $x $y $z $h $w $tab) + ',"visual":{"visualType":"tableEx","query":{"queryState":{"Values":{"projections":[' + $fieldsJson + ']}}},' + (title $titleText) + '}}'
}

# pages.json
Write-Json "$base\pages.json" ('{"$schema":"' + $sm + '","pageOrder":["965f4a0e8701b6becab8","a2b3c4d5e6f7a8b9c0d1","b3c4d5e6f7a8b9c0d1e2","c4d5e6f7a8b9c0d1e2f3"],"activePageName":"965f4a0e8701b6becab8"}')

foreach ($pg in @(
    @{guid="965f4a0e8701b6becab8"; name="Summary"},
    @{guid="a2b3c4d5e6f7a8b9c0d1"; name="Monthly Trends"},
    @{guid="b3c4d5e6f7a8b9c0d1e2"; name="Quarterly Trends"},
    @{guid="c4d5e6f7a8b9c0d1e2f3"; name="Detail"}
)) {
    Write-Json "$base\$($pg.guid)\page.json" ('{"$schema":"' + $sp + '","name":"' + $pg.guid + '","displayName":"' + $pg.name + '","displayOption":"FitToPage","height":720,"width":1280}')
}

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 1 — SUMMARY
# Layout: Left slicers (x=10,w=188) | 4 KPI cards top | 2 tables bottom
# Canvas: 1280x720
# ═══════════════════════════════════════════════════════════════════════════════
$p1 = "965f4a0e8701b6becab8"
# Clear existing visuals on p1 (diagnostic test left it empty)
Get-ChildItem "$base\$p1\visuals" -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force

# Slicers — left column x=10, w=188
wv $p1 "slicer_p1_month"   (mksl "slicer_p1_month"    10  10 1000  80 188  0 "Column" "MonthYear"    "sg_month"   "Month")
wv $p1 "slicer_p1_quarter" (mksl "slicer_p1_quarter"  10 100 2000  80 188  1 "Column" "YearQuarter"  "sg_quarter" "Quarter")
wv $p1 "slicer_p1_status"  (mksl "slicer_p1_status"   10 190 3000  80 188  2 "Column" "Status"       "sg_status"  "Status")
wv $p1 "slicer_p1_usecase" (mksl "slicer_p1_usecase"  10 280 4000 110 188  3 "Column" "UseCase"      "sg_usecase" "Use Case")
wv $p1 "slicer_p1_org"     (mksl "slicer_p1_org"      10 400 5000 110 188  4 "Column" "Organization" "sg_org"     "Organization")

# KPI Cards — y=10, h=100, evenly spaced across content area (x=208 to x=1258)
wv $p1 "card_p1_totalsubs"  (mkcard "card_p1_totalsubs"   208  10 6000 100 255  5 "Total Submissions"    "Total Submissions")
wv $p1 "card_p1_comprate"   (mkcard "card_p1_comprate"    473  10 7000 100 255  6 "Completion Rate"      "Completion Rate")
wv $p1 "card_p1_activeorgs" (mkcard "card_p1_activeorgs"  738  10 8000 100 255  7 "Active Organizations" "Active Organizations")
wv $p1 "card_p1_errorrate"  (mkcard "card_p1_errorrate"  1003  10 9000 100 255  8 "Error Rate"           "Error Rate")

# Tables — y=120, h=590, side by side
$wl = (colproj "Organization") + "," + (msproj "Total Submissions") + "," + (msproj "Completed Submissions") + "," + (msproj "Completion Rate") + "," + (msproj "Total Errors") + "," + (msproj "Error Rate")
wv $p1 "table_p1_watchlist"  (mktable "table_p1_watchlist"  208 120 10000 590 520  9 $wl "Watch List - Low Completion Rate")

$t20 = (colproj "Organization") + "," + (msproj "Total Submissions") + "," + (msproj "Completion Rate") + "," + (msproj "Error Rate")
wv $p1 "table_p1_top20orgs"  (mktable "table_p1_top20orgs"  738 120 11000 590 530 10 $t20 "Top Organizations by Volume")

Write-Host "Page 1 done (11 visuals)"

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 2 — MONTHLY TRENDS
# Layout: Left slicers | Top: full-width monthly bar | Bottom: 3 trend charts
# FIXED: Removed overlapping table (was also at y=10)
# ═══════════════════════════════════════════════════════════════════════════════
$p2 = "a2b3c4d5e6f7a8b9c0d1"
Get-ChildItem "$base\$p2\visuals" -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force

wv $p2 "slicer_p2_month"   (mksl "slicer_p2_month"    10  10 1000  80 188  0 "Column" "MonthYear"    "sg_month"   "Month")
wv $p2 "slicer_p2_quarter" (mksl "slicer_p2_quarter"  10 100 2000  80 188  1 "Column" "YearQuarter"  "sg_quarter" "Quarter")
wv $p2 "slicer_p2_status"  (mksl "slicer_p2_status"   10 190 3000  80 188  2 "Column" "Status"       "sg_status"  "Status")
wv $p2 "slicer_p2_usecase" (mksl "slicer_p2_usecase"  10 280 4000 110 188  3 "Column" "UseCase"      "sg_usecase" "Use Case")
wv $p2 "slicer_p2_org"     (mksl "slicer_p2_org"      10 400 5000 110 188  4 "Column" "Organization" "sg_org"     "Organization")

# Top: full-width monthly submissions column chart
wv $p2 "col_p2_monthsubs"  (mkchart "col_p2_monthsubs"  "columnChart" 208  10 6000 340 1060  5 "MonthYear"    "Column" "Total Submissions"    "Monthly Submissions")

# Bottom row: 3 KPI trend charts  (y=360, h=350 — fits within 720)
wv $p2 "line_p2_comprate"  (mkchart "line_p2_comprate"  "lineChart"   208 360 7000 350  340  6 "MonthYear"    "Column" "Completion Rate"      "Completion Rate Trend")
wv $p2 "col_p2_activeorgs" (mkchart "col_p2_activeorgs" "columnChart" 558 360 8000 350  340  7 "MonthYear"    "Column" "Active Organizations" "Active Organizations by Month")
wv $p2 "line_p2_errorrate" (mkchart "line_p2_errorrate" "lineChart"   908 360 9000 350  360  8 "MonthYear"    "Column" "Error Rate"           "Error Rate Trend")

Write-Host "Page 2 done (9 visuals)"

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 3 — QUARTERLY TRENDS
# Layout: Left slicers | Top: quarterly bar chart | Bottom: quarterly table
# ═══════════════════════════════════════════════════════════════════════════════
$p3 = "b3c4d5e6f7a8b9c0d1e2"
Get-ChildItem "$base\$p3\visuals" -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force

wv $p3 "slicer_p3_month"   (mksl "slicer_p3_month"    10  10 1000  80 188  0 "Column" "MonthYear"    "sg_month"   "Month")
wv $p3 "slicer_p3_quarter" (mksl "slicer_p3_quarter"  10 100 2000  80 188  1 "Column" "YearQuarter"  "sg_quarter" "Quarter")
wv $p3 "slicer_p3_status"  (mksl "slicer_p3_status"   10 190 3000  80 188  2 "Column" "Status"       "sg_status"  "Status")
wv $p3 "slicer_p3_usecase" (mksl "slicer_p3_usecase"  10 280 4000 110 188  3 "Column" "UseCase"      "sg_usecase" "Use Case")
wv $p3 "slicer_p3_org"     (mksl "slicer_p3_org"      10 400 5000 110 188  4 "Column" "Organization" "sg_org"     "Organization")

wv $p3 "col_p3_quartsubs"  (mkchart "col_p3_quartsubs" "columnChart"  208  10 6000 340 1060  5 "YearQuarter" "Column" "Total Submissions" "Quarterly Submissions")

$qt = (colproj "YearQuarter") + "," + (msproj "Total Submissions") + "," + (msproj "Completed Submissions") + "," + (msproj "Completion Rate") + "," + (msproj "Active Organizations") + "," + (msproj "Error Rate")
wv $p3 "table_p3_quarterly" (mktable "table_p3_quarterly" 208 360 7000 350 1060  6 $qt "Quarterly Summary")

Write-Host "Page 3 done (7 visuals)"

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 4 — DETAIL
# Layout: Left slicers | Top: 2 bar charts side by side | Bottom: matrix
# ═══════════════════════════════════════════════════════════════════════════════
$p4 = "c4d5e6f7a8b9c0d1e2f3"
Get-ChildItem "$base\$p4\visuals" -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force

wv $p4 "slicer_p4_month"   (mksl "slicer_p4_month"    10  10 1000  80 188  0 "Column" "MonthYear"    "sg_month"   "Month")
wv $p4 "slicer_p4_quarter" (mksl "slicer_p4_quarter"  10 100 2000  80 188  1 "Column" "YearQuarter"  "sg_quarter" "Quarter")
wv $p4 "slicer_p4_status"  (mksl "slicer_p4_status"   10 190 3000  80 188  2 "Column" "Status"       "sg_status"  "Status")
wv $p4 "slicer_p4_usecase" (mksl "slicer_p4_usecase"  10 280 4000 110 188  3 "Column" "UseCase"      "sg_usecase" "Use Case")
wv $p4 "slicer_p4_org"     (mksl "slicer_p4_org"      10 400 5000 110 188  4 "Column" "Organization" "sg_org"     "Organization")

# Two horizontal bar charts side by side (Status and UseCase breakdowns)
wv $p4 "bar_p4_status"  (mkchart "bar_p4_status"  "barChart" 208  10 6000 340 510  5 "Status"  "Column" "Total Submissions" "Submissions by Status")
wv $p4 "bar_p4_usecase" (mkchart "bar_p4_usecase" "barChart" 728  10 7000 340 540  6 "UseCase" "Column" "Total Submissions" "Submissions by Use Case")

# Matrix: Use Case (rows) x Status (columns) x Total Submissions
wv $p4 "matrix_p4_ucstatus" ('{"$schema":"' + $sv + '","name":"matrix_p4_ucstatus",' + (pos 208 360 8000 350 1060 7) + ',"visual":{"visualType":"pivotTable","query":{"queryState":{"Rows":{"projections":[' + (colproj "UseCase") + ']},"Columns":{"projections":[' + (colproj "Status") + ']},"Values":{"projections":[' + (msproj "Total Submissions") + ']}}},' + (title "Use Case by Status") + '}}')

Write-Host "Page 4 done (8 visuals)"

# ═══════════════════════════════════════════════════════════════════════════════
# VERIFY
# ═══════════════════════════════════════════════════════════════════════════════
$c1 = (Get-ChildItem "$base\965f4a0e8701b6becab8\visuals").Count
$c2 = (Get-ChildItem "$base\a2b3c4d5e6f7a8b9c0d1\visuals").Count
$c3 = (Get-ChildItem "$base\b3c4d5e6f7a8b9c0d1e2\visuals").Count
$c4 = (Get-ChildItem "$base\c4d5e6f7a8b9c0d1e2f3\visuals").Count
$b0 = [System.IO.File]::ReadAllBytes("$base\pages.json")[0]
$bv = [System.IO.File]::ReadAllBytes("$base\965f4a0e8701b6becab8\visuals\card_p1_totalsubs\visual.json")[0]
Write-Host ("Counts: P1=$c1 P2=$c2 P3=$c3 P4=$c4  |  byte0: pages.json=$b0 card=$bv  (123=OK)")
Write-Host "ALL DONE"
