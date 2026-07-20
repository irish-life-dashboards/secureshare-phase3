$BASE = 'c:\Users\m347\Documents\Power BI Desktop\SecureShare Report\secureshare-phase3'
$DIST = "$BASE\dist"

$REPORTS = @(
    @{ folder = 'Operational Reporting'; file = 'Operational Reporting.html' },
    @{ folder = 'SecureShare Deployment Stats'; file = 'SecureShare Deployment Stats.html' },
    @{ folder = 'Leavers Report'; file = 'Leavers Report.html' },
    @{ folder = 'Finance Performance'; file = 'Finance Performance.html' },
    @{ folder = 'EFT Confirmation'; file = 'EFT Confirmation.html' },
    @{ folder = 'VRDD Appropriations'; file = 'VRDD Appropriations.html' },
    @{ folder = 'Mothly Engagement'; file = 'Monthly Engagement.html' }
)

if (-not (Test-Path $DIST)) { New-Item -ItemType Directory $DIST | Out-Null }

foreach ($r in $REPORTS) {
    $htmlPath = "$BASE\$($r.folder)\$($r.file)"
    if (-not (Test-Path $htmlPath)) { Write-Host "SKIP: $htmlPath"; continue }

    Write-Host "`nBuilding: $($r.folder)\$($r.file)"

    $content = [System.IO.File]::ReadAllText($htmlPath)

    # Find and replace data script references with inline content
    $pattern = '<script\s+src=[''"]([^''"]+)[''"](?:\s*/>|>\s*</script>)'
    $regex = [regex]::new($pattern, 'IgnoreCase')
    $matches = $regex.Matches($content)

    foreach ($m in $matches) {
        $src = $m.Groups[1].Value
        $isLocalData = $src.StartsWith('./data/') -or $src.StartsWith('data/')
        $isSharedData = $src.StartsWith('../Operational Reporting/data/')
        if (-not ($isLocalData -or $isSharedData)) { continue }

        if ($isSharedData) {
            $dataFile = Join-Path "$BASE\\$($r.folder)" ($src.Replace('/', '\\'))
            $dataFile = [System.IO.Path]::GetFullPath($dataFile)
        } else {
            $dataFile = "$BASE\$($r.folder)\$($src.Replace('/', '\'))"
        }

        if (Test-Path $dataFile) {
            $size = (Get-Item $dataFile).Length
            Write-Host "  + Inlining $src ($([math]::Round($size/1KB)) KB)"
            $dataContent = [System.IO.File]::ReadAllText($dataFile)
            $inlined = "<script>`n$dataContent`n</script>"
            $content = $content.Replace($m.Value, $inlined)
        } else {
            Write-Host "  WARNING: $dataFile not found"
        }
    }

    $distFolder = "$DIST\$($r.folder)"
    if (-not (Test-Path $distFolder)) { New-Item -ItemType Directory $distFolder | Out-Null }

    $outPath = "$distFolder\$($r.file)"
    [System.IO.File]::WriteAllText($outPath, $content)
    $outSize = (Get-Item $outPath).Length
    Write-Host "  => Output: $outPath ($([math]::Round($outSize/1KB)) KB)"

    # Copy assets
    $assetsSrc = "$BASE\$($r.folder)\assets"
    if (Test-Path $assetsSrc) {
        Copy-Item $assetsSrc -Destination "$distFolder\assets" -Recurse -Force
        Write-Host "  => assets copied"
    }
}

# Copy landing page + root assets
$lp = "$BASE\Landing page.html"
if (Test-Path $lp) {
    Copy-Item $lp "$DIST\Landing page.html" -Force
    Write-Host "`nLanding page copied"
}
$ra = "$BASE\assets"
if (Test-Path $ra) {
    Copy-Item $ra "$DIST\assets" -Recurse -Force
    Write-Host "Root assets copied"
}

Write-Host "`nDone. Standalone reports in: $DIST"
Write-Host "Copy the dist\ folder contents to OneDrive."
