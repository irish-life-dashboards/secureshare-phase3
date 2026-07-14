$src = "c:\Users\m347\Documents\Power BI Desktop\SecureShare Report\secureshare-phase3"
$dst = "c:\Users\m347\OneDrive - GWLE\James, Stephen's files - Secure Share Dashboard reports\SecureShare Phase 3 reports"

if (-not (Test-Path $dst)) {
    New-Item -ItemType Directory -Path $dst -Force | Out-Null
}

$items = @(
    "Landing page.html",
    "assets",
    "Operational Reporting",
    "Finance Performance",
    "EFT Confirmation",
    "VRDD Appropriations",
    "Leavers Report",
    "Mothly Engagement"
)

foreach ($item in $items) {
    $s = Join-Path $src $item
    $d = Join-Path $dst $item

    if (Test-Path $s) {
        Copy-Item -Path $s -Destination $d -Recurse -Force
        Write-Output "Copied: $item"
    } else {
        Write-Output "Missing source: $item"
    }
}
