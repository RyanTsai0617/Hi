# Create junctions from .claude/plugins/general-reports/skills/ → OneDrive skill source
# Uses char codes to avoid Chinese path encoding issues in PowerShell 5.1

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$desktop = [char]0x684c + [char]0x9762  # 桌面
$DST = "$env:USERPROFILE\.claude\plugins\general-reports\skills"
$SRC = "$env:OneDrive\$desktop\Coding\.claude\skills"

Write-Host "SRC: $SRC"
Write-Host "DST: $DST"
Write-Host ""

$skills = @("weekly-report", "daily-market-report", "verification-rules")

Write-Host "=== Create junctions ==="
foreach ($s in $skills) {
    $link = Join-Path $DST $s
    $target = Join-Path $SRC $s
    if (Test-Path $link) {
        Write-Host "  $s : exists, skip"
        continue
    }
    if (-not (Test-Path $target)) {
        Write-Host "  $s : target missing -> $target"
        continue
    }
    New-Item -ItemType Junction -Path $link -Target $target | Out-Null
    Write-Host "  $s : junction created"
}

Write-Host ""
Write-Host "=== Verify ==="
foreach ($s in $skills) {
    $f = Join-Path $DST "$s\SKILL.md"
    if (Test-Path $f) {
        Write-Host "  $s/SKILL.md : $((Get-Item $f).Length) bytes OK"
    } else {
        Write-Host "  $s/SKILL.md : FAIL"
    }
}
