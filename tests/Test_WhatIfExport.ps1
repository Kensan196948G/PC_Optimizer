# ==============================================================
# Test_WhatIfExport.ps1
# Runs WhatIf twice and verifies exported deleted-path list is stable
# ==============================================================

param(
    [string]$RepoRoot = (Split-Path $PSScriptRoot -Parent)
)

$scriptPath = Join-Path $RepoRoot "PC_Optimizer.ps1"
$logsDir = Join-Path $RepoRoot "logs"
$outDir = Join-Path $env:TEMP ("PCOptDeletedPaths_" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

$pass = 0
$fail = 0
function Assert-True {
    param([string]$Name, [bool]$Condition, [string]$Detail = "")
    if ($Condition) {
        $script:pass++
        Write-Host "  [PASS] $Name" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "  [FAIL] $Name" -ForegroundColor Red
        if ($Detail) { Write-Host "         => $Detail" -ForegroundColor DarkYellow }
    }
}

Assert-True "WE-01: script exists" (Test-Path $scriptPath) $scriptPath
if ($fail -gt 0) { exit 1 }

function Invoke-WhatIfRun {
    param([ValidateSet('json','csv')][string]$Format)
    $before = @()
    $pattern = if ($Format -eq 'json') { "DeletedPaths_*.json" } else { "DeletedPaths_*.csv" }
    $before = @(Get-ChildItem -Path $outDir -Filter $pattern -File -ErrorAction SilentlyContinue)

    & powershell -NoProfile -ExecutionPolicy Bypass -File $scriptPath `
        -NonInteractive -WhatIf -NoRebootPrompt `
        -Tasks "1-3,7" -ExportDeletedPaths $Format -ExportDeletedPathsPath $outDir | Out-Null
    $exitCode = $LASTEXITCODE

    $after = @(Get-ChildItem -Path $outDir -Filter $pattern -File | Sort-Object LastWriteTimeUtc)
    $newFile = $after | Where-Object { $before.FullName -notcontains $_.FullName } | Select-Object -Last 1

    return [PSCustomObject]@{
        ExitCode = $exitCode
        File = $newFile
    }
}

$r1 = Invoke-WhatIfRun -Format json
$r2 = Invoke-WhatIfRun -Format json
$r3 = Invoke-WhatIfRun -Format csv
& powershell -NoProfile -ExecutionPolicy Bypass -File $scriptPath `
    -NonInteractive -WhatIf -NoRebootPrompt `
    -Tasks "1-3,7" -ExportDeletedPaths json -ExportDeletedPathsPath "C:\temp\bad|path" | Out-Null
$invalidPathExit = $LASTEXITCODE

Assert-True "WE-02: first run exit code is 0" ($r1.ExitCode -eq 0) "exit=$($r1.ExitCode)"
Assert-True "WE-03: second run exit code is 0" ($r2.ExitCode -eq 0) "exit=$($r2.ExitCode)"
Assert-True "WE-04: first run exported file exists" ($null -ne $r1.File -and (Test-Path $r1.File.FullName)) "file=$($r1.File.FullName)"
Assert-True "WE-05: second run exported file exists" ($null -ne $r2.File -and (Test-Path $r2.File.FullName)) "file=$($r2.File.FullName)"
Assert-True "WE-05b: csv run exit code is 0" ($r3.ExitCode -eq 0) "exit=$($r3.ExitCode)"
Assert-True "WE-05c: csv exported file exists" ($null -ne $r3.File -and (Test-Path $r3.File.FullName)) "file=$($r3.File.FullName)"
Assert-True "WE-05d: invalid ExportDeletedPathsPath exits 3" ($invalidPathExit -eq 3) "exit=$invalidPathExit"

if ($fail -eq 0) {
    $j1 = Get-Content -Path $r1.File.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
    $j2 = Get-Content -Path $r2.File.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
    $t1 = @($j1.tasks | Sort-Object)
    $t2 = @($j2.tasks | Sort-Object)

    Assert-True "WE-06: exported task list count is stable" ($t1.Count -eq $t2.Count) "count1=$($t1.Count), count2=$($t2.Count)"
    Assert-True "WE-07: exported task list content is stable" (($t1 -join "|") -eq ($t2 -join "|")) "list differs"
    Assert-True "WE-08: expected Chrome cache path exists" (($t1 -join "|") -match 'Google\\Chrome\\User Data\\Default\\Cache') "expected path missing"

    $csvRows = @(Import-Csv -Path $r3.File.FullName)
    Assert-True "WE-09: csv has rows" ($csvRows.Count -gt 0) "rows=$($csvRows.Count)"
    Assert-True "WE-10: csv includes Chrome path" ((@($csvRows.path) -join "|") -match 'Google\\Chrome\\User Data\\Default\\Cache') "expected path missing"
}

Write-Host ""
Write-Host "PASS: $pass / $($pass + $fail)" -ForegroundColor Green
Write-Host "FAIL: $fail / $($pass + $fail)" -ForegroundColor Yellow

if ($fail -gt 0) { exit 1 }
exit 0
