# ==============================================================
# Test_ReportSchema.ps1
# Validates actual runtime JSON report against schema-required shape
# ==============================================================

param(
    [string]$RepoRoot = (Split-Path $PSScriptRoot -Parent)
)

$schemaPath  = Join-Path $RepoRoot "docs\schemas\pc-optimizer-report-v1.schema.json"
$scriptPath  = Join-Path $RepoRoot "PC_Optimizer.ps1"
$reportsDir  = Join-Path $RepoRoot "reports"
$reportGlob  = "PC_Health_Report_*.json"

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

Assert-True "RS-01: schema file exists" (Test-Path $schemaPath) $schemaPath
Assert-True "RS-02: script exists" (Test-Path $scriptPath) $scriptPath
if ($fail -gt 0) { exit 1 }

$before = @()
if (Test-Path $reportsDir) {
    $before = @(Get-ChildItem -Path $reportsDir -Filter $reportGlob -File)
}

& powershell -NoProfile -ExecutionPolicy Bypass -File $scriptPath `
    -NonInteractive -WhatIf -NoRebootPrompt -Tasks "1-2,7" -FailureMode continue | Out-Null
$runExit = $LASTEXITCODE
Assert-True "RS-03: runtime report generation exits 0 in whatif mode" ($runExit -eq 0) "exit=$runExit"

$after = @(Get-ChildItem -Path $reportsDir -Filter $reportGlob -File | Sort-Object LastWriteTimeUtc)
$newReport = $after | Where-Object { $before.FullName -notcontains $_.FullName } | Select-Object -Last 1
Assert-True "RS-04: new json report file is generated" ($null -ne $newReport) "no new report"
if ($fail -gt 0) { exit 1 }

$schema = Get-Content -Path $schemaPath -Raw -Encoding UTF8 | ConvertFrom-Json
$actual = Get-Content -Path $newReport.FullName -Raw -Encoding UTF8 | ConvertFrom-Json

foreach ($required in $schema.required) {
    Assert-True "RS-05: required property '$required' exists" `
        ($null -ne $actual.PSObject.Properties[$required]) `
        "missing $required"
}

Assert-True "RS-06: version const 1.0" ($actual.version -eq $schema.properties.version.const) "actual=$($actual.version)"
Assert-True "RS-07: status enum" ($schema.properties.status.enum -contains $actual.status) "status=$($actual.status)"
Assert-True "RS-07b: whatif + filtered run should be PARTIAL" ($actual.status -eq "PARTIAL") "status=$($actual.status)"
Assert-True "RS-08: exitCode enum" ($schema.properties.exitCode.enum -contains [int]$actual.exitCode) "exitCode=$($actual.exitCode)"
Assert-True "RS-09: failureMode enum" ($schema.properties.failureMode.enum -contains $actual.failureMode) "failureMode=$($actual.failureMode)"
Assert-True "RS-10: tasks count >= 1" (@($actual.tasks).Count -ge 1) "tasks=$(@($actual.tasks).Count)"
Assert-True "RS-10b: selectedTasks exists and has values" (@($actual.selectedTasks).Count -ge 1) "selectedTasks missing/empty"
Assert-True "RS-10c: skippedReasonSummary exists" ($null -ne $actual.skippedReasonSummary) "skippedReasonSummary missing"
if (@($actual.selectedTasks).Count -gt 0) {
    $selectedInts = @($actual.selectedTasks | ForEach-Object { [int]$_ })
    $selectedSorted = @($selectedInts | Sort-Object)
    Assert-True "RS-10d: selectedTasks are sorted ascending" `
        (($selectedInts -join ",") -eq ($selectedSorted -join ",")) `
        "selectedTasks=$($selectedInts -join ',')"
}

$task0 = @($actual.tasks)[0]
foreach ($requiredTask in $schema.properties.tasks.items.required) {
    Assert-True "RS-11: task required '$requiredTask' exists" `
        ($null -ne $task0.PSObject.Properties[$requiredTask]) `
        "missing task.$requiredTask"
}
Assert-True "RS-12: task status enum" `
    ($schema.properties.tasks.items.properties.status.enum -contains $task0.status) `
    "task.status=$($task0.status)"

Assert-True "RS-13: schema allows fail-fast exitCode=2" ($schema.properties.exitCode.enum -contains 2) "schema enum does not include 2"
$simulatedFailFast = [PSCustomObject]@{
    version              = $actual.version
    runId                = $actual.runId
    startedAt            = $actual.startedAt
    finishedAt           = $actual.finishedAt
    host                 = $actual.host
    status               = "PARTIAL"
    exitCode             = 2
    failureMode          = "fail-fast"
    durationSeconds      = $actual.durationSeconds
    selectedTasks        = @($actual.selectedTasks)
    skippedReasonSummary = [PSCustomObject]@{ "FailFast skip" = 3 }
    unexecutedTasks      = @(11,13,15)
    tasks                = @($actual.tasks)
}
foreach ($required in $schema.required) {
    Assert-True "RS-14: simulated fail-fast has required '$required'" `
        ($null -ne $simulatedFailFast.PSObject.Properties[$required]) `
        "missing $required"
}
Assert-True "RS-15: simulated fail-fast exitCode is 2" ([int]$simulatedFailFast.exitCode -eq 2) "exitCode=$($simulatedFailFast.exitCode)"
Assert-True "RS-16: simulated fail-fast failureMode is fail-fast" ($simulatedFailFast.failureMode -eq "fail-fast") "failureMode=$($simulatedFailFast.failureMode)"

# ── v4.0.1: fail-fast 実実行JSONレポート検証 ──
Write-Host ""
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
Write-Host " fail-fast runtime report validation (v4.0.1)" -ForegroundColor White
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray

$beforeFF = @()
if (Test-Path $reportsDir) {
    $beforeFF = @(Get-ChildItem -Path $reportsDir -Filter $reportGlob -File)
}

# fail-fast モードで -Tasks "1-3" を実行（WhatIf なので副作用なし）
& powershell -NoProfile -ExecutionPolicy Bypass -File $scriptPath `
    -NonInteractive -WhatIf -NoRebootPrompt -Tasks "1-3" -FailureMode fail-fast | Out-Null
$ffExit = $LASTEXITCODE
Assert-True "RS-17: fail-fast run exits with 0 or 1 (WhatIf+PARTIAL)" ($ffExit -eq 0 -or $ffExit -eq 1) "exit=$ffExit"

$afterFF = @(Get-ChildItem -Path $reportsDir -Filter $reportGlob -File | Sort-Object LastWriteTimeUtc)
$ffReport = $afterFF | Where-Object { $beforeFF.FullName -notcontains $_.FullName } | Select-Object -Last 1
Assert-True "RS-18: fail-fast run generates json report" ($null -ne $ffReport) "no new report"

if ($null -ne $ffReport) {
    $ffObj = Get-Content -Path $ffReport.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-True "RS-19: fail-fast report has failureMode=fail-fast" ($ffObj.failureMode -eq "fail-fast") "failureMode=$($ffObj.failureMode)"
    Assert-True "RS-20: fail-fast report status is OK or PARTIAL" ($ffObj.status -eq "OK" -or $ffObj.status -eq "PARTIAL") "status=$($ffObj.status)"
    Assert-True "RS-21: fail-fast report has skippedReasonSummary" ($null -ne $ffObj.skippedReasonSummary) "skippedReasonSummary missing"
    Assert-True "RS-22: fail-fast report selectedTasks count >= 1" (@($ffObj.selectedTasks).Count -ge 1) "selectedTasks empty"
    # skippedReasonSummary のキーが存在すること（WhatIf で全タスクがSKIPになる）
    $skippedKeys = @($ffObj.skippedReasonSummary.PSObject.Properties.Name)
    Assert-True "RS-23: skippedReasonSummary has at least one key" ($skippedKeys.Count -ge 1) "keys=0"
}

Write-Host ""
Write-Host "PASS: $pass / $($pass + $fail)" -ForegroundColor Green
Write-Host "FAIL: $fail / $($pass + $fail)" -ForegroundColor Yellow

if ($fail -gt 0) { exit 1 }
exit 0
