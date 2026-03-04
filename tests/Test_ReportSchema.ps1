# ==============================================================
# Test_ReportSchema.ps1
# Validates actual runtime JSON report against schema-required shape
# ==============================================================

param(
    [string]$RepoRoot = (Split-Path $PSScriptRoot -Parent)
)

$schemaPath = Join-Path $RepoRoot "docs\schemas\pc-optimizer-report-v1.schema.json"
$scriptPath = Join-Path $RepoRoot "PC_Optimizer.ps1"
$logsDir = Join-Path $RepoRoot "logs"

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
if (Test-Path $logsDir) {
    $before = @(Get-ChildItem -Path $logsDir -Filter "PC_Optimizer_Report_*.json" -File)
}

& powershell -NoProfile -ExecutionPolicy Bypass -File $scriptPath `
    -NonInteractive -WhatIf -NoRebootPrompt -Tasks "1-2,7" -FailureMode continue | Out-Null
$runExit = $LASTEXITCODE
Assert-True "RS-03: runtime report generation exits 0 in whatif mode" ($runExit -eq 0) "exit=$runExit"

$after = @(Get-ChildItem -Path $logsDir -Filter "PC_Optimizer_Report_*.json" -File | Sort-Object LastWriteTimeUtc)
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
Assert-True "RS-08: exitCode enum" ($schema.properties.exitCode.enum -contains [int]$actual.exitCode) "exitCode=$($actual.exitCode)"
Assert-True "RS-09: failureMode enum" ($schema.properties.failureMode.enum -contains $actual.failureMode) "failureMode=$($actual.failureMode)"
Assert-True "RS-10: tasks count >= 1" (@($actual.tasks).Count -ge 1) "tasks=$(@($actual.tasks).Count)"

$task0 = @($actual.tasks)[0]
foreach ($requiredTask in $schema.properties.tasks.items.required) {
    Assert-True "RS-11: task required '$requiredTask' exists" `
        ($null -ne $task0.PSObject.Properties[$requiredTask]) `
        "missing task.$requiredTask"
}
Assert-True "RS-12: task status enum" `
    ($schema.properties.tasks.items.properties.status.enum -contains $task0.status) `
    "task.status=$($task0.status)"

Write-Host ""
Write-Host "PASS: $pass / $($pass + $fail)" -ForegroundColor Green
Write-Host "FAIL: $fail / $($pass + $fail)" -ForegroundColor Yellow

if ($fail -gt 0) { exit 1 }
exit 0
