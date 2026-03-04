# ==============================================================
# Test_ReportSchema.ps1
# Validates JSON report against schema-required structure
# ==============================================================

param(
    [string]$RepoRoot = (Split-Path $PSScriptRoot -Parent)
)

$schemaPath = Join-Path $RepoRoot "docs\schemas\pc-optimizer-report-v1.schema.json"
$modulePath = Join-Path $RepoRoot "modules\Report.psm1"
$tmpDir = Join-Path $env:TEMP ("PCOptSchema_" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null

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
Assert-True "RS-02: report module exists" (Test-Path $modulePath) $modulePath
if ($fail -gt 0) { exit 1 }

$schema = Get-Content -Path $schemaPath -Raw -Encoding UTF8 | ConvertFrom-Json
Import-Module $modulePath -Force

$sample = [PSCustomObject]@{
    version         = "1.0"
    runId           = "12345678"
    startedAt       = "2026-03-04T00:00:00Z"
    finishedAt      = "2026-03-04T00:00:01Z"
    host            = [PSCustomObject]@{
        hostname = "pc-01"
        os = "Windows"
        psVersion = "7.5.0"
    }
    status          = "OK"
    durationSeconds = 1.0
    tasks           = @(
        [PSCustomObject]@{
            id = 1
            name = "task-1"
            status = "OK"
            duration = 0.5
            errors = @()
            previewOnly = $false
        }
    )
}

$outJson = Join-Path $tmpDir "report.json"
Export-OptimizerReport -ReportData $sample -Format json -Path $outJson | Out-Null
$actual = Get-Content -Path $outJson -Raw -Encoding UTF8 | ConvertFrom-Json

foreach ($required in $schema.required) {
    Assert-True "RS-03: required property '$required' exists" `
        ($null -ne $actual.PSObject.Properties[$required]) `
        "missing $required"
}

Assert-True "RS-04: version const 1.0" ($actual.version -eq $schema.properties.version.const) "actual=$($actual.version)"
Assert-True "RS-05: status enum" ($schema.properties.status.enum -contains $actual.status) "status=$($actual.status)"
Assert-True "RS-06: tasks count >= 1" (@($actual.tasks).Count -ge 1) "tasks=$(@($actual.tasks).Count)"

$task0 = @($actual.tasks)[0]
foreach ($requiredTask in $schema.properties.tasks.items.required) {
    Assert-True "RS-07: task required '$requiredTask' exists" `
        ($null -ne $task0.PSObject.Properties[$requiredTask]) `
        "missing task.$requiredTask"
}
Assert-True "RS-08: task status enum" `
    ($schema.properties.tasks.items.properties.status.enum -contains $task0.status) `
    "task.status=$($task0.status)"

Write-Host ""
Write-Host "PASS: $pass / $($pass + $fail)" -ForegroundColor Green
Write-Host "FAIL: $fail / $($pass + $fail)" -ForegroundColor Yellow

if ($fail -gt 0) { exit 1 }
exit 0
