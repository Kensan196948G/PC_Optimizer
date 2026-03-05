param(
    [ValidateSet("powershell","pwsh")]
    [string]$Engine = "powershell",
    [int]$Iterations = 3,
    [int]$MaxAllowedP95Seconds = 120
)
Set-StrictMode -Version Latest

$repoRoot = Split-Path $PSScriptRoot -Parent
$scriptPath = Join-Path $repoRoot "PC_Optimizer.ps1"
$reportsDir = Join-Path $repoRoot "reports"
$resultDir = Join-Path $repoRoot "test-results"
if (-not (Test-Path $resultDir)) { New-Item -ItemType Directory -Path $resultDir -Force | Out-Null }

$durations = @()
for ($i = 1; $i -le $Iterations; $i++) {
    Write-Host "[LOAD] iteration $i/$Iterations"
    $sw = [Diagnostics.Stopwatch]::StartNew()
    & $Engine -NoProfile -ExecutionPolicy Bypass -File $scriptPath -NonInteractive -WhatIf -NoRebootPrompt -Mode diagnose -ExecutionProfile agent-teams -Tasks "20" -FailureMode continue | Out-Null
    $sw.Stop()
    $sec = [math]::Round($sw.Elapsed.TotalSeconds, 3)
    $durations += $sec
    if ($LASTEXITCODE -ne 0) {
        throw "Load iteration failed: engine=$Engine iteration=$i exitCode=$LASTEXITCODE"
    }
}

$sorted = @($durations | Sort-Object)
$idx = [math]::Ceiling(@($sorted).Count * 0.95) - 1
if ($idx -lt 0) { $idx = 0 }
$p95 = [double]$sorted[$idx]
$avg = [math]::Round((($durations | Measure-Object -Average).Average), 3)

$latestAudit = Get-ChildItem -Path $reportsDir -Filter "Audit_Run_*.json" -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTimeUtc | Select-Object -Last 1
if (-not $latestAudit) { throw "Audit file not found after load test." }
$auditObj = Get-Content -Path $latestAudit.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
if (-not $auditObj.agentTeams) { throw "agentTeams section missing after load test." }

$summary = [PSCustomObject]@{
    engine = $Engine
    iterations = $Iterations
    durationsSeconds = @($durations)
    averageSeconds = $avg
    p95Seconds = $p95
    maxAllowedP95Seconds = $MaxAllowedP95Seconds
    generatedAt = (Get-Date).ToString("s")
}
$summaryPath = Join-Path $resultDir ("agent_teams_load_{0}.json" -f $Engine)
$summary | ConvertTo-Json -Depth 6 | Set-Content -Path $summaryPath -Encoding utf8

if ($p95 -gt $MaxAllowedP95Seconds) {
    throw "Load threshold failed: p95=$p95 sec > $MaxAllowedP95Seconds sec"
}

Write-Host "[LOAD] success engine=$Engine avg=$avg sec p95=$p95 sec"
