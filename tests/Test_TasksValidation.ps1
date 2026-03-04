# ==============================================================
# Test_TasksValidation.ps1
# Runtime validation tests for invalid -Tasks arguments
# ==============================================================

param(
    [string]$RepoRoot = (Split-Path $PSScriptRoot -Parent)
)

$scriptPath = Join-Path $RepoRoot "PC_Optimizer.ps1"
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

Assert-True "TV-01: script exists" (Test-Path $scriptPath) $scriptPath
if ($fail -gt 0) { exit 1 }

$cases = @(
    @{ Name = "TV-02: -Tasks '1,,,' returns InvalidArgs(3)"; Tasks = "1,,," },
    @{ Name = "TV-03: -Tasks '1-' returns InvalidArgs(3)";   Tasks = "1-"   },
    @{ Name = "TV-04: -Tasks '-3' returns InvalidArgs(3)";   Tasks = "-3"   },
    @{ Name = "TV-05: -Tasks '1-3-5' returns InvalidArgs(3)"; Tasks = "1-3-5" },
    @{ Name = "TV-06: -Tasks '1-5,3' returns InvalidArgs(3)"; Tasks = "1-5,3" }
)

foreach ($c in $cases) {
    & powershell -NoProfile -ExecutionPolicy Bypass -File $scriptPath `
        -NonInteractive -WhatIf -NoRebootPrompt -Tasks $c.Tasks | Out-Null
    $ec = $LASTEXITCODE
    Assert-True $c.Name ($ec -eq 3) "exit=$ec tasks=$($c.Tasks)"
}

& powershell -NoProfile -ExecutionPolicy Bypass -File $scriptPath `
    -NonInteractive -WhatIf -NoRebootPrompt -Tasks "1, 2" | Out-Null
$validSpaceExit = $LASTEXITCODE
Assert-True "TV-07: -Tasks '1, 2' (space mixed) exits 0" ($validSpaceExit -eq 0) "exit=$validSpaceExit"

Write-Host ""
Write-Host "PASS: $pass / $($pass + $fail)" -ForegroundColor Green
Write-Host "FAIL: $fail / $($pass + $fail)" -ForegroundColor Yellow

if ($fail -gt 0) { exit 1 }
exit 0
