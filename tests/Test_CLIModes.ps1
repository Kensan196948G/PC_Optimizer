# ==============================================================
# Test_CLIModes.ps1
# Static regression checks for v4.0 CLI integration
# ==============================================================

param(
    [string]$ScriptPath = (Join-Path (Split-Path $PSScriptRoot -Parent) "PC_Optimizer.ps1")
)

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

if (-not (Test-Path $ScriptPath)) {
    Write-Host "Script not found: $ScriptPath" -ForegroundColor Red
    exit 1
}

$content = Get-Content -Path $ScriptPath -Encoding UTF8 -Raw

Write-Host "============================================================" -ForegroundColor White
Write-Host " CLI Modes Regression Tests" -ForegroundColor White
Write-Host "============================================================" -ForegroundColor White

# 1) config loading
Assert-True "CM-01: imports Common module" `
    ($content -match 'Import-Module\s+\$commonModulePath') `
    "Import-Module with commonModulePath not found"

Assert-True "CM-02: calls Get-OptimizerConfig" `
    ($content -match 'Get-OptimizerConfig\s+-Path\s+\$ConfigPath') `
    "Get-OptimizerConfig -Path $ConfigPath not found"

# 2) NonInteractive mode
Assert-True "NI-01: has -NonInteractive parameter" `
    ($content -match '\[switch\]\$NonInteractive') `
    "param [switch]$NonInteractive not found"

Assert-True "NI-02: has Get-UserChoice function" `
    ($content -match 'function\s+Get-UserChoice') `
    "function Get-UserChoice not found"

Assert-True "NI-03: skips exit wait when non-interactive" `
    ($content -match 'if\s*\(\$script:IsNonInteractive\)\s*\{\s*Write-Log\s*"\[NonInteractive\]\s*') `
    "non-interactive exit wait skip branch not found"

Assert-True "NI-04: retains Read-Host for interactive path" `
    ($content -match 'Read-Host\s+"Enter\s+.*"') `
    "interactive Read-Host end prompt not found"

Assert-True "NI-05: has -Tasks parameter" `
    ($content -match '\[string\]\$Tasks\s*=\s*"all"') `
    "param [string]$Tasks = \"all\" not found"

Assert-True "NI-06: non-selected tasks are recorded as SKIP" `
    ($content -match 'TaskFiltered skip') `
    "TaskFiltered skip record not found"

# 3) WhatIf mode
Assert-True "WI-01: has -WhatIf parameter" `
    ($content -match '\[switch\]\$WhatIf') `
    "param [switch]$WhatIf not found"

Assert-True "WI-02: Try-Step contains WhatIf control" `
    ($content -match 'if\s*\(\$script:IsWhatIfMode\s*-and\s*-not\s*\(Test-ReadOnlyTask') `
    "Try-Step WhatIf condition not found"

Assert-True "WI-03: writes SKIP status in WhatIf path" `
    ($content -match 'Status\s*=\s*"SKIP"') `
    'Status = "SKIP" not found'

Assert-True "WI-04: has -ExportDeletedPaths parameter" `
    ($content -match '\[string\]\$ExportDeletedPaths') `
    "param [string]$ExportDeletedPaths not found"

Assert-True "FM-01: has -FailureMode parameter" `
    ($content -match '\[string\]\$FailureMode') `
    "param [string]$FailureMode not found"

Assert-True "FM-02: uses exit code map 0/1/2/3/4" `
    ($content -match 'Success\s*=\s*0' -and
     $content -match 'Partial\s*=\s*1' -and
     $content -match 'Fatal\s*=\s*2' -and
     $content -match 'InvalidArgs\s*=\s*3' -and
     $content -match 'Permission\s*=\s*4') `
    "exit code map not found"

Write-Host ""
Write-Host "PASS: $pass / $($pass + $fail)" -ForegroundColor Green
Write-Host "FAIL: $fail / $($pass + $fail)" -ForegroundColor Yellow

if ($fail -gt 0) { exit 1 }
exit 0
