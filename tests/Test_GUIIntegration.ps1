# ==============================================================
# Test_GUIIntegration.ps1
# Static and runtime validation for GUI + CLI coexistence
# ==============================================================

param(
    [string]$RepoRoot = (Split-Path $PSScriptRoot -Parent)
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

$enginePath = Join-Path $RepoRoot "PC_Optimizer.ps1"
$taskCatalogPath = Join-Path $RepoRoot "modules\TaskCatalog.psm1"
$guiScriptPath = Join-Path $RepoRoot "gui\PCOptimizer.Gui.ps1"
$guiXamlPath = Join-Path $RepoRoot "gui\MainWindow.xaml"
$guiBatPath = Join-Path $RepoRoot "Run_PC_Optimizer_GUI.bat"

Write-Host "============================================================" -ForegroundColor White
Write-Host " GUI Integration Validation" -ForegroundColor White
Write-Host "============================================================" -ForegroundColor White

Assert-True "GUI-01: GUI launcher exists" (Test-Path $guiBatPath) $guiBatPath
Assert-True "GUI-02: GUI script exists" (Test-Path $guiScriptPath) $guiScriptPath
Assert-True "GUI-03: GUI XAML exists" (Test-Path $guiXamlPath) $guiXamlPath
Assert-True "GUI-04: task catalog module exists" (Test-Path $taskCatalogPath) $taskCatalogPath

$guiScriptContent = Get-Content -Path $guiScriptPath -Raw -Encoding UTF8
$guiBatContent = Get-Content -Path $guiBatPath -Raw -Encoding UTF8

Assert-True "GUI-05: GUI imports task catalog" `
    ($guiScriptContent -match 'Import-Module .*TaskCatalog\.psm1') `
    "TaskCatalog import missing"

Assert-True "GUI-06: GUI launches PC_Optimizer.ps1 engine" `
    ($guiScriptContent -match 'PC_Optimizer\.ps1') `
    "engine path reference missing"

Assert-True "GUI-07: GUI launcher targets gui script" `
    ($guiBatContent -match 'gui\\PCOptimizer\.Gui\.ps1') `
    "GUI launcher target missing"

$taskCount = & powershell -NoProfile -ExecutionPolicy Bypass -Command "Import-Module '$taskCatalogPath' -Force; @(Get-PCOptimizerTaskCatalog).Count"
Assert-True "GUI-08: task catalog returns 20 tasks" ($taskCount -eq 20) "count=$taskCount"

$selection = & powershell -NoProfile -ExecutionPolicy Bypass -Command "Import-Module '$taskCatalogPath' -Force; ConvertTo-PCOptimizerTaskSelection -TaskIds @(1..20)"
Assert-True "GUI-09: full task set collapses to all" ($selection -eq 'all') "selection=$selection"

& powershell -NoProfile -ExecutionPolicy Bypass -Command "Import-Module '$taskCatalogPath' -Force; ConvertTo-PCOptimizerTaskSelection -TaskIds @(0,21) | Out-Null" | Out-Null
$invalidSelectionExit = $LASTEXITCODE
Assert-True "GUI-09b: invalid task ids are rejected in task catalog" ($invalidSelectionExit -ne 0) "exit=$invalidSelectionExit"

$argList = & powershell -NoProfile -ExecutionPolicy Bypass -Command "Import-Module '$taskCatalogPath' -Force; @(New-PCOptimizerArgumentList -Mode diagnose -ExecutionProfile classic -FailureMode continue -TaskIds @(20) -NonInteractive -NoRebootPrompt -WhatIfMode -EmitUiEvents) -join '|'"
Assert-True "GUI-10: argument builder includes GUI-safe flags" `
    ($argList -match '\-NonInteractive' -and $argList -match '\-NoRebootPrompt' -and $argList -match '\-WhatIf' -and $argList -match '\-EmitUiEvents') `
    $argList

$tokens = $null
$errors = $null
[System.Management.Automation.Language.Parser]::ParseFile($enginePath, [ref]$tokens, [ref]$errors) | Out-Null
Assert-True "GUI-11: engine script parses cleanly" (@($errors).Count -eq 0) (@($errors | ForEach-Object { $_.Message }) -join '; ')

$tokens = $null
$errors = $null
[System.Management.Automation.Language.Parser]::ParseFile($guiScriptPath, [ref]$tokens, [ref]$errors) | Out-Null
Assert-True "GUI-12: GUI script parses cleanly" (@($errors).Count -eq 0) (@($errors | ForEach-Object { $_.Message }) -join '; ')

$emitOutput = & powershell -NoProfile -ExecutionPolicy Bypass -File $enginePath -NonInteractive -WhatIf -NoRebootPrompt -Tasks "20" -EmitUiEvents 2>&1
$emitExit = $LASTEXITCODE
$emitJoined = ($emitOutput | Out-String)
Assert-True "GUI-13: UI event run exits 0" ($emitExit -eq 0) "exit=$emitExit"
Assert-True "GUI-14: UI event prefix emitted" ($emitJoined -match '##PCOPT_UI##') "prefix missing"
Assert-True "GUI-15: run_start emitted" ($emitJoined -match '"type":"run_start"') "run_start missing"
Assert-True "GUI-16: task_start emitted" ($emitJoined -match '"type":"task_start"') "task_start missing"
Assert-True "GUI-17: task_finish emitted" ($emitJoined -match '"type":"task_finish"') "task_finish missing"
Assert-True "GUI-18: run_complete emitted" ($emitJoined -match '"type":"run_complete"') "run_complete missing"

Write-Host ""
Write-Host "PASS: $pass / $($pass + $fail)" -ForegroundColor Green
Write-Host "FAIL: $fail / $($pass + $fail)" -ForegroundColor Yellow

if ($fail -gt 0) { exit 1 }
exit 0
