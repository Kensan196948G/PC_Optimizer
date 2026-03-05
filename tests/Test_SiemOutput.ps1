#Requires -Version 5.1
<#
.SYNOPSIS
    SIEM 出力機能の単体テスト（Test_SiemOutput.ps1）
    Orchestration.psm1 の SIEM 出力関数 (Convert-HookEntryToSiemLine / Export-HookSiemLines) を検証する。
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$psver = $PSVersionTable.PSVersion.Major
$isPS7Plus = ($psver -ge 7)

# ========== Helper ==========
$script:PassCount = 0
$script:FailCount = 0

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if ($Condition) {
        Write-Host "  [PASS] $Message" -ForegroundColor Green
        $script:PassCount++
    } else {
        Write-Host "  [FAIL] $Message" -ForegroundColor Red
        $script:FailCount++
    }
}

function Assert-Contains {
    param([string]$Text, [string]$Pattern, [string]$Message)
    Assert-True -Condition ($Text -match $Pattern) -Message $Message
}

function Assert-NotContains {
    param([string]$Text, [string]$Pattern, [string]$Message)
    Assert-True -Condition ($Text -notmatch $Pattern) -Message $Message
}

# ========== Load Orchestration module ==========
$orchPath = Join-Path $PSScriptRoot '..\modules\Orchestration.psm1'
if (-not (Test-Path $orchPath)) {
    Write-Error "Orchestration.psm1 not found at: $orchPath"
    exit 1
}
Import-Module $orchPath -Force

# ========== Test helpers for private functions (invoke by calling exported ConvertHook/Export) ==========
# Note: Convert-HookEntryToSiemLine and Export-HookSiemLines are not exported by default.
# We test them via the module's exported Invoke-AgentHookEvent and by calling the module functions directly.
# We use module scope invocation.

$orchModule = Get-Module -Name 'Orchestration'
if (-not $orchModule) {
    Write-Error "Orchestration module not loaded."
    exit 1
}

# Helper to call non-exported function via module scope
function Invoke-OrchestratorPrivate {
    param([string]$FunctionName, [hashtable]$Params)
    & $orchModule { param($fn, $p) & $fn @p } $FunctionName $Params
}

Write-Host ""
Write-Host "=== SI-01: Convert-HookEntryToSiemLine - JSONL format ===" -ForegroundColor Cyan

$entry = [PSCustomObject]@{
    event    = "post_task"
    action   = "notify"
    status   = "Success"
    detail   = "test detail"
    finishedAt = "2026-01-01T10:00:00"
    hookPayload = [PSCustomObject]@{
        transactionId = "tx-001"
        nodeId        = "collector.security"
        runId         = "run-001"
    }
}

$lineJsonl = & $orchModule { param($e) Convert-HookEntryToSiemLine -Entry $e -Format "jsonl" } $entry
Assert-True -Condition ($null -ne $lineJsonl) -Message "SI-01-1: JSONL line is not null"
Assert-Contains -Text "$lineJsonl" -Pattern '"event"' -Message "SI-01-2: JSONL contains event key"
Assert-Contains -Text "$lineJsonl" -Pattern '"post_task"' -Message "SI-01-3: JSONL contains event value"
Assert-Contains -Text "$lineJsonl" -Pattern '"status"' -Message "SI-01-4: JSONL contains status key"

Write-Host ""
Write-Host "=== SI-02: Convert-HookEntryToSiemLine - CEF format ===" -ForegroundColor Cyan

$lineCef = & $orchModule { param($e) Convert-HookEntryToSiemLine -Entry $e -Format "cef" } $entry
Assert-True -Condition ($null -ne $lineCef) -Message "SI-02-1: CEF line is not null"
Assert-Contains -Text "$lineCef" -Pattern "^CEF:" -Message "SI-02-2: CEF line starts with CEF:"
Assert-Contains -Text "$lineCef" -Pattern "PC_Optimizer" -Message "SI-02-3: CEF contains vendor"
Assert-Contains -Text "$lineCef" -Pattern "tx-001" -Message "SI-02-4: CEF contains transactionId"
Assert-Contains -Text "$lineCef" -Pattern "collector.security" -Message "SI-02-5: CEF contains nodeId"

Write-Host ""
Write-Host "=== SI-03: Convert-HookEntryToSiemLine - LEEF format ===" -ForegroundColor Cyan

$lineLeef = & $orchModule { param($e) Convert-HookEntryToSiemLine -Entry $e -Format "leef" } $entry
Assert-True -Condition ($null -ne $lineLeef) -Message "SI-03-1: LEEF line is not null"
Assert-Contains -Text "$lineLeef" -Pattern "^LEEF:" -Message "SI-03-2: LEEF line starts with LEEF:"
Assert-Contains -Text "$lineLeef" -Pattern "sev=" -Message "SI-03-3: LEEF contains sev"
Assert-Contains -Text "$lineLeef" -Pattern "tx=tx-001" -Message "SI-03-4: LEEF contains tx"

Write-Host ""
Write-Host "=== SI-04: Convert-HookEntryToSiemLine - Failed status severity ===" -ForegroundColor Cyan

$failEntry = [PSCustomObject]@{
    event    = "on_error"
    action   = "notify"
    status   = "Failed"
    detail   = "error occurred"
    finishedAt = "2026-01-01T10:01:00"
    hookPayload = [PSCustomObject]@{
        transactionId = "tx-002"
        nodeId        = "analyzer.security"
        runId         = "run-002"
    }
}

$failCef = & $orchModule { param($e) Convert-HookEntryToSiemLine -Entry $e -Format "cef" } $failEntry
Assert-Contains -Text "$failCef" -Pattern "\|7\|" -Message "SI-04-1: Failed entry has severity 7 in CEF"

$okEntry = $entry.PSObject.Copy()
$okCef = & $orchModule { param($e) Convert-HookEntryToSiemLine -Entry $e -Format "cef" } $entry
Assert-Contains -Text "$okCef" -Pattern "\|3\|" -Message "SI-04-2: Success entry has severity 3 in CEF"

Write-Host ""
Write-Host "=== SI-05: Export-HookSiemLines - creates output files ===" -ForegroundColor Cyan

$tempDir = Join-Path $env:TEMP ("SiemTest_{0}" -f [guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

try {
    $siemConfig = [PSCustomObject]@{
        siem = [PSCustomObject]@{
            enabled   = $true
            formats   = @("jsonl", "cef", "leef")
            outputDir = "siem_out"
        }
    }

    $entries = @($entry, $failEntry)
    & $orchModule {
        param($en, $hc, $ld, $ri)
        Export-HookSiemLines -Entries $en -HooksConfig $hc -LogsDir $ld -RunId $ri
    } $entries $siemConfig $tempDir "runSI05"

    $siemOutDir = Join-Path $tempDir "siem_out"
    $files = @(Get-ChildItem -Path $siemOutDir -File -ErrorAction SilentlyContinue)
    Assert-True -Condition ($files.Count -eq 3) -Message "SI-05-1: 3 output files created (jsonl, cef, leef)"

    $jsonlFile = @($files | Where-Object { $_.Extension -eq ".jsonl" })[0]
    $cefFile   = @($files | Where-Object { $_.Extension -eq ".cef" })[0]
    $leefFile  = @($files | Where-Object { $_.Extension -eq ".leef" })[0]

    Assert-True -Condition ($null -ne $jsonlFile) -Message "SI-05-2: JSONL file exists"
    Assert-True -Condition ($null -ne $cefFile)   -Message "SI-05-3: CEF file exists"
    Assert-True -Condition ($null -ne $leefFile)  -Message "SI-05-4: LEEF file exists"

    if ($null -ne $jsonlFile) {
        $jsonlContent = Get-Content -Path $jsonlFile.FullName -Raw
        Assert-Contains -Text $jsonlContent -Pattern '"post_task"' -Message "SI-05-5: JSONL file contains event data"
    }
    if ($null -ne $cefFile) {
        $cefContent = Get-Content -Path $cefFile.FullName -Raw
        Assert-Contains -Text $cefContent -Pattern "^CEF:" -Message "SI-05-6: CEF file contains CEF format data"
    }
} finally {
    Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host ""
Write-Host "=== SI-06: Export-HookSiemLines - disabled siem skips output ===" -ForegroundColor Cyan

$tempDir2 = Join-Path $env:TEMP ("SiemTest_{0}" -f [guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Path $tempDir2 -Force | Out-Null

try {
    $siemConfigDisabled = [PSCustomObject]@{
        siem = [PSCustomObject]@{
            enabled   = $false
            formats   = @("jsonl")
            outputDir = "siem_disabled"
        }
    }

    & $orchModule {
        param($en, $hc, $ld, $ri)
        Export-HookSiemLines -Entries $en -HooksConfig $hc -LogsDir $ld -RunId $ri
    } @($entry) $siemConfigDisabled $tempDir2 "runSI06"

    $siemDisabledDir = Join-Path $tempDir2 "siem_disabled"
    $noFiles = @(Get-ChildItem -Path $siemDisabledDir -File -ErrorAction SilentlyContinue)
    Assert-True -Condition ($noFiles.Count -eq 0) -Message "SI-06-1: No files created when siem is disabled"
} finally {
    Remove-Item -Path $tempDir2 -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host ""
Write-Host "=== SI-07: Export-HookSiemLines - empty entries produces no output ===" -ForegroundColor Cyan

$tempDir3 = Join-Path $env:TEMP ("SiemTest_{0}" -f [guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Path $tempDir3 -Force | Out-Null

try {
    $siemConfigEnabled = [PSCustomObject]@{
        siem = [PSCustomObject]@{
            enabled   = $true
            formats   = @("jsonl")
            outputDir = "siem_empty"
        }
    }

    & $orchModule {
        param($en, $hc, $ld, $ri)
        Export-HookSiemLines -Entries $en -HooksConfig $hc -LogsDir $ld -RunId $ri
    } @() $siemConfigEnabled $tempDir3 "runSI07"

    $siemEmptyDir = Join-Path $tempDir3 "siem_empty"
    $emptyFiles = @(Get-ChildItem -Path $siemEmptyDir -File -ErrorAction SilentlyContinue)
    Assert-True -Condition ($emptyFiles.Count -eq 0) -Message "SI-07-1: No files created for empty entries"
} finally {
    Remove-Item -Path $tempDir3 -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host ""
Write-Host "=== SI-08: HooksConfig without siem key - no error ===" -ForegroundColor Cyan

try {
    $noSiemConfig = [PSCustomObject]@{
        pre_task = @()
        on_error = @()
    }
    & $orchModule {
        param($en, $hc, $ld, $ri)
        Export-HookSiemLines -Entries $en -HooksConfig $hc -LogsDir $ld -RunId $ri
    } @($entry) $noSiemConfig "." "runSI08"
    Assert-True -Condition $true -Message "SI-08-1: No exception when HooksConfig lacks siem key"
} catch {
    Assert-True -Condition $false -Message "SI-08-1: Exception when HooksConfig lacks siem key: $_"
}

# ========== Summary ==========
Write-Host ""
Write-Host "=============================" -ForegroundColor White
Write-Host "  SIEM Test Summary" -ForegroundColor White
Write-Host "=============================" -ForegroundColor White
Write-Host "  PASS: $($script:PassCount)" -ForegroundColor Green
Write-Host "  FAIL: $($script:FailCount)" -ForegroundColor Red
Write-Host "=============================" -ForegroundColor White

if ($script:FailCount -gt 0) {
    Write-Host "[RESULT] FAILED" -ForegroundColor Red
    exit 1
} else {
    Write-Host "[RESULT] ALL PASSED" -ForegroundColor Green
    exit 0
}
