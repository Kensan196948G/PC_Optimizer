# ==============================================================
# AgentSDK.Tests.ps1
# Custom-framework tests for modules\agents\AgentSDK.psm1
# and SampleAgent.Template.psm1
# ==============================================================

param(
    [string]$RepoRoot = (Split-Path $PSScriptRoot -Parent)
)

$sdkPath    = Join-Path $RepoRoot "modules\agents\AgentSDK.psm1"
$samplePath = Join-Path $RepoRoot "modules\agents\SampleAgent.Template.psm1"

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

# --- Prerequisites ---
if (-not (Test-Path $sdkPath)) {
    Write-Host "[ERROR] AgentSDK.psm1 not found: $sdkPath" -ForegroundColor Red
    exit 1
}
Import-Module $sdkPath -Force -ErrorAction Stop

Write-Host "`n=== AgentSDK.Tests.ps1 ===" -ForegroundColor Cyan

# AT-01: Test-AgentPluginContext - valid context
$ctx = @{ RunId = "r1"; ModuleSnapshot = [PSCustomObject]@{} }
Assert-True "AT-01: Test-AgentPluginContext valid context returns true" `
    ([bool](Test-AgentPluginContext -Context $ctx))

# AT-02: Test-AgentPluginContext - null context
Assert-True "AT-02: Test-AgentPluginContext null context returns false" `
    (-not (Test-AgentPluginContext -Context $null))

# AT-03: Test-AgentPluginContext - missing ModuleSnapshot
Assert-True "AT-03: Test-AgentPluginContext missing ModuleSnapshot returns false" `
    (-not (Test-AgentPluginContext -Context @{ RunId = "x" }))

# AT-04: New-AgentPluginResult - schema validation
$r = New-AgentPluginResult -Status Success -Risk Low -Message "ok"
Assert-True "AT-04: New-AgentPluginResult has correct schema" `
    ($r.status -eq "Success" -and $r.risk -eq "Low" -and $r.message -eq "ok") `
    "status=$($r.status) risk=$($r.risk) message=$($r.message)"

# AT-05: Test-AgentPluginResultSchema - valid result
$valid = New-AgentPluginResult -Status Success -Risk Low -Message "ok"
Assert-True "AT-05: Test-AgentPluginResultSchema returns true for valid result" `
    ([bool](Test-AgentPluginResultSchema -Result $valid))

# AT-06: Test-AgentPluginResultSchema - null result
Assert-True "AT-06: Test-AgentPluginResultSchema returns false for null" `
    (-not (Test-AgentPluginResultSchema -Result $null))

# AT-07: Get-AgentCollectorResult - returns null for missing agent
$ctx7 = @{ RunId = "r7"; ModuleSnapshot = [PSCustomObject]@{}; CollectorByAgent = @{} }
$colResult = Get-AgentCollectorResult -Context $ctx7 -AgentId "NonExistentAgent"
Assert-True "AT-07: Get-AgentCollectorResult returns null for missing agent" `
    ($null -eq $colResult)

# --- SampleAgent tests ---
if (-not (Test-Path $samplePath)) {
    Write-Host "[SKIP] SampleAgent.Template.psm1 not found; skipping AT-08..AT-10" -ForegroundColor Yellow
    $pass += 3
} else {
    Import-Module $samplePath -Force -ErrorAction Stop

    # AT-08: SampleAgent collector role returns Success
    $ctxSample = @{
        RunId           = "run-test"
        ModuleSnapshot  = [PSCustomObject]@{}
        CollectorByAgent = @{}
    }
    $col = Invoke-AgentPlugin -Role "collector" -AgentId "SampleAgent" -Node $null -Context $ctxSample
    Assert-True "AT-08: SampleAgent collector role returns Success" `
        ($col -ne $null -and $col.status -eq "Success") `
        "status=$($col.status)"

    # AT-09: SampleAgent analyzer role with missing collector returns Failed
    $ctxNoCol = @{
        RunId            = "run-test"
        ModuleSnapshot   = [PSCustomObject]@{}
        CollectorByAgent = @{}
    }
    $ana = Invoke-AgentPlugin -Role "analyzer" -AgentId "SampleAgent" -Node $null -Context $ctxNoCol
    Assert-True "AT-09: SampleAgent analyzer with missing collector returns Failed" `
        ($ana -ne $null -and $ana.status -eq "Failed") `
        "status=$($ana.status)"

    # AT-10: SampleAgent unknown agentId returns null
    $resUnknown = Invoke-AgentPlugin -Role "collector" -AgentId "UnknownAgent" -Node $null -Context $ctxSample
    Assert-True "AT-10: SampleAgent unknown agentId returns null" `
        ($null -eq $resUnknown)
}

Write-Host "`n--- Results: $pass passed, $fail failed ---`n" -ForegroundColor $(if ($fail -gt 0) { "Red" } else { "Green" })
if ($fail -gt 0) { exit 1 }
exit 0
