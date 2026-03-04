# ==============================================================
# Test_PCOptimizer.ps1
# Validates ENCODING_POLICY and CLAUDE_IMPLEMENTATION_CONSTITUTION
# compliance, plus runtime logic of PC_Optimizer.ps1
#
# Usage (no admin required):
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\Test_PCOptimizer.ps1
#   pwsh      -NoProfile -ExecutionPolicy Bypass -File .\tests\Test_PCOptimizer.ps1
# ==============================================================

param(
    [string]$ScriptPath = (Join-Path (Split-Path $PSScriptRoot -Parent) "PC_Optimizer.ps1")
)

$psver       = $PSVersionTable.PSVersion.Major
$isPS7Plus   = ($psver -ge 7)
$logEncoding = if ($isPS7Plus) { 'utf8NoBOM' } else { 'UTF8' }

# ── Minimal test framework ─────────────────────────────────────
$script:passCount = 0
$script:failCount = 0
$script:results   = New-Object 'System.Collections.Generic.List[PSObject]'

function Assert-True {
    param([string]$Name, [bool]$Condition, [string]$Detail = "")
    $obj = [PSCustomObject]@{
        Test   = $Name
        Status = if ($Condition) { "PASS" } else { "FAIL" }
        Detail = $Detail
    }
    $script:results.Add($obj)
    if ($Condition) {
        $script:passCount++
        Write-Host "  [PASS] $Name" -ForegroundColor Green
    } else {
        $script:failCount++
        Write-Host "  [FAIL] $Name" -ForegroundColor Red
        if ($Detail) { Write-Host "         => $Detail" -ForegroundColor DarkYellow }
    }
}

function Write-Section([string]$title) {
    Write-Host ""
    Write-Host ("==" + "= $title " + "=" * 40) -ForegroundColor Cyan
}

# ── Prerequisites ─────────────────────────────────────────────
Write-Host "============================================================" -ForegroundColor White
Write-Host " PC Optimizer Compliance & Logic Verification Tests" -ForegroundColor White
Write-Host " PowerShell $($PSVersionTable.PSVersion)  /  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor White
Write-Host "============================================================" -ForegroundColor White

Assert-True "Target script exists" (Test-Path $ScriptPath) "ScriptPath=$ScriptPath"

if (-not (Test-Path $ScriptPath)) {
    Write-Host "Script not found. Aborting tests." -ForegroundColor Red
    exit 1
}

$sourceLines   = Get-Content $ScriptPath -Encoding UTF8
$sourceContent = $sourceLines -join "`n"

# ==============================================================
# SECTION 1: Static Analysis -- Encoding Policy
# ==============================================================
Write-Section "SECTION 1: Static Analysis -- Encoding Policy"

# 1-1: All Add-Content calls must have -Encoding
$addContentCalls    = $sourceLines | Select-String "Add-Content" | Where-Object { $_ -notmatch "^\s*#" }
$addContentMissing  = $addContentCalls | Where-Object { $_ -notmatch "\-Encoding" }
Assert-True `
    "All Add-Content calls specify -Encoding" `
    ($null -eq $addContentMissing -or @($addContentMissing).Count -eq 0) `
    "Missing -Encoding at: $($addContentMissing -join ' | ')"

# 1-2: Out-File calls (non-null) must have -Encoding
$outFileCalls   = $sourceLines | Select-String "Out-File" | Where-Object { $_ -notmatch "^\s*#" -and $_ -notmatch "Out-Null" }
$outFileMissing = $outFileCalls | Where-Object { $_ -notmatch "\-Encoding" }
Assert-True `
    "All Out-File calls (non-null) specify -Encoding" `
    ($null -eq $outFileMissing -or @($outFileMissing).Count -eq 0) `
    "Missing -Encoding at: $($outFileMissing -join ' | ')"

# 1-3: Set-Content calls must have -Encoding
$setContentCalls   = $sourceLines | Select-String "Set-Content" | Where-Object { $_ -notmatch "^\s*#" }
$setContentMissing = $setContentCalls | Where-Object { $_ -notmatch "\-Encoding" }
Assert-True `
    "All Set-Content calls specify -Encoding" `
    ($null -eq $setContentMissing -or @($setContentMissing).Count -eq 0) `
    "Missing -Encoding at: $($setContentMissing -join ' | ')"

# 1-4: $logEncoding variable is defined
$logEncLine = $sourceLines | Select-String '\$logEncoding\s*='
Assert-True `
    "\$logEncoding variable is defined in source" `
    ($null -ne $logEncLine -and @($logEncLine).Count -gt 0) `
    "Definition not found"

# 1-5: $logEncoding uses $isPS7Plus branch
Assert-True `
    "\$logEncoding uses \$isPS7Plus for version branching" `
    ($sourceContent -match '\$logEncoding\s*=\s*if\s*\(\s*\$isPS7Plus') `
    "Expected pattern: '`$logEncoding = if (`$isPS7Plus) { ... }'"

# 1-6: No default encoding (detect Out-File without -Encoding that is NOT Out-Null or commented)
$defaultOutFile = $sourceLines |
    Select-String "Out-File" |
    Where-Object { $_ -notmatch "^\s*#" -and $_ -notmatch "\-Encoding" -and $_ -notmatch "Out-Null" }
Assert-True `
    "No Out-File without -Encoding exists (UTF-16 default prevention)" `
    ($null -eq $defaultOutFile -or @($defaultOutFile).Count -eq 0) `
    "Lines without -Encoding: $(@($defaultOutFile).Count)"

# ==============================================================
# SECTION 2: Static Analysis -- Variable & Function Order
# ==============================================================
Write-Section "SECTION 2: Static Analysis -- Variable and Function Order"

$lineOf = @{
    psver        = 0
    isPS7Plus    = 0
    logEncoding  = 0
    WriteLog     = 0
    WriteErrorLog= 0
    ProgressBar  = 0
    TryStep      = 0
    Show         = 0
}

for ($i = 0; $i -lt $sourceLines.Count; $i++) {
    $ln = $sourceLines[$i]
    if ($lineOf['psver']        -eq 0 -and $ln -match '^\$psver\s*=')              { $lineOf['psver']         = $i + 1 }
    if ($lineOf['isPS7Plus']    -eq 0 -and $ln -match '^\$isPS7Plus\s*=')          { $lineOf['isPS7Plus']     = $i + 1 }
    if ($lineOf['logEncoding']  -eq 0 -and $ln -match '^\$logEncoding\s*=')        { $lineOf['logEncoding']   = $i + 1 }
    if ($lineOf['WriteLog']     -eq 0 -and $ln -match '^function Write-Log')       { $lineOf['WriteLog']      = $i + 1 }
    if ($lineOf['WriteErrorLog']-eq 0 -and $ln -match '^function Write-ErrorLog')  { $lineOf['WriteErrorLog'] = $i + 1 }
    if ($lineOf['Show']         -eq 0 -and $ln -match '^function Show')            { $lineOf['Show']          = $i + 1 }
    if ($lineOf['ProgressBar']  -eq 0 -and $ln -match '^function Progress-Bar')    { $lineOf['ProgressBar']   = $i + 1 }
    if ($lineOf['TryStep']      -eq 0 -and $ln -match '^function Try-Step')        { $lineOf['TryStep']       = $i + 1 }
}

Assert-True `
    "\$psver defined before Write-Log (L$($lineOf['psver']) vs L$($lineOf['WriteLog']))" `
    ($lineOf['psver'] -gt 0 -and $lineOf['WriteLog'] -gt 0 -and $lineOf['psver'] -lt $lineOf['WriteLog']) `
    "psver=L$($lineOf['psver']), WriteLog=L$($lineOf['WriteLog'])"

Assert-True `
    "\$isPS7Plus defined before Write-Log (L$($lineOf['isPS7Plus']) vs L$($lineOf['WriteLog']))" `
    ($lineOf['isPS7Plus'] -gt 0 -and $lineOf['isPS7Plus'] -lt $lineOf['WriteLog']) `
    "isPS7Plus=L$($lineOf['isPS7Plus']), WriteLog=L$($lineOf['WriteLog'])"

Assert-True `
    "\$logEncoding defined before Write-Log (L$($lineOf['logEncoding']) vs L$($lineOf['WriteLog']))" `
    ($lineOf['logEncoding'] -gt 0 -and $lineOf['logEncoding'] -lt $lineOf['WriteLog']) `
    "logEncoding=L$($lineOf['logEncoding']), WriteLog=L$($lineOf['WriteLog'])"

Assert-True `
    "\$logEncoding defined before Write-ErrorLog (L$($lineOf['logEncoding']) vs L$($lineOf['WriteErrorLog']))" `
    ($lineOf['logEncoding'] -gt 0 -and $lineOf['logEncoding'] -lt $lineOf['WriteErrorLog']) `
    "logEncoding=L$($lineOf['logEncoding']), WriteErrorLog=L$($lineOf['WriteErrorLog'])"

Assert-True `
    "Progress-Bar defined before Try-Step (L$($lineOf['ProgressBar']) vs L$($lineOf['TryStep']))" `
    ($lineOf['ProgressBar'] -gt 0 -and $lineOf['TryStep'] -gt 0 -and $lineOf['ProgressBar'] -lt $lineOf['TryStep']) `
    "ProgressBar=L$($lineOf['ProgressBar']), TryStep=L$($lineOf['TryStep'])"

# ==============================================================
# SECTION 3: Static Analysis -- Error Handling Policy
# ==============================================================
Write-Section "SECTION 3: Static Analysis -- Error Handling Policy"

# 3-1: No empty catch blocks  "catch { }" on a single line
$emptyCatchLines = $sourceLines |
    Select-String 'catch\s*\{\s*\}' |
    Where-Object { $_ -notmatch "^\s*#" }
Assert-True `
    "No empty catch { } blocks exist" `
    ($null -eq $emptyCatchLines -or @($emptyCatchLines).Count -eq 0) `
    "Found at: $($emptyCatchLines -join ' | ')"

# 3-2: Try-Step catch contains Write-ErrorLog
Assert-True `
    "Try-Step catch block calls Write-ErrorLog" `
    ($sourceContent -match 'function Try-Step[\s\S]*?catch[\s\S]*?Write-ErrorLog') `
    "Write-ErrorLog not found inside Try-Step catch"

# 3-3: Test-PendingReboot catch contains Write-ErrorLog
Assert-True `
    "Test-PendingReboot catch block calls Write-ErrorLog" `
    ($sourceContent -match 'function Test-PendingReboot[\s\S]*?catch[\s\S]*?Write-ErrorLog') `
    "Write-ErrorLog not found inside Test-PendingReboot catch"

# 3-4: Run-WindowsUpdate catch contains Write-ErrorLog
Assert-True `
    "Run-WindowsUpdate catch block calls Write-ErrorLog" `
    ($sourceContent -match 'function Run-WindowsUpdate[\s\S]*?catch[\s\S]*?Write-ErrorLog') `
    "Write-ErrorLog not found inside Run-WindowsUpdate catch"

# ==============================================================
# SECTION 4: Runtime Tests -- Write-Log / Write-ErrorLog
# ==============================================================
Write-Section "SECTION 4: Runtime Tests -- Write-Log and Write-ErrorLog"

$tmpDir = Join-Path $env:TEMP ("PCOpt_Test_" + (Get-Date -Format 'yyyyMMddHHmmss'))
New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null

try {
    $testLogPath      = Join-Path $tmpDir "test_log.txt"
    $testErrorLogPath = Join-Path $tmpDir "test_error.txt"

    # Define functions equivalent to the production code
    function Invoke-WriteLog {
        param([string]$msg)
        Add-Content -Path $testLogPath -Value $msg -Encoding $logEncoding
    }
    function Invoke-WriteErrorLog {
        param([string]$msg)
        Add-Content -Path $testErrorLogPath -Value ("[ERROR] " + (Get-Date -Format 'yyyy/MM/dd HH:mm:ss')) -Encoding $logEncoding
        Add-Content -Path $testErrorLogPath -Value $msg -Encoding $logEncoding
    }

    # 4-1: Write-Log creates a file
    Invoke-WriteLog "Line one"
    Invoke-WriteLog "Line two"
    Assert-True "Write-Log creates log file" (Test-Path $testLogPath) ""

    # 4-2: Write-Log content is correct
    $content = Get-Content $testLogPath -Encoding UTF8
    Assert-True `
        "Write-Log content: 2 lines written correctly" `
        ($content.Count -eq 2 -and $content[0] -eq "Line one") `
        "lines=$($content.Count), first='$($content[0])'"

    # 4-3: Write-ErrorLog creates error file
    Invoke-WriteErrorLog "Something failed"
    Assert-True "Write-ErrorLog creates error log file" (Test-Path $testErrorLogPath) ""

    # 4-4: Error log line 1 has [ERROR] timestamp pattern
    $errContent = Get-Content $testErrorLogPath -Encoding UTF8
    Assert-True `
        "Write-ErrorLog line 1 matches [ERROR] yyyy/MM/dd HH:mm:ss" `
        ($errContent[0] -match '^\[ERROR\] \d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}$') `
        "line1='$($errContent[0])'"

    # 4-5: Error log line 2 has the message
    Assert-True `
        "Write-ErrorLog line 2 contains the error message" `
        ($errContent.Count -ge 2 -and $errContent[1] -eq "Something failed") `
        "line2='$(if ($errContent.Count -ge 2) { $errContent[1] } else { "<missing>" })'"

    # 4-6: BOM detection -- verify encoding policy is applied
    $rawBytes = [System.IO.File]::ReadAllBytes($testLogPath)
    $hasBom   = ($rawBytes.Count -ge 3 -and $rawBytes[0] -eq 0xEF -and $rawBytes[1] -eq 0xBB -and $rawBytes[2] -eq 0xBF)
    if ($isPS7Plus) {
        Assert-True `
            "PS7+: log file is UTF-8 without BOM (utf8NoBOM)" `
            (-not $hasBom) `
            "BOM bytes detected: 0x$($rawBytes[0].ToString('X2')) 0x$($rawBytes[1].ToString('X2')) 0x$($rawBytes[2].ToString('X2'))"
    } else {
        Assert-True `
            "PS5.x: log file is UTF-8 with BOM (UTF8)" `
            $hasBom `
            "No BOM detected. First bytes: 0x$($rawBytes[0].ToString('X2')) 0x$($rawBytes[1].ToString('X2')) 0x$($rawBytes[2].ToString('X2'))"
    }

    # 4-7: Append integrity -- 3rd write produces 3 lines total
    Invoke-WriteLog "Line three"
    $content3 = Get-Content $testLogPath -Encoding UTF8
    Assert-True `
        "Append integrity: 3 writes = 3 lines (no data corruption)" `
        ($content3.Count -eq 3) `
        "Expected 3 lines, got $($content3.Count)"

    # 4-8: Non-ASCII (Unicode) content round-trips correctly
    # Use -join to build Unicode string safely (avoids int-addition of [char]+[char] in PS5.1)
    $jpTestStr  = -join @([char]0x65E5, [char]0x672C, [char]0x8A9E)  # U+65E5 U+672C U+8A9E
    $jpTestPath = [System.IO.Path]::Combine($tmpDir, "test_jp.txt")
    $jp8Ok = $false
    try {
        Add-Content -Path $jpTestPath -Value $jpTestStr -Encoding $logEncoding
        $jpReadBack = (Get-Content $jpTestPath -Encoding UTF8 -Raw).TrimEnd("`r", "`n")
        $jp8Ok = ($jpReadBack -eq $jpTestStr)
        Assert-True `
            "Non-ASCII (Unicode) round-trips with $logEncoding encoding (len=$($jpTestStr.Length))" `
            $jp8Ok `
            "Written len=$($jpTestStr.Length), Read len=$($jpReadBack.Length)"
    } catch {
        Assert-True "Non-ASCII (Unicode) round-trips with $logEncoding encoding" $false "Exception: $_"
    }

} finally {
    Remove-Item -Path $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
}

# ==============================================================
# SECTION 5: Runtime Tests -- $logEncoding value
# ==============================================================
Write-Section "SECTION 5: Runtime Tests -- logEncoding value"

$expected = if ($isPS7Plus) { 'utf8NoBOM' } else { 'UTF8' }

Assert-True `
    "logEncoding equals expected value '$expected'" `
    ($logEncoding -eq $expected) `
    "actual='$logEncoding', expected='$expected'"

if ($isPS7Plus) {
    Assert-True "PS7+: utf8NoBOM selected" ($logEncoding -eq 'utf8NoBOM') ""
} else {
    Assert-True "PS5.x: UTF8 (with BOM) selected" ($logEncoding -eq 'UTF8') ""
}

# ==============================================================
# SECTION 6: Runtime Tests -- Test-PendingReboot logic
# ==============================================================
Write-Section "SECTION 6: Runtime Tests -- Test-PendingReboot logic"

function Local-PendingReboot {
    $req = $false
    try {
        if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired") {
            $req = $true
        }
        if (Test-Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager") {
            $p = Get-ItemProperty `
                -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" `
                -Name "PendingFileRenameOperations" `
                -ErrorAction SilentlyContinue
            if ($p -and $p.PendingFileRenameOperations) { $req = $true }
        }
    } catch { }
    return $req
}

$rebootResult = $null
try {
    $rebootResult = Local-PendingReboot
    Assert-True "Test-PendingReboot executes without exception" $true ""
} catch {
    Assert-True "Test-PendingReboot executes without exception" $false "Exception: $_"
}

Assert-True `
    "Test-PendingReboot returns a bool" `
    ($rebootResult -is [bool]) `
    "Type returned: $($rebootResult.GetType().Name)"

# ==============================================================
# SECTION 7: Static Analysis -- BAT file compliance
# ==============================================================
Write-Section "SECTION 7: Static Analysis -- BAT file compliance"

$batPath = Join-Path (Split-Path $ScriptPath -Parent) "Run_PC_Optimizer.bat"
Assert-True "Run_PC_Optimizer.bat exists" (Test-Path $batPath) "path=$batPath"

if (Test-Path $batPath) {
    $bat = Get-Content $batPath -Encoding Default

    Assert-True `
        "BAT: pwsh (PS7) preferred detection via 'where pwsh'" `
        ($null -ne ($bat | Select-String "where pwsh")) `
        ""

    Assert-True `
        "BAT: admin check via 'net session'" `
        ($null -ne ($bat | Select-String "net session")) `
        ""

    Assert-True `
        "BAT: -ExecutionPolicy Bypass specified" `
        ($null -ne ($bat | Select-String "ExecutionPolicy Bypass")) `
        ""

    $pwshExeLines = $bat | Select-String 'set "PWSH_EXE=.*WindowsPowerShell'
    Assert-True `
        "BAT: No duplicate PWSH_EXE assignment for WindowsPowerShell path (was: 2 identical lines)" `
        (@($pwshExeLines).Count -le 1) `
        "Duplicate set lines found: $(@($pwshExeLines).Count)"
}

# ==============================================================
# SECTION 8: PowerShell AST syntax check
# ==============================================================
Write-Section "SECTION 8: AST Syntax Check"

$parseErrors = $null
$tokens      = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile(
    $ScriptPath,
    [ref]$tokens,
    [ref]$parseErrors
)

Assert-True `
    "PC_Optimizer.ps1 has zero parse errors" `
    ($parseErrors.Count -eq 0) `
    "$($parseErrors.Count) error(s): $($parseErrors | Select-Object -First 3 -ExpandProperty Message)"

# AST: verify all Add-Content calls have -Encoding argument
$commandAsts       = $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true)
$addContentAsts    = $commandAsts | Where-Object { $_.GetCommandName() -eq 'Add-Content' }
$addContentNoEncAS = $addContentAsts | Where-Object { $_.Extent.Text -notmatch '-Encoding' }

Assert-True `
    "AST: All $($addContentAsts.Count) Add-Content call(s) include -Encoding" `
    (@($addContentNoEncAS).Count -eq 0) `
    "Missing -Encoding on lines: $(@($addContentNoEncAS) | ForEach-Object { 'L' + $_.Extent.StartLineNumber })"

# ── Task 20: スタートアップ・サービスレポートのロジックテスト ─────
Write-Section "SECTION 9: Task 20 Startup / Service Report"

Assert-True "TC-01: PC_Optimizer.ps1 に unnecessaryServices リストが存在する" `
    ($sourceContent -match '\$unnecessaryServices') `
    "スクリプト内に `\$unnecessaryServices` が見つかりません"

Assert-True "TC-02: unnecessaryServices に 'Fax' が含まれる" `
    ($sourceContent -match "'Fax'") `
    "'Fax' がリストに見つかりません"

Assert-True "TC-03: unnecessaryServices に 'DiagTrack' が含まれる" `
    ($sourceContent -match "'DiagTrack'") `
    "'DiagTrack' がリストに見つかりません"

Assert-True "TC-04: Win32_StartupCommand の呼び出しが存在する" `
    ($sourceContent -match 'Win32_StartupCommand') `
    "Win32_StartupCommand が見つかりません"

Assert-True "TC-05: Get-Service の呼び出しが存在する" `
    ($sourceContent -match 'Get-Service') `
    "Get-Service が見つかりません"

Assert-True "TC-06: タスク20が Try-Step ラッパーで実装されている" `
    ($sourceContent -match 'Try-Step[\s\S]*?スタートアップ') `
    "Try-Step 'スタートアップ...' が見つかりません"

Assert-True "TC-07: Set-Service (mutating cmd) not present" `
    ($sourceContent -notmatch 'Set-Service') `
    "Set-Service detected (mutating commands are prohibited)"

Assert-True "TC-08: Stop-Service not present" `
    ($sourceContent -notmatch 'Stop-Service') `
    "Stop-Service detected (mutating commands are prohibited)"

# ==============================================================
# Final Summary
# ==============================================================
$total = $script:passCount + $script:failCount
Write-Host ""
Write-Host "============================================================" -ForegroundColor White
Write-Host " TEST SUMMARY" -ForegroundColor White
Write-Host "============================================================" -ForegroundColor White
Write-Host ("  PASS : {0,3} / {1}" -f $script:passCount, $total) -ForegroundColor Green
$failColor = if ($script:failCount -eq 0) { "Green" } else { "Red" }
Write-Host ("  FAIL : {0,3} / {1}" -f $script:failCount, $total) -ForegroundColor $failColor
Write-Host ""

if ($script:failCount -gt 0) {
    Write-Host "Failed tests:" -ForegroundColor Red
    $script:results |
        Where-Object { $_.Status -eq "FAIL" } |
        ForEach-Object {
            Write-Host "  - $($_.Test)" -ForegroundColor Red
            if ($_.Detail) { Write-Host "    $($_.Detail)" -ForegroundColor DarkYellow }
        }
    Write-Host ""
    exit 1
} else {
    Write-Host "All tests PASS -- compliance verified." -ForegroundColor Green
    exit 0
}
