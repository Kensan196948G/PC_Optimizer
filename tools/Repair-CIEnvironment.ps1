# ==============================================================
# Repair-CIEnvironment.ps1
# CI 自動修復スクリプト
#
# 概要:
#   テスト実行前にリポジトリ内のよくある問題を自動検出・修復する。
#   修復内容をログに記録し、修復後のテスト再実行をサポートする。
#
# 戻り値:
#   0 = すべて正常（修復不要 or 修復成功）
#   1 = 修復できない問題あり（手動対応が必要）
#
# 使い方:
#   pwsh -NoProfile -ExecutionPolicy Bypass -File tools/Repair-CIEnvironment.ps1
#   pwsh -NoProfile -ExecutionPolicy Bypass -File tools/Repair-CIEnvironment.ps1 -RepoRoot "D:\PC_Optimizer"
# ==============================================================

param(
    [string]$RepoRoot = (Split-Path $PSScriptRoot -Parent),
    [switch]$DryRun,
    [string]$StatusFile = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── カラー出力ヘルパー ──────────────────────────────────────────
function Write-Ok([string]$msg)    { Write-Host "  [OK]     $msg" -ForegroundColor Green }
function Write-Fixed([string]$msg) { Write-Host "  [FIXED]  $msg" -ForegroundColor Cyan }
function Write-Warn([string]$msg)  { Write-Host "  [WARN]   $msg" -ForegroundColor Yellow }
function Write-Err([string]$msg)   { Write-Host "  [ERROR]  $msg" -ForegroundColor Red }
function Write-Info([string]$msg)  { Write-Host "  [INFO]   $msg" -ForegroundColor Gray }

# ── 修復カウンター ───────────────────────────────────────────────
$script:FixedCount    = 0
$script:ErrorCount    = 0
$script:CheckedCount  = 0

function Register-Fix([string]$desc) {
    $script:FixedCount++
    Write-Fixed $desc
}
function Register-Error([string]$desc) {
    $script:ErrorCount++
    Write-Err $desc
}
function Register-Check([string]$desc) {
    $script:CheckedCount++
    Write-Ok $desc
}

# ==============================================================
# CHECK 1: Run_PC_Optimizer.bat — CRLF 行末コード
# ==============================================================
Write-Host ""
Write-Host "== [CHECK 1] Run_PC_Optimizer.bat — CRLF 行末コード ==" -ForegroundColor White

$batPath = Join-Path $RepoRoot "Run_PC_Optimizer.bat"
if (-not (Test-Path $batPath)) {
    Register-Error "Run_PC_Optimizer.bat が見つかりません: $batPath"
} else {
    [byte[]]$batBytes = [System.IO.File]::ReadAllBytes($batPath)
    $lfCount   = 0
    $crlfCount = 0
    for ($i = 0; $i -lt $batBytes.Length; $i++) {
        if ($batBytes[$i] -eq 0x0A) {
            if ($i -gt 0 -and $batBytes[$i - 1] -eq 0x0D) {
                $crlfCount++
            } else {
                $lfCount++
            }
        }
    }

    if ($lfCount -gt 0 -and $crlfCount -eq 0) {
        Write-Info "LF のみ行末を検出（$lfCount 行）→ CRLF に変換します"
        if (-not $DryRun) {
            # LF → CRLF 変換（バイト列操作でエンコーディングを保持）
            $newBytes = [System.Collections.Generic.List[byte]]::new($batBytes.Length + $lfCount)
            for ($i = 0; $i -lt $batBytes.Length; $i++) {
                if ($batBytes[$i] -eq 0x0A -and ($i -eq 0 -or $batBytes[$i - 1] -ne 0x0D)) {
                    $newBytes.Add(0x0D)
                }
                $newBytes.Add($batBytes[$i])
            }
            [System.IO.File]::WriteAllBytes($batPath, $newBytes.ToArray())
        }
        Register-Fix "Run_PC_Optimizer.bat: LF → CRLF 変換（$lfCount 行）"
    } elseif ($lfCount -gt 0) {
        Write-Info "混在行末を検出（CRLF: $crlfCount, LF のみ: $lfCount）→ 全行 CRLF に統一します"
        if (-not $DryRun) {
            $sjis = [System.Text.Encoding]::GetEncoding(932)
            $text = $sjis.GetString($batBytes)
            $text = $text -replace "`r`n", "`n"  # まず CRLF → LF に正規化
            $text = $text -replace "`n", "`r`n"  # LF → CRLF
            [System.IO.File]::WriteAllBytes($batPath, $sjis.GetBytes($text))
        }
        Register-Fix "Run_PC_Optimizer.bat: 混在行末 → 全行 CRLF に統一"
    } else {
        Register-Check "Run_PC_Optimizer.bat: CRLF 行末 OK（$crlfCount 行）"
    }
}

# ==============================================================
# CHECK 2: PC_Optimizer.ps1 — UTF-8 BOM 付き
# ==============================================================
Write-Host ""
Write-Host "== [CHECK 2] PC_Optimizer.ps1 — UTF-8 BOM 付き ==" -ForegroundColor White

$ps1Path = Join-Path $RepoRoot "PC_Optimizer.ps1"
if (-not (Test-Path $ps1Path)) {
    Register-Error "PC_Optimizer.ps1 が見つかりません: $ps1Path"
} else {
    [byte[]]$ps1Bytes = [System.IO.File]::ReadAllBytes($ps1Path)
    $hasBom = ($ps1Bytes.Length -ge 3 -and
               $ps1Bytes[0] -eq 0xEF -and
               $ps1Bytes[1] -eq 0xBB -and
               $ps1Bytes[2] -eq 0xBF)

    if (-not $hasBom) {
        Write-Info "UTF-8 BOM なしを検出 → BOM を付加します"
        if (-not $DryRun) {
            $bom = [byte[]]@(0xEF, 0xBB, 0xBF)
            $newBytes = $bom + $ps1Bytes
            [System.IO.File]::WriteAllBytes($ps1Path, $newBytes)
        }
        Register-Fix "PC_Optimizer.ps1: UTF-8 BOM を付加しました"
    } else {
        Register-Check "PC_Optimizer.ps1: UTF-8 BOM あり OK"
    }
}

# ==============================================================
# CHECK 3: PC_Optimizer.ps1 — PowerShell 構文チェック
# ==============================================================
Write-Host ""
Write-Host "== [CHECK 3] PC_Optimizer.ps1 — PowerShell 構文チェック ==" -ForegroundColor White

if (Test-Path $ps1Path) {
    try {
        $tokens = $null
        $errors = $null
        $null = [System.Management.Automation.Language.Parser]::ParseFile(
            $ps1Path,
            [ref]$tokens,
            [ref]$errors
        )
        if ($errors.Count -eq 0) {
            Register-Check "PC_Optimizer.ps1: 構文エラーなし OK"
        } else {
            $errSample = ($errors | Select-Object -First 3 | ForEach-Object {
                "L$($_.Extent.StartLineNumber): $($_.Message)"
            }) -join "; "
            Register-Error "PC_Optimizer.ps1: 構文エラー $($errors.Count) 件 — $errSample"
        }
    } catch {
        Register-Error "PC_Optimizer.ps1: 構文チェック中に例外 — $_"
    }
}

# ==============================================================
# CHECK 4: tests/ 内のスクリプトファイルの存在確認
# ==============================================================
Write-Host ""
Write-Host "== [CHECK 4] テストファイルの存在確認 ==" -ForegroundColor White

$requiredTestFiles = @(
    "tests\Test_PCOptimizer.ps1",
    "tests\Test_CLIModes.ps1",
    "tests\Test_ReportSchema.ps1",
    "tests\Test_WhatIfExport.ps1",
    "Test_Encoding.ps1"
)
foreach ($relPath in $requiredTestFiles) {
    $fullPath = Join-Path $RepoRoot $relPath
    if (Test-Path $fullPath) {
        Register-Check "$($relPath): 存在 OK"
    } else {
        Register-Error "$($relPath): ファイルが見つかりません"
    }
}

# ==============================================================
# CHECK 5: modules/ 内のモジュールファイルの存在確認
# ==============================================================
Write-Host ""
Write-Host "== [CHECK 5] モジュールファイルの存在確認 ==" -ForegroundColor White

$requiredModules = @(
    "modules\Common.psm1",
    "modules\Cleanup.psm1",
    "modules\Update.psm1",
    "modules\Report.psm1"
)
foreach ($relPath in $requiredModules) {
    $fullPath = Join-Path $RepoRoot $relPath
    if (Test-Path $fullPath) {
        Register-Check "$($relPath): 存在 OK"
    } else {
        Register-Error "$($relPath): ファイルが見つかりません"
    }
}

# ==============================================================
# CHECK 6: config/ 内の config.json の存在確認
# ==============================================================
Write-Host ""
Write-Host "== [CHECK 6] config.json の存在確認 ==" -ForegroundColor White

$configPath = Join-Path $RepoRoot "config\config.json"
if (Test-Path $configPath) {
    Register-Check "config\config.json: 存在 OK"
    # JSON の妥当性チェック
    try {
        $null = Get-Content -Path $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
        Register-Check "config\config.json: JSON 形式 OK"
    } catch {
        Register-Error "config\config.json: JSON パースエラー — $_"
    }
} else {
    Register-Error "config\config.json: ファイルが見つかりません"
}

# ==============================================================
# サマリー
# ==============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor White
Write-Host " 自動修復 サマリー" -ForegroundColor White
Write-Host "============================================================" -ForegroundColor White
Write-Host "  チェック済み : $script:CheckedCount 項目"
Write-Host "  自動修復     : $script:FixedCount 項目" -ForegroundColor $(if ($script:FixedCount -gt 0) { 'Cyan' } else { 'Gray' })
Write-Host "  未解決エラー : $script:ErrorCount 項目" -ForegroundColor $(if ($script:ErrorCount -gt 0) { 'Red' } else { 'Gray' })
if ($DryRun) {
    Write-Host "  ※ DryRun モード — 実際のファイル変更は行いませんでした" -ForegroundColor Yellow
}
Write-Host "============================================================" -ForegroundColor White
Write-Host ""

# StatusFile が指定されていれば、カウントを JSON で書き出す（CI パース用）
if ($StatusFile -ne "") {
    $statusData = @{
        checkedCount = $script:CheckedCount
        fixedCount   = $script:FixedCount
        errorCount   = $script:ErrorCount
    } | ConvertTo-Json -Compress
    [System.IO.File]::WriteAllText($StatusFile, $statusData, [System.Text.Encoding]::UTF8)
}

if ($script:ErrorCount -gt 0) {
    Write-Host "未解決の問題があります。手動での対応が必要です。" -ForegroundColor Red
    exit 1
}

if ($script:FixedCount -gt 0) {
    Write-Host "自動修復を $script:FixedCount 件実施しました。" -ForegroundColor Cyan
}

exit 0
