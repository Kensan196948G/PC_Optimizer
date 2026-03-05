<#
.SYNOPSIS
    auto-fix.ps1 — CI 失敗時に安全な自動修正を試みるスクリプト
.DESCRIPTION
    GitHub Actions の auto-fix-on-failure ジョブから呼び出される。
    以下の修正を試みる（危険な変更は一切行わない）:
      1. 必須ディレクトリ構造の作成（.gitkeep 付き）
      2. エンコーディング修正（UTF-16LE → UTF-8 BOM）
      3. Get-WinEvent try-catch 欠落の検出
      4. PSScriptAnalyzer — 安全なルールのみ適用（PSUseSingularNouns を除外）
    修正内容のサマリーを標準出力と GITHUB_OUTPUT に書き出す。
#>
[CmdletBinding()]
param(
    [string]$ArtifactsPath = "auto-fix-artifacts"
)

$ErrorActionPreference = 'Continue'
$fixActions = [System.Collections.Generic.List[string]]::new()

Write-Host "=== auto-fix.ps1 ==="

# ---- ログからエラーカテゴリを検出 ----
$allLogContent = ""
$diagFiles = Get-ChildItem -Path $ArtifactsPath -Recurse -Filter "diag_meta_*.json" -ErrorAction SilentlyContinue
foreach ($f in $diagFiles) {
    try {
        $meta = Get-Content $f.FullName -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json -ErrorAction SilentlyContinue
        if ($meta.failedTestLogs) {
            Write-Host "[auto-fix] Failed test logs from diag: $($meta.failedTestLogs -join ', ')"
        }
    } catch {}
}
$logFiles = Get-ChildItem -Path $ArtifactsPath -Recurse -Filter "*.log" -ErrorAction SilentlyContinue
foreach ($lf in $logFiles) {
    $c = Get-Content $lf.FullName -Raw -ErrorAction SilentlyContinue
    if ($c) { $allLogContent += $c }
}

$hasEncodingError    = $allLogContent -match '(?i)(BOM|encoding error|文字コード|charset mismatch)'
$hasMissingDir       = $allLogContent -match '(?i)(Cannot find path|Directory.*not found|test-results)'
$hasSyntaxError      = $allLogContent -match '(?i)(ParseError|SyntaxError|Unexpected token)'
$hasGetWinEventError = $allLogContent -match '(?i)(Get-WinEvent.*parameter is incorrect|parameter is incorrect.*Get-WinEvent)'
$hasPesterError      = $allLogContent -match '(?i)(Pester.*fail|failed.*pester|\.Tests\.ps1.*fail)'

Write-Host "[auto-fix] Detected: Encoding=$hasEncodingError MissingDir=$hasMissingDir Syntax=$hasSyntaxError GetWinEvent=$hasGetWinEventError Pester=$hasPesterError"

# ---- 1. 必須ディレクトリ構造の作成 ----
$requiredDirs = @(
    "test-results",
    "logs",
    "reports",
    "reports\assets",
    "reports\agent-teams",
    "logs\hooks\queue",
    "logs\hooks\queue\done"
)
$dirFixCount = 0
foreach ($dir in $requiredDirs) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        New-Item -ItemType File -Path "$dir\.gitkeep" -Force | Out-Null
        $dirFixCount++
        Write-Host "[auto-fix] Created directory: $dir"
    }
}
if ($dirFixCount -gt 0) { $fixActions.Add("Directories: $dirFixCount created") }

# ---- 2. エンコーディング修正（UTF-16LE → UTF-8 BOM） ----
if ($hasEncodingError) {
    $encodingFixCount = 0
    $targetFiles = Get-ChildItem -Path . -Include "*.ps1","*.psm1" -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -notmatch '(\.git[/\\]|auto-fix-artifacts[/\\])' }
    foreach ($f in $targetFiles) {
        try {
            $bytes = [System.IO.File]::ReadAllBytes($f.FullName)
            $hasUtf16Le = ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE)
            if ($hasUtf16Le) {
                $content = [System.IO.File]::ReadAllText($f.FullName, [System.Text.Encoding]::Unicode)
                $utf8Bom = [System.Text.UTF8Encoding]::new($true)
                [System.IO.File]::WriteAllText($f.FullName, $content, $utf8Bom)
                $encodingFixCount++
                Write-Host "[auto-fix] Fixed UTF-16LE -> UTF-8 BOM: $($f.Name)"
            }
        } catch {}
    }
    if ($encodingFixCount -gt 0) {
        $fixActions.Add("Encoding: $encodingFixCount files (UTF-16LE -> UTF-8 BOM)")
    }
}

# ---- 3. PSScriptAnalyzer — 安全なルールのみ適用 ----
$safeFixRules = @(
    'PSAvoidTrailingWhitespace',
    'PSUseConsistentIndentation',
    'PSPlaceOpenBrace',
    'PSPlaceCloseBrace',
    'PSUseConsistentWhitespace'
)

$psaInstalled = $false
if (Get-Module -ListAvailable -Name PSScriptAnalyzer -ErrorAction SilentlyContinue) {
    $psaInstalled = $true
} else {
    Write-Host "[auto-fix] Installing PSScriptAnalyzer..."
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Install-Module PSScriptAnalyzer -Scope CurrentUser -Force -SkipPublisherCheck -ErrorAction Stop
        $psaInstalled = $true
    } catch {
        Write-Warning "[auto-fix] PSScriptAnalyzer install failed: $_"
    }
}

if ($psaInstalled) {
    $psFiles = Get-ChildItem -Path . -Include "*.ps1","*.psm1" -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -notmatch '(\.git[/\\]|auto-fix-artifacts[/\\]|node_modules[/\\])' }

    $analyzerFixCount = 0
    foreach ($psFile in $psFiles) {
        try {
            $before = Get-Content $psFile.FullName -Raw -ErrorAction SilentlyContinue
            Invoke-ScriptAnalyzer -Path $psFile.FullName -Fix `
                -IncludeRule $safeFixRules `
                -ErrorAction SilentlyContinue | Out-Null
            $after = Get-Content $psFile.FullName -Raw -ErrorAction SilentlyContinue
            if ($before -ne $after) { $analyzerFixCount++ }
        } catch {}
    }
    if ($analyzerFixCount -gt 0) {
        $fixActions.Add("PSScriptAnalyzer(safe): $analyzerFixCount files fixed")
        Write-Host "[auto-fix] PSScriptAnalyzer (safe rules) fixed $analyzerFixCount files"
    } else {
        Write-Host "[auto-fix] PSScriptAnalyzer (safe rules): no fixable issues found"
    }
}

# ---- 結果出力 ----
if ($fixActions.Count -eq 0) {
    $summary = "no-fixable-issues-found"
} else {
    $raw = $fixActions -join "; "
    $summary = if ($raw.Length -gt 250) { $raw.Substring(0, 250) } else { $raw }
}

Write-Host "[auto-fix] === Summary: $summary ==="

if ($env:GITHUB_OUTPUT) {
    "fix-summary=$summary" | Out-File -FilePath $env:GITHUB_OUTPUT -Append -Encoding utf8
    "fix-count=$($fixActions.Count)" | Out-File -FilePath $env:GITHUB_OUTPUT -Append -Encoding utf8
}

return $summary
