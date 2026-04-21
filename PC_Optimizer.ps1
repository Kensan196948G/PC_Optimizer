# 高性能 PC 最適化ツール - PowerShell 3.0 / 5.1 / 7.x 対応
# 追加機能:
#  - 配信最適化キャッシュの削除
#  - Windows Update キャッシュの削除
#  - エラーレポート・ログ・不要キャッシュの削除
#  - OneDrive / Teams / Office キャッシュの削除
#  - SSD ヘルスチェック
#  - ブラウザキャッシュ削除（Chrome / Edge / Firefox / Brave / Opera / Vivaldi）
#  - サムネイルキャッシュ削除
#  - Microsoft Store キャッシュクリア
#  - Windows イベントログのクリア
#  - システムファイル整合性チェック・修復（SFC）
#  - Windows コンポーネントストア診断（DISM）
#  - 電源プランの最適化
#  - 実行前後のディスク空き容量比較
# エンコーディング: UTF-8 BOM 付き (PS5.1) / UTF-8 BOM なし (PS7+) -- docs/文字コード規約.md 参照

param(
    [switch]$NonInteractive,
    [switch]$WhatIf,
    [switch]$NoRebootPrompt,
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config\config.json"),
    [ValidateSet("repair","diagnose")]
    [string]$Mode = "repair",
    [string]$Tasks = "all",
    [string]$FailureMode = "continue",
    [string]$ExportDeletedPaths = "",
    [string]$ExportDeletedPathsPath = "",
    [bool]$EnableAIDiagnosis = $true,
    [switch]$UseAnthropicAI,
    [string]$AnthropicApiKey = "",
    [string]$AnthropicModel = "claude-sonnet-4-6",
    [string]$AssetAggregateInputDir = "",
    [string]$AssetAggregateOutputDir = "",
    [string]$RemoteComputerListPath = "",
    [string]$RemoteDiagnosticsOutputDir = "",
    [ValidateSet("","create","update","delete")]
    [string]$ScheduleTaskAction = "",
    [string]$ScheduleTaskName = "PC_Optimizer_Monthly",
    [int]$ScheduleDayOfMonth = 1,
    [string]$ScheduleTime = "03:00",
    [ValidateSet("classic","agent-teams")]
    [string]$ExecutionProfile = "classic",
    [switch]$UseLocalChartJs,
    [string]$ChartJsLocalRelativePath = "assets/chart.umd.min.js",
    [switch]$ExportPowerBIJson,
    [switch]$EmitUiEvents,
    [string]$UiEventPrefix = "##PCOPT_UI##"
)

# ログパスの構築
$now     = Get-Date -Format "yyyyMMddHHmm"
$logsDir = Join-Path -Path $PSScriptRoot -ChildPath "logs"
$reportsDir = Join-Path -Path $PSScriptRoot -ChildPath "reports"
if (-not (Test-Path $logsDir)) {
    New-Item -ItemType Directory -Path $logsDir -Force | Out-Null
}
if (-not (Test-Path $reportsDir)) {
    New-Item -ItemType Directory -Path $reportsDir -Force | Out-Null
}
$logPath      = Join-Path -Path $logsDir -ChildPath "PC_Optimizer_Log_${now}.txt"
$errorLogPath = Join-Path -Path $logsDir -ChildPath "PC_Optimizer_Error_${now}.txt"

# ── バージョン検出（ログ関数定義前に記述すること）──
$psver     = $PSVersionTable.PSVersion.Major
$isPS7Plus = ($psver -ge 7)

# ANSI 太文字（PS7+ の対応ターミナルのみ有効、PS5.1 では空文字列にフォールバック）
# ESC[1m  = 太文字 ON
# ESC[22m = 太文字のみ OFF（ForegroundColor の色は保持）← 行途中の切り替えに使用
# ESC[0m  = 全属性リセット ← 行末に置く場合に使用
$B  = if ($isPS7Plus) { "$([char]27)[1m"  } else { "" }
$RB = if ($isPS7Plus) { "$([char]27)[22m" } else { "" }
$R  = if ($isPS7Plus) { "$([char]27)[0m"  } else { "" }

# エンコーディング規約: 暗黙のデフォルト禁止 — 常に明示指定すること
# PS5.1 : 'UTF8'      → UTF-8 BOM 付き  (メモ帳 / Excel で文字化けしない)
# PS7.x : 'utf8NoBOM' → UTF-8 BOM なし  (IETF 標準)
$logEncoding = if ($isPS7Plus) { 'utf8NoBOM' } else { 'UTF8' }
$script:IsNonInteractive  = [bool]$NonInteractive
$script:IsWhatIfMode      = [bool]$WhatIf
$script:IsNoRebootPrompt  = [bool]$NoRebootPrompt
$script:Config            = $null
$script:TaskCounter       = 0
$script:SelectedTaskSet   = $null
$script:FailureMode       = $FailureMode.ToLowerInvariant()
$script:RunMode           = $Mode.ToLowerInvariant()
$script:ExportDeletedFmt  = $ExportDeletedPaths.ToLowerInvariant()
$script:ExportDeletedPath = $ExportDeletedPathsPath
$script:HadTaskFailure    = $false
$script:FatalStopRequested= $false
$script:ExitCode          = 0
$script:RunId             = ([guid]::NewGuid().ToString("N"))
$script:DeletedPathSet    = New-Object 'System.Collections.Generic.HashSet[string]'
$script:DeletedPathList   = New-Object 'System.Collections.Generic.List[string]'
$script:LoadedModules     = New-Object 'System.Collections.Generic.List[string]'
$script:ModuleSnapshot    = $null
$script:HealthScore       = $null
$script:GuardrailState    = $null
$script:AIDiagnosis       = $null
$script:UpdateErrorClassification = @()
$script:M365Connectivity  = @()
$script:EventAnomaly      = $null
$script:BootTrend         = $null
$script:HookHistory       = @()
$script:AgentTeamsResult  = $null
$script:AgentTeamsSummary = $null
$script:McpResults        = @()
$script:ExecutionProfile  = $ExecutionProfile
$script:EmitUiEvents      = [bool]$EmitUiEvents
$script:UiEventPrefix     = $UiEventPrefix

$script:ExitCodes = @{
    Success       = 0
    Partial       = 1
    Fatal         = 2
    InvalidArgs   = 3
    Permission    = 4
}

# HTMLレポート用タスク結果収集
$script:taskResults     = @()
$script:scriptStartTime = Get-Date
$script:sysInfo         = $null    # CIM 取得失敗時のフォールバック用

if ($psver -lt 3) {
    Write-Host "警告: PowerShell 3.0 以上を推奨します。" -ForegroundColor Yellow
    Write-Host "現在のバージョン: $($PSVersionTable.PSVersion.ToString())" -ForegroundColor Yellow
    Write-Host "一部の機能が正常に動作しない場合があります。" -ForegroundColor Yellow
}

function Write-Log($msg) {
    # FileShare.ReadWrite でロック競合を回避（Defender 等の同時スキャンに対応）
    try {
        $enc = if ($isPS7Plus) { [System.Text.UTF8Encoding]::new($false) } else { [System.Text.UTF8Encoding]::new($true) }
        $fs  = [System.IO.FileStream]::new($logPath, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
        $sw  = [System.IO.StreamWriter]::new($fs, $enc)
        $sw.WriteLine($msg)
        $sw.Dispose()
    } catch {
        Write-Host "[ログ書き込み失敗（スキップ）] $msg" -ForegroundColor DarkGray
    }
}
function Write-ErrorLog($msg) {
    try {
        $enc = if ($isPS7Plus) { [System.Text.UTF8Encoding]::new($false) } else { [System.Text.UTF8Encoding]::new($true) }
        $fs  = [System.IO.FileStream]::new($errorLogPath, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
        $sw  = [System.IO.StreamWriter]::new($fs, $enc)
        $sw.WriteLine("[ERROR] $(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')")
        $sw.WriteLine($msg)
        $sw.Dispose()
    } catch {
        Write-Host "[エラーログ書き込み失敗（スキップ）] $msg" -ForegroundColor DarkGray
    }
}

function Show($msg, $color = "White") {
    Write-Host $msg -ForegroundColor $color
}

function Write-UiEvent {
    param(
        [Parameter(Mandatory)][string]$Type,
        [hashtable]$Data = @{}
    )

    if (-not $script:EmitUiEvents) { return }

    $payload = [ordered]@{
        type      = $Type
        runId     = $script:RunId
        timestamp = (Get-Date).ToString("o")
    }
    foreach ($key in $Data.Keys) {
        $payload[$key] = $Data[$key]
    }

    try {
        $json = [PSCustomObject]$payload | ConvertTo-Json -Compress -Depth 8
        [Console]::Out.WriteLine(("{0}{1}" -f $script:UiEventPrefix, $json))
    } catch {
        Write-Log "[ui-event] シリアライズ失敗: type=$Type error=$_"
    }
}

function Convert-HealthStatusToJapanese {
    param([string]$Status)
    $safeStatus = if ($null -eq $Status) { "" } else { "$Status" }
    switch ($safeStatus.ToLowerInvariant()) {
        "excellent" { "優秀" }
        "good" { "良好" }
        "warning" { "注意" }
        "critical" { "重大" }
        default { if ([string]::IsNullOrWhiteSpace($Status)) { "不明" } else { $Status } }
    }
}

function Get-SystemDriveFreeGB {
    [CmdletBinding()]
    param([string]$DriveLetter = "C:")

    $disk = Get-CimInstance Win32_LogicalDisk -Filter ("DeviceID='{0}'" -f $DriveLetter) -ErrorAction SilentlyContinue
    if ($disk -and $null -ne $disk.FreeSpace) {
        return [math]::Round(($disk.FreeSpace / 1GB), 2)
    }

    $psDrive = Get-PSDrive -Name ($DriveLetter.TrimEnd(':')) -ErrorAction SilentlyContinue
    if ($psDrive -and $null -ne $psDrive.Free) {
        return [math]::Round(($psDrive.Free / 1GB), 2)
    }

    return $null
}

function Ensure-ReportChartAsset {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ReportsDir,
        [Parameter(Mandatory)][string]$ScriptRoot,
        [Parameter(Mandatory)][string]$RelativePath
    )

    $targetPath = Join-Path $ReportsDir $RelativePath
    $targetDir = Split-Path -Parent $targetPath
    if (-not (Test-Path $targetDir)) {
        New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
    }

    $sourcePath = Join-Path $ScriptRoot $RelativePath
    if (Test-Path $sourcePath) {
        Copy-Item -Path $sourcePath -Destination $targetPath -Force
        return $true
    }
    return $false
}

function Import-DotEnvToProcess {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$OverwriteExisting
    )

    if (-not (Test-Path -LiteralPath $Path)) { return 0 }
    $count = 0
    $lines = Get-Content -LiteralPath $Path -Encoding UTF8 -ErrorAction SilentlyContinue
    foreach ($raw in @($lines)) {
        if ($null -eq $raw) { continue }
        $line = "$raw".Trim()
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        if ($line.StartsWith("#") -or $line.StartsWith(";")) { continue }

        $eq = $line.IndexOf("=")
        if ($eq -lt 1) { continue }

        $key = $line.Substring(0, $eq).Trim()
        $val = $line.Substring($eq + 1).Trim()
        if ($key.StartsWith("export ", [System.StringComparison]::OrdinalIgnoreCase)) { $key = $key.Substring(7).Trim() }
        if ($key.StartsWith("set ", [System.StringComparison]::OrdinalIgnoreCase)) { $key = $key.Substring(4).Trim() }
        if ([string]::IsNullOrWhiteSpace($key)) { continue }

        if ($val.Length -ge 2 -and $val.StartsWith('"') -and $val.EndsWith('"')) {
            $val = $val.Substring(1, $val.Length - 2)
        }

        $current = [Environment]::GetEnvironmentVariable($key, "Process")
        if (-not $OverwriteExisting -and -not [string]::IsNullOrWhiteSpace($current)) { continue }
        [Environment]::SetEnvironmentVariable($key, $val, "Process")
        $count++
    }
    return $count
}

# .env を実行時にも読込（bat経由でない実行や昇格時の環境引継ぎ差異を吸収）
try {
    $dotenvPath = Join-Path $PSScriptRoot ".env"
    $imported = Import-DotEnvToProcess -Path $dotenvPath
    if ($imported -gt 0) {
        Write-Log "[env] .env 読込: $imported 件"
    }
    if (-not $PSBoundParameters.ContainsKey("AnthropicModel")) {
        $envModel = [Environment]::GetEnvironmentVariable("ANTHROPIC_MODEL", "Process")
        if (-not [string]::IsNullOrWhiteSpace($envModel)) {
            $AnthropicModel = $envModel.Trim()
        }
    }
} catch {
    Write-ErrorLog ".env 読込に失敗しました: $_"
}

# v4.0 foundation: 共通モジュールと設定を接続
try {
    $commonModulePath = Join-Path $PSScriptRoot "modules\Common.psm1"
    if (Test-Path $commonModulePath) {
        Import-Module $commonModulePath -Force -ErrorAction Stop
        if (-not ($script:LoadedModules -contains "modules\Common.psm1")) {
            [void]$script:LoadedModules.Add("modules\Common.psm1")
        }
    }
    $moduleCandidates = @(
        "modules\TaskCatalog.psm1",
        "modules\Cleanup.psm1",
        "modules\Diagnostics.psm1",
        "modules\Performance.psm1",
        "modules\Network.psm1",
        "modules\Security.psm1",
        "modules\Update.psm1",
        "modules\Advanced.psm1",
        "modules\Orchestration.psm1",
        "modules\Report.psm1"
    )
    foreach ($moduleRelPath in $moduleCandidates) {
        $modulePath = Join-Path $PSScriptRoot $moduleRelPath
        if (-not (Test-Path $modulePath)) {
            Write-Log "[module] 未検出: $moduleRelPath"
            continue
        }
        Import-Module $modulePath -Force -ErrorAction Stop
        [void]$script:LoadedModules.Add($moduleRelPath)
    }
    if (Get-Command Get-OptimizerConfig -ErrorAction SilentlyContinue) {
        $script:Config = Get-OptimizerConfig -Path $ConfigPath
        Write-Log "[config] 読込成功: $ConfigPath"
        if ($script:Config -and $script:Config.PSObject.Properties["useLocalChartJs"] -and -not $PSBoundParameters.ContainsKey("UseLocalChartJs")) {
            $UseLocalChartJs = [bool]$script:Config.useLocalChartJs
        }
        if ($script:Config -and $script:Config.PSObject.Properties["chartJsLocalRelativePath"] -and $script:Config.chartJsLocalRelativePath -and -not $PSBoundParameters.ContainsKey("ChartJsLocalRelativePath")) {
            $ChartJsLocalRelativePath = "$($script:Config.chartJsLocalRelativePath)"
        }
        if ($script:Config -and $script:Config.PSObject.Properties["executionProfile"] -and $script:Config.executionProfile -and -not $PSBoundParameters.ContainsKey("ExecutionProfile")) {
            $script:ExecutionProfile = "$($script:Config.executionProfile)"
        }
    }
} catch {
    Write-ErrorLog "共通モジュールまたは config 読込に失敗しました: $_"
}
if ($script:LoadedModules.Count -gt 0) {
    Write-Log ("[module] 読込済み: {0}" -f ($script:LoadedModules -join ", "))
}

function Get-UserChoice {
    param(
        [string]$Prompt,
        [ValidateSet('Y', 'N')]
        [string]$Default = 'N'
    )

    if ($script:IsNonInteractive) {
        Write-Log "[NonInteractive] 自動応答: $Default / Prompt=$Prompt"
        return $Default
    }

    $choice = Read-Host $Prompt
    if ([string]::IsNullOrWhiteSpace($choice)) { return $Default }
    return $choice
}

function Test-ChoiceYes {
    param([string]$Choice)
    return ($Choice -match '^[Yy]$')
}

function Test-ReadOnlyTask {
    param([string]$TaskName)
    return ($TaskName -match 'スタートアップ・サービスレポート')
}

function Test-IsAdministrator {
    try {
        $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        Write-ErrorLog "管理者権限判定に失敗しました: $_"
        return $false
    }
}

function Test-ReservedPathSegment {
    param([string]$Segment)
    if ([string]::IsNullOrWhiteSpace($Segment)) { return $false }
    $trimmed = $Segment.TrimEnd(':')
    return ($trimmed -match '^(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(\..*)?$')
}

function Initialize-ExecutionOption {
    if ($script:RunMode -notin @('repair', 'diagnose')) {
        Show "不正な -Mode です: $Mode" Red
        $script:ExitCode = $script:ExitCodes.InvalidArgs
        exit $script:ExitCode
    }
    if ($script:FailureMode -notin @('continue', 'fail-fast')) {
        Show "不正な -FailureMode です: $FailureMode" Red
        $script:ExitCode = $script:ExitCodes.InvalidArgs
        exit $script:ExitCode
    }
    if ($ScheduleDayOfMonth -lt 1 -or $ScheduleDayOfMonth -gt 31) {
        Show "不正な -ScheduleDayOfMonth です: $ScheduleDayOfMonth" Red
        $script:ExitCode = $script:ExitCodes.InvalidArgs
        exit $script:ExitCode
    }
    if ($ScheduleTime -notmatch '^\d{2}:\d{2}$') {
        Show "不正な -ScheduleTime です: $ScheduleTime (HH:mm)" Red
        $script:ExitCode = $script:ExitCodes.InvalidArgs
        exit $script:ExitCode
    }
    if ($script:ExecutionProfile -notin @('classic','agent-teams')) {
        Show "不正な -ExecutionProfile です: $($script:ExecutionProfile)" Red
        $script:ExitCode = $script:ExitCodes.InvalidArgs
        exit $script:ExitCode
    }
    if ($script:ExportDeletedFmt -and $script:ExportDeletedFmt -notin @('csv', 'json')) {
        Show "不正な -ExportDeletedPaths です: $ExportDeletedPaths" Red
        $script:ExitCode = $script:ExitCodes.InvalidArgs
        exit $script:ExitCode
    }
    if ($script:ExportDeletedPath) {
        $rawPath = $script:ExportDeletedPath.Trim()
        if ([string]::IsNullOrWhiteSpace($rawPath)) {
            Show "不正な -ExportDeletedPathsPath です（空文字）: $ExportDeletedPathsPath" Red
            $script:ExitCode = $script:ExitCodes.InvalidArgs
            exit $script:ExitCode
        }
        $invalidChars = [IO.Path]::GetInvalidPathChars()
        if ($rawPath.IndexOfAny($invalidChars) -ge 0) {
            Show "不正な -ExportDeletedPathsPath です: $ExportDeletedPathsPath" Red
            $script:ExitCode = $script:ExitCodes.InvalidArgs
            exit $script:ExitCode
        }
        $resolvedPath = $rawPath
        if (-not [IO.Path]::IsPathRooted($resolvedPath)) {
            $resolvedPath = Join-Path $PSScriptRoot $resolvedPath
            Write-Log "[WhatIfExport] 相対パスを絶対パスへ解決: $rawPath -> $resolvedPath"
        }
        try {
            $resolvedPath = [IO.Path]::GetFullPath($resolvedPath)
        } catch {
            Show "不正な -ExportDeletedPathsPath です（解決失敗）: $ExportDeletedPathsPath" Red
            $script:ExitCode = $script:ExitCodes.InvalidArgs
            exit $script:ExitCode
        }
        $segments = @($resolvedPath -split '[\\/]')
        foreach ($seg in $segments) {
            if (Test-ReservedPathSegment -Segment $seg) {
                Show "不正な -ExportDeletedPathsPath です（予約語を含む）: $ExportDeletedPathsPath" Red
                $script:ExitCode = $script:ExitCodes.InvalidArgs
                exit $script:ExitCode
            }
        }
        $script:ExportDeletedPath = $resolvedPath
    }

    if ($Tasks -and $Tasks -ne 'all') {
        if ($Tasks -match ',\s*$') {
            Show "-Tasks の末尾にカンマがあります: $Tasks" Red
            $script:ExitCode = $script:ExitCodes.InvalidArgs
            exit $script:ExitCode
        }
        $set = New-Object 'System.Collections.Generic.HashSet[int]'
        $tokens = $Tasks -split ','
        for ($idx = 0; $idx -lt $tokens.Count; $idx++) {
            $token = $tokens[$idx]
            $trimmed = $token.Trim()
            if ([string]::IsNullOrWhiteSpace($trimmed)) {
                Show "-Tasks に空要素があります（位置: $($idx+1)）: $Tasks" Red
                $script:ExitCode = $script:ExitCodes.InvalidArgs
                exit $script:ExitCode
            }
            if ($trimmed -match '^(\d+)-(\d+)$') {
                $startId = [int]$matches[1]
                $endId   = [int]$matches[2]
                if ($startId -gt $endId) {
                    Show "不正な -Tasks 範囲です: $trimmed" Red
                    $script:ExitCode = $script:ExitCodes.InvalidArgs
                    exit $script:ExitCode
                }
                if ($startId -lt 1 -or $endId -gt 20) {
                    Show "-Tasks は 1-20 の範囲で指定してください: $Tasks" Red
                    $script:ExitCode = $script:ExitCodes.InvalidArgs
                    exit $script:ExitCode
                }
                for ($id = $startId; $id -le $endId; $id++) {
                    if (-not $set.Add($id)) {
                        Show "-Tasks に重複指定があります: $id (入力: $trimmed)" Red
                        $script:ExitCode = $script:ExitCodes.InvalidArgs
                        exit $script:ExitCode
                    }
                }
                continue
            }
            if ($trimmed -match '^\d+$') {
                $id = [int]$trimmed
                if ($id -lt 1 -or $id -gt 20) {
                    Show "-Tasks は 1-20 の範囲で指定してください: $Tasks" Red
                    $script:ExitCode = $script:ExitCodes.InvalidArgs
                    exit $script:ExitCode
                }
                if (-not $set.Add($id)) {
                    Show "-Tasks に重複指定があります: $id (入力: $trimmed)" Red
                    $script:ExitCode = $script:ExitCodes.InvalidArgs
                    exit $script:ExitCode
                }
                continue
            }
            Show "不正な -Tasks 指定です: $Tasks" Red
            $script:ExitCode = $script:ExitCodes.InvalidArgs
            exit $script:ExitCode
        }
        $script:SelectedTaskSet = $set
    }
}

function Add-DeletedPathCandidate {
    param([string]$Path)
    if (-not $Path) { return }
    if ($script:DeletedPathSet.Add($Path)) {
        [void]$script:DeletedPathList.Add($Path)
    }
}

function Resolve-ConfiguredPath {
    param([string]$Path)
    if (-not $Path) { return $Path }
    $resolved = $Path
    $matches = [regex]::Matches($Path, '%([^%]+)%')
    foreach ($m in $matches) {
        $name = $m.Groups[1].Value
        $val = [Environment]::GetEnvironmentVariable($name)
        if ($val) {
            $resolved = $resolved.Replace($m.Value, $val)
        }
    }
    return $resolved
}

function Register-TaskPlannedDeletedPath {
    param([int]$TaskId)
    $configured = $null
    if ($script:Config -and $script:Config.whatIfDeletedPathMap) {
        $configured = $script:Config.whatIfDeletedPathMap.PSObject.Properties["$TaskId"]
    }
    if ($configured) {
        foreach ($p in @($configured.Value)) {
            Add-DeletedPathCandidate -Path (Resolve-ConfiguredPath -Path "$p")
        }
        return
    }
    switch ($TaskId) {
        1 {
            Add-DeletedPathCandidate "$env:SystemRoot\Temp"
            Add-DeletedPathCandidate $env:TEMP
            Add-DeletedPathCandidate "$env:USERPROFILE\AppData\Local\Temp"
        }
        2 {
            Add-DeletedPathCandidate "$env:SystemRoot\Prefetch"
            Add-DeletedPathCandidate "$env:SystemRoot\SoftwareDistribution\Download"
            Add-DeletedPathCandidate "$env:SystemRoot\System32\DeliveryOptimization\Cache"
        }
        3 { Add-DeletedPathCandidate "$env:SystemRoot\SoftwareDistribution\DeliveryOptimization\Cache" }
        4 { Add-DeletedPathCandidate "$env:SystemRoot\SoftwareDistribution\Download" }
        5 {
            Add-DeletedPathCandidate "$env:ProgramData\Microsoft\Windows\WER\ReportArchive"
            Add-DeletedPathCandidate "$env:ProgramData\Microsoft\Windows\WER\ReportQueue"
            Add-DeletedPathCandidate "$env:SystemRoot\Logs\CBS"
        }
        6 {
            Add-DeletedPathCandidate "$env:LOCALAPPDATA\Microsoft\OneDrive\logs"
            Add-DeletedPathCandidate "$env:APPDATA\Microsoft\Teams\Cache"
            Add-DeletedPathCandidate "$env:LOCALAPPDATA\Microsoft\Office\16.0\OfficeFileCache"
        }
        7 {
            Add-DeletedPathCandidate "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache"
            Add-DeletedPathCandidate "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache"
            Add-DeletedPathCandidate "$env:APPDATA\Mozilla\Firefox\Profiles\*\cache2\entries"
            Add-DeletedPathCandidate "$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\User Data\Default\Cache"
            Add-DeletedPathCandidate "$env:APPDATA\Opera Software\Opera Stable\Cache"
            Add-DeletedPathCandidate "$env:LOCALAPPDATA\Vivaldi\User Data\Default\Cache"
        }
        8 { Add-DeletedPathCandidate "$env:LOCALAPPDATA\Microsoft\Windows\Explorer\thumbcache_*.db" }
        9 {
            Add-DeletedPathCandidate "$env:LOCALAPPDATA\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe\LocalCache"
            Add-DeletedPathCandidate "$env:LOCALAPPDATA\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe\LocalState\Cache"
        }
        10 { Add-DeletedPathCandidate '$Recycle.Bin' }
    }
}

function Export-DeletedPathCandidate {
    if (-not $script:ExportDeletedFmt) { return $null }
    $outputDir = if ($script:ExportDeletedPath) { $script:ExportDeletedPath } else { $logsDir }
    if (-not (Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }
    $stamp = Get-Date -Format "yyyyMMddHHmmss"
    $sorted = @($script:DeletedPathList | Sort-Object -Unique)
    if ($script:ExportDeletedFmt -eq 'json') {
        $path = Join-Path $outputDir "DeletedPaths_${stamp}.json"
        [PSCustomObject]@{
            runId = $script:RunId
            mode  = if ($script:IsWhatIfMode) { 'whatif' } else { 'execute' }
            tasks = $sorted
        } | ConvertTo-Json -Depth 4 | Set-Content -Path $path -Encoding $logEncoding
        Write-Log "[DeletedPaths] JSON 出力: $path"
        return $path
    }
    $path = Join-Path $outputDir "DeletedPaths_${stamp}.csv"
    $sorted | ForEach-Object { [PSCustomObject]@{ path = $_ } } |
        Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
    Write-Log "[DeletedPaths] CSV 出力: $path"
    return $path
}

function Invoke-StandaloneOperation {
    $executed = $false

    if ($ScheduleTaskAction) {
        $executed = $true
        $scriptPath = Join-Path $PSScriptRoot "PC_Optimizer.ps1"
        if ($ScheduleTaskAction -in @("create", "update")) {
            $result = "Failed"
            if (Get-Command Set-MonthlyMaintenanceTask -ErrorAction SilentlyContinue) {
                try {
                    $scheduledArgs = "-NonInteractive -Mode repair -ExecutionProfile $($script:ExecutionProfile)"
                    $result = Set-MonthlyMaintenanceTask -TaskName $ScheduleTaskName -ScriptPath $scriptPath -ScriptArguments $scheduledArgs -DayOfMonth $ScheduleDayOfMonth -At $ScheduleTime
                } catch {
                    $result = "Failed: $($_.Exception.Message)"
                }
            }
            Show "Task Scheduler 設定結果: $result" Cyan
            Write-Log "[scheduler] action=$ScheduleTaskAction result=$result"
        } elseif ($ScheduleTaskAction -eq "delete") {
            $result = "Failed"
            if (Get-Command Remove-MonthlyMaintenanceTask -ErrorAction SilentlyContinue) {
                try {
                    $result = Remove-MonthlyMaintenanceTask -TaskName $ScheduleTaskName
                } catch {
                    $result = "Failed: $($_.Exception.Message)"
                }
            }
            Show "Task Scheduler 削除結果: $result" Yellow
            Write-Log "[scheduler] action=delete result=$result"
        }
    }

    if ($AssetAggregateInputDir) {
        $executed = $true
        $outDir = if ($AssetAggregateOutputDir) { $AssetAggregateOutputDir } else { Join-Path $reportsDir "asset-aggregate" }
        if (-not [IO.Path]::IsPathRooted($outDir)) { $outDir = Join-Path $PSScriptRoot $outDir }
        $inDir = $AssetAggregateInputDir
        if (-not [IO.Path]::IsPathRooted($inDir)) { $inDir = Join-Path $PSScriptRoot $inDir }
        $agg = $null
        if (Get-Command Invoke-AssetCentralAggregation -ErrorAction SilentlyContinue) {
            try {
                $agg = Invoke-AssetCentralAggregation -InputDir $inDir -OutputDir $outDir
            } catch {
                Write-ErrorLog "[asset] 集約失敗: $_"
            }
        }
        if ($agg) {
            Show ("資産集約完了: {0}件" -f $agg.Count) Green
            Write-Log "[asset] aggregated=$($agg.Count) csv=$($agg.AggregatedCsv)"
        }
    }

    if ($RemoteComputerListPath) {
        $executed = $true
        $listPath = $RemoteComputerListPath
        if (-not [IO.Path]::IsPathRooted($listPath)) { $listPath = Join-Path $PSScriptRoot $listPath }
        $outDir = if ($RemoteDiagnosticsOutputDir) { $RemoteDiagnosticsOutputDir } else { Join-Path $reportsDir "remote" }
        if (-not [IO.Path]::IsPathRooted($outDir)) { $outDir = Join-Path $PSScriptRoot $outDir }
        if (-not (Test-Path $listPath)) {
            Show "RemoteComputerListPath が見つかりません: $RemoteComputerListPath" Red
            $script:ExitCode = $script:ExitCodes.InvalidArgs
            return $true
        }
        $computers = @(Get-Content -Path $listPath -Encoding UTF8 | ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith('#') })
        $sum = $null
        if (Get-Command Invoke-WinRMRemoteDiagnosticsBatch -ErrorAction SilentlyContinue) {
            try {
                $sum = Invoke-WinRMRemoteDiagnosticsBatch -ComputerName $computers -OutputDir $outDir -RetryCount 1 -ThrottleLimit 8
            } catch {
                Write-ErrorLog "[remote] 一括診断失敗: $_"
            }
        }
        if ($sum) {
            Show ("WinRM診断完了: 成功={0}, 失敗={1}" -f @($sum.Succeeded).Count, @($sum.Failed).Count) Cyan
            Write-Log "[remote] success=$(@($sum.Succeeded).Count) ng=$(@($sum.Failed).Count)"
        }
    }
    return $executed
}

Initialize-ExecutionOption
Write-UiEvent -Type "run_start" -Data @{
    mode             = $script:RunMode
    executionProfile = $script:ExecutionProfile
    whatIf           = $script:IsWhatIfMode
    nonInteractive   = $script:IsNonInteractive
}
if (Invoke-StandaloneOperation) {
    if ($script:ExitCode -eq 0) { $script:ExitCode = $script:ExitCodes.Success }
    exit $script:ExitCode
}
if (-not (Test-IsAdministrator) -and -not ($script:IsWhatIfMode -and $script:IsNonInteractive)) {
    Show "管理者権限で実行してください。" Red
    $script:ExitCode = $script:ExitCodes.Permission
    exit $script:ExitCode
}
if (Get-Command Start-RepairGuardrails -ErrorAction SilentlyContinue) {
    $script:GuardrailState = Start-RepairGuardrails -Mode $script:RunMode -RootPath $PSScriptRoot -LogsDir $logsDir -WhatIfMode:$script:IsWhatIfMode
    Write-Log "[guardrail] started mode=$($script:RunMode) restorePoint=$($script:GuardrailState.RestorePointStatus)"
}

# ハードウェア・システム情報の収集（24H2 / 22H2 等のマーケティングバージョンを含む）
try {
    $hostname = $env:COMPUTERNAME
    $username = $env:USERNAME

    # OS 名とマーケティングバージョン（24H2/22H2 など）
    $osinfo = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $osName = if ($osinfo) { $osinfo.Caption } else { $env:OS }
    $osVer  = if ($osinfo) { $osinfo.Version } else { "Unknown" }
    $osMarketingVer = ""
    try {
        $verRegPath     = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
        $displayVersion = (Get-ItemProperty -Path $verRegPath -Name "DisplayVersion" -ErrorAction SilentlyContinue).DisplayVersion
        if (-not $displayVersion) {
            $displayVersion = (Get-ItemProperty -Path $verRegPath -Name "ReleaseId" -ErrorAction SilentlyContinue).ReleaseId
        }
        if ($displayVersion) { $osMarketingVer = " $displayVersion" }
    } catch {
        # ベストエフォートのレジストリ読み取り — 失敗時はマーケティングバージョンを省略
        Write-ErrorLog "OS マーケティングバージョンのレジストリ読み取り失敗（非致命的）: $_"
    }
    $os = "$osName$osMarketingVer ($osVer)"

    # その他の情報（アクセス拒否時は既定値で継続）
    $cpuinfo = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
    $cpu = if ($cpuinfo) { $cpuinfo.Name } else { "Unknown CPU" }
    $cpuCores = if ($cpuinfo) { $cpuinfo.NumberOfCores } else { "?" }
    $cpuThreads = if ($cpuinfo) { $cpuinfo.NumberOfLogicalProcessors } else { "?" }
    $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
    $mem = if ($cs -and $cs.TotalPhysicalMemory) { [Math]::Round(($cs.TotalPhysicalMemory / 1GB), 2) } else { 0 }
    $disk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction SilentlyContinue
    $diskId = if ($disk) { $disk.DeviceID } else { "C:" }
    $free = if ($disk) { [math]::Round(($disk.FreeSpace / 1GB), 1) } else { 0 }
    $total = if ($disk) { [math]::Round(($disk.Size / 1GB), 1) } else { 0 }
    $pwv        = $PSVersionTable.PSVersion.ToString()

    # コンソール出力 — PS7+ は絵文字、PS5.x 以前は ASCII 記号
    if ($isPS7Plus) {
        Show "🖥️  ${B}ホスト:${RB} $hostname" Green
        Show "👤  ${B}ユーザー:${RB} $username" Cyan
        Show "🏷️  ${B}OS:${RB} $os" Yellow
        Show "💻  ${B}CPU:${RB} $cpu  $cpuCores コア / $cpuThreads スレッド" Magenta
        Show "🧠  ${B}メモリ:${RB} ${mem} GB" Blue
        Show "💾  ${B}ディスク ${diskId} 空き:${RB} ${free}GB / ${total}GB" White
    } else {
        Show "* ホスト: $hostname" Green
        Show "* ユーザー: $username" Cyan
        Show "* OS: $os" Yellow
        Show "* CPU: $cpu  $cpuCores コア / $cpuThreads スレッド" Magenta
        Show "* メモリ: ${mem} GB" Blue
        Show "* ディスク ${diskId} 空き: ${free}GB / ${total}GB" White
    }
    Show "-----------------------------------------------------" Gray

    Write-Log "===== 高性能 PC 最適化ツール 実行ログ ====="
    Write-Log "[開始] $(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')"
    Write-Log "[ホスト] $hostname"
    Write-Log "[ユーザー] $username"
    Write-Log "[OS] $os"
    Write-Log "[CPU] $cpu  $cpuCores コア / $cpuThreads スレッド"
    Write-Log "[メモリ] $mem GB"
    Write-Log "[ディスク] ${diskId} 空き: $free GB / $total GB"
    Write-Log "[PowerShell バージョン] $pwv"
    $script:sysInfo = @{
        Hostname  = $hostname
        Username  = $username
        OS        = $os
        CPU       = "$cpu  $cpuCores コア / $cpuThreads スレッド"
        RAM       = "${mem} GB"
        Disk      = "${diskId} 空き ${free}GB / ${total}GB"
        PSVersion = $pwv
    }
    Write-Log "-----------------------------"
} catch {
    Write-ErrorLog "PC 情報の取得に失敗しました: $_"
}

# プログレスバー — PS7+ はブロック文字、PS5.x 以前はハッシュバー
function Progress-Bar ($msg, $percent) {
    $bars = [math]::Round($percent / 10)
    if ($isPS7Plus) {
        $filled = ([string][char]0x2588) * $bars
        $empty  = ([string][char]0x2591) * (10 - $bars)
        Write-Host "[$filled$empty] ${B}$msg${R}" -ForegroundColor Green
    } else {
        $bar = ('#' * $bars).PadRight(10, '-')
        Write-Host "[$bar] 進捗: $msg" -ForegroundColor Green
    }
}

# ディレクトリ内容を高速削除（Remove-Item -Recurse は大量ファイルでハングするため rd /s /q を使用）
function Clear-DirContents ([string]$Path) {
    Add-DeletedPathCandidate -Path $Path
    if ($script:GuardrailState -and (Get-Command Test-RepairAllowListPath -ErrorAction SilentlyContinue)) {
        $isAllowed = Test-RepairAllowListPath -State $script:GuardrailState -Path $Path
        if (-not $isAllowed) {
            $script:GuardrailState.BlockedPaths = @($script:GuardrailState.BlockedPaths) + @($Path)
            Write-Log "[guardrail] AllowList拒否: $Path"
            return
        }
    }
    if ($script:IsWhatIfMode) { return }
    if (-not (Test-Path $Path)) { return }
    Get-ChildItem -LiteralPath $Path -Force -ErrorAction SilentlyContinue |
        Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
    $null = New-Item -ItemType Directory -Path $Path -Force -ErrorAction SilentlyContinue
}

# タイムアウト付きサービス停止（Stop-Service は無限待機するため sc.exe で代替）
function Stop-ServiceSafe ([string]$Name, [int]$TimeoutSec = 20) {
    $svc = Get-Service -Name $Name -ErrorAction SilentlyContinue
    if (-not $svc -or $svc.Status -eq 'Stopped') { return }
    $null = & sc.exe stop $Name 2>$null          # 即時返却（待機しない）
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
        Start-Sleep -Milliseconds 500
        $svc.Refresh()
        if ($svc.Status -eq 'Stopped') { return }
    }
    Write-Log "[警告] $Name が ${TimeoutSec}秒以内に停止しませんでした（削除を試みます）"
}

# 各最適化ステップのラッパー
function Try-Step ($desc, [ScriptBlock]$action) {
    $script:TaskCounter++
    $taskId = $script:TaskCounter
    $start = Get-Date
    $sw    = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Log "[$($start.ToString('HH:mm:ss'))] $desc 開始..."
    if (Get-Command Invoke-AgentHookEvent -ErrorAction SilentlyContinue) {
        try {
            $hookContext = [PSCustomObject]@{ runId = $script:RunId; taskId = $taskId; taskName = $desc; mode = $script:RunMode; profile = $script:ExecutionProfile; stage = "start" }
            $script:HookHistory += @(Invoke-AgentHookEvent -EventName "pre_task" -Context $hookContext -HooksConfig $(if ($script:Config) { $script:Config.hooks } else { $null }) -RunId $script:RunId -LogsDir $logsDir)
        } catch {
            Write-Warning "[Hook] pre_task フック実行エラー（スキップ）: $_"
        }
    }
    Progress-Bar "$desc..." 0
    if ($script:FatalStopRequested) {
        $sw.Stop()
        $msg = "[FailFast] 先行タスク失敗のためスキップ: $desc"
        Show $msg Yellow
        Write-Log $msg
        $script:taskResults += [PSCustomObject]@{
            Id         = $taskId
            Name       = $desc
            Status     = "SKIP"
            Duration   = [int]$sw.Elapsed.TotalSeconds
            Error      = "FailFast skip"
            Errors     = @("FailFast skip")
            PreviewOnly= $false
        }
        Write-UiEvent -Type "task_finish" -Data @{
            taskId      = $taskId
            taskName    = $desc
            status      = "SKIP"
            durationSec = [int]$sw.Elapsed.TotalSeconds
            detail      = "FailFast skip"
        }
        return
    }
    if ($script:SelectedTaskSet -and -not $script:SelectedTaskSet.Contains($taskId)) {
        $sw.Stop()
        $msg = "[Tasks] 対象外タスクのためスキップ: #${taskId} $desc"
        Show $msg DarkGray
        Write-Log $msg
        $script:taskResults += [PSCustomObject]@{
            Id         = $taskId
            Name       = $desc
            Status     = "SKIP"
            Duration   = [int]$sw.Elapsed.TotalSeconds
            Error      = "TaskFiltered skip"
            Errors     = @("TaskFiltered skip")
            PreviewOnly= $script:IsWhatIfMode
        }
        return
    }
    Write-UiEvent -Type "task_start" -Data @{
        taskId   = $taskId
        taskName = $desc
    }
    if ($script:RunMode -eq 'diagnose' -and -not (Test-ReadOnlyTask -TaskName $desc)) {
        $sw.Stop()
        $msg = "[Diagnose] $desc は診断モードのため変更をスキップしました。"
        Show $msg Yellow
        Write-Log $msg
        $script:taskResults += [PSCustomObject]@{
            Id         = $taskId
            Name       = $desc
            Status     = "SKIP"
            Duration   = [int]$sw.Elapsed.TotalSeconds
            Error      = "DiagnoseMode skip"
            Errors     = @("DiagnoseMode skip")
            PreviewOnly= $true
        }
        Write-UiEvent -Type "task_finish" -Data @{
            taskId      = $taskId
            taskName    = $desc
            status      = "SKIP"
            durationSec = [int]$sw.Elapsed.TotalSeconds
            detail      = "DiagnoseMode skip"
        }
        return
    }
    if ($script:IsWhatIfMode -and -not (Test-ReadOnlyTask -TaskName $desc)) {
        $sw.Stop()
        $msg = "[WhatIf] $desc はプレビュー実行のため変更をスキップしました。"
        Show $msg Yellow
        Write-Log $msg
        if (Get-Command Register-TaskPlannedDeletedPath -ErrorAction SilentlyContinue) {
            Register-TaskPlannedDeletedPath -TaskId $taskId
        }
        $script:taskResults += [PSCustomObject]@{
            Id         = $taskId
            Name       = $desc
            Status     = "SKIP"
            Duration   = [int]$sw.Elapsed.TotalSeconds
            Error      = "WhatIf skip"
            Errors     = @("WhatIf skip")
            PreviewOnly= $true
        }
        Write-UiEvent -Type "task_finish" -Data @{
            taskId      = $taskId
            taskName    = $desc
            status      = "SKIP"
            durationSec = [int]$sw.Elapsed.TotalSeconds
            detail      = "WhatIf skip"
        }
        return
    }
    try {
        & $action
        $sw.Stop()
        Progress-Bar "$desc 完了" 100
        Write-Log "[$((Get-Date).ToString('HH:mm:ss'))] $desc 完了"
        $script:taskResults += [PSCustomObject]@{
            Id         = $taskId
            Name       = $desc
            Status     = "OK"
            Duration   = [int]$sw.Elapsed.TotalSeconds
            Error      = ""
            Errors     = @()
            PreviewOnly= $false
        }
        Write-UiEvent -Type "task_finish" -Data @{
            taskId      = $taskId
            taskName    = $desc
            status      = "OK"
            durationSec = [int]$sw.Elapsed.TotalSeconds
        }
        if (Get-Command Invoke-AgentHookEvent -ErrorAction SilentlyContinue) {
            try {
                $hookContext = [PSCustomObject]@{ runId = $script:RunId; taskId = $taskId; taskName = $desc; mode = $script:RunMode; profile = $script:ExecutionProfile; stage = "success" }
                $script:HookHistory += @(Invoke-AgentHookEvent -EventName "post_task" -Context $hookContext -HooksConfig $(if ($script:Config) { $script:Config.hooks } else { $null }) -RunId $script:RunId -LogsDir $logsDir)
            } catch {
                Write-Warning "[Hook] post_task フック実行エラー（スキップ）: $_"
            }
        }
    } catch {
        $sw.Stop()
        Progress-Bar "$desc 失敗" 100
        Write-Log "[$((Get-Date).ToString('HH:mm:ss'))] $desc 失敗"
        Write-ErrorLog "$desc : $_"
        $script:HadTaskFailure = $true
        if ($script:FailureMode -eq 'fail-fast') {
            $script:FatalStopRequested = $true
        }
        $script:taskResults += [PSCustomObject]@{
            Id         = $taskId
            Name       = $desc
            Status     = "NG"
            Duration   = [int]$sw.Elapsed.TotalSeconds
            Error      = "$_"
            Errors     = @("$_")
            PreviewOnly= $false
        }
        Write-UiEvent -Type "task_finish" -Data @{
            taskId      = $taskId
            taskName    = $desc
            status      = "NG"
            durationSec = [int]$sw.Elapsed.TotalSeconds
            detail      = "$_"
        }
        if (Get-Command Invoke-AgentHookEvent -ErrorAction SilentlyContinue) {
            try {
                $hookContext = [PSCustomObject]@{ runId = $script:RunId; taskId = $taskId; taskName = $desc; mode = $script:RunMode; profile = $script:ExecutionProfile; stage = "error"; error = "$_" }
                $script:HookHistory += @(Invoke-AgentHookEvent -EventName "on_error" -Context $hookContext -HooksConfig $(if ($script:Config) { $script:Config.hooks } else { $null }) -RunId $script:RunId -LogsDir $logsDir)
            } catch {
                Write-Warning "[Hook] on_error フック実行エラー（スキップ）: $_"
            }
        }
    }
}

# ── HTML レポート生成 ─────────────────────────────────────────────────
function New-HtmlReport {
    param(
        [PSCustomObject[]]$Results,
        [double]$DiskBefore,
        [double]$DiskAfter,
        [hashtable]$SysInfo,
        [pscustomobject]$AIDiagnosis
    )

    try {
        if (-not $isPS7Plus) { Add-Type -AssemblyName System.Web }
        $reportTime  = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
        $reportStamp = Get-Date -Format "yyyyMMddHHmm"
        $reportPath  = Join-Path $logsDir "PC_Optimizer_Report_${reportStamp}.html"
        $diskFreed   = [math]::Round($DiskAfter - $DiskBefore, 2)
        $diskSign    = [System.Web.HttpUtility]::HtmlEncode($(if ($diskFreed -ge 0) { "+${diskFreed}" } else { "${diskFreed}" }))
        $diskColor   = [System.Web.HttpUtility]::HtmlEncode($(if ($diskFreed -ge 0) { "#4caf50" } else { "#ff9800" }))

        # タスク結果の行 HTML を生成
        $taskRows = ($Results | ForEach-Object {
            $icon    = if ($_.Status -eq "OK") { "&#x2705;" } elseif ($_.Status -eq "SKIP") { "&#x23ED;&#xFE0F;" } else { "&#x274C;" }
            $rowCls  = if ($_.Status -eq "OK") { "row-ok" } elseif ($_.Status -eq "SKIP") { "row-skip" } else { "row-ng" }
            $nameEsc = [System.Web.HttpUtility]::HtmlEncode($_.Name)
            $errEsc  = [System.Web.HttpUtility]::HtmlEncode($_.Error)
            $errCell = if ($errEsc) { "<span class='err-detail'>$errEsc</span>" } else { "" }
            "<tr class='$rowCls'><td>$icon</td><td>$nameEsc</td><td>$([System.Web.HttpUtility]::HtmlEncode($_.Status))</td><td>$($_.Duration)s</td><td>$errCell</td></tr>"
        }) -join "`n"

        # システム情報行 HTML を生成
        $sysRows = ""
        if ($SysInfo) {
            $sysRows = ($SysInfo.GetEnumerator() | Sort-Object Key | ForEach-Object {
                $kEsc = [System.Web.HttpUtility]::HtmlEncode($_.Key)
                $vEsc = [System.Web.HttpUtility]::HtmlEncode($_.Value)
                "<tr><td class='sys-key'>$kEsc</td><td>$vEsc</td></tr>"
            }) -join "`n"
        }

        # OK / NG / SKIP カウント
        $okCount = ($Results | Where-Object { $_.Status -eq "OK" }).Count
        $ngCount = ($Results | Where-Object { $_.Status -eq "NG" }).Count
        $skipCount = ($Results | Where-Object { $_.Status -eq "SKIP" }).Count
        $unexecutedRows = ($Results | Where-Object { $_.Status -eq "SKIP" -and $_.Error -eq "FailFast skip" } | ForEach-Object {
            "<tr><td>#$($_.Id)</td><td>$([System.Web.HttpUtility]::HtmlEncode($_.Name))</td><td>$([System.Web.HttpUtility]::HtmlEncode($_.Error))</td></tr>"
        }) -join "`n"
        $unexecutedSection = ""
        if ($unexecutedRows) {
            $unexecutedSection = @"
<div class="card">
  <h2>&#x26A0;&#xFE0F; Fail-Fast 未実行タスク</h2>
  <table>
    <tr><th>Task ID</th><th>タスク名</th><th>理由</th></tr>
    $unexecutedRows
  </table>
</div>
"@
        }

        $aiSection = ""
        if ($AIDiagnosis) {
            $aiHeadline = [System.Web.HttpUtility]::HtmlEncode("$($AIDiagnosis.Headline)")
            $aiEval = [System.Web.HttpUtility]::HtmlEncode("$($AIDiagnosis.Evaluation)")
            $aiSummary = if ($AIDiagnosis.PSObject.Properties["Summary"] -and $AIDiagnosis.Summary) { [System.Web.HttpUtility]::HtmlEncode("$($AIDiagnosis.Summary)") } else { "" }
            $aiSource = [System.Web.HttpUtility]::HtmlEncode("$($AIDiagnosis.Source)")
            $aiPromptVersion = if ($AIDiagnosis.PSObject.Properties["PromptVersion"] -and $AIDiagnosis.PromptVersion) { [System.Web.HttpUtility]::HtmlEncode("$($AIDiagnosis.PromptVersion)") } else { "N/A" }
            $aiConfidence = if ($AIDiagnosis.PSObject.Properties["Confidence"] -and $null -ne $AIDiagnosis.Confidence) { [System.Web.HttpUtility]::HtmlEncode(("{0:P0}" -f [double]$AIDiagnosis.Confidence)) } else { "N/A" }
            $aiDataTimestamp = if ($AIDiagnosis.PSObject.Properties["DataTimestamp"] -and $AIDiagnosis.DataTimestamp) { [System.Web.HttpUtility]::HtmlEncode("$($AIDiagnosis.DataTimestamp)") } else { "N/A" }
            $aiFallback = if ($AIDiagnosis.PSObject.Properties["FallbackReason"] -and $AIDiagnosis.FallbackReason) { [System.Web.HttpUtility]::HtmlEncode("$($AIDiagnosis.FallbackReason)") } else { "" }
            $aiNarrativeRaw = if ($AIDiagnosis.PSObject.Properties["Narrative"] -and $AIDiagnosis.Narrative) {
                "$($AIDiagnosis.Narrative)".Trim()
            } else {
                ""
            }
            $aiNarrative = ""
            if (-not [string]::IsNullOrWhiteSpace($aiNarrativeRaw)) {
                $looksJsonBlock = $aiNarrativeRaw.StartsWith('```json') -or $aiNarrativeRaw.StartsWith('```') -or ($aiNarrativeRaw.StartsWith('{') -and $aiNarrativeRaw.EndsWith('}'))
                if (-not $looksJsonBlock) {
                    $aiNarrative = [System.Web.HttpUtility]::HtmlEncode($aiNarrativeRaw)
                }
            }
            $aiFindings = @($AIDiagnosis.Findings)
            $aiRecommendations = @($AIDiagnosis.Recommendations)
            $findingRows = if (@($aiFindings).Count -gt 0) {
                (@($aiFindings | ForEach-Object { "<li>$([System.Web.HttpUtility]::HtmlEncode("$_"))</li>" }) -join "`n")
            } else {
                "<li>該当なし</li>"
            }
            $actionRows = if (@($aiRecommendations).Count -gt 0) {
                (@($aiRecommendations | ForEach-Object { "<li>$([System.Web.HttpUtility]::HtmlEncode("$_"))</li>" }) -join "`n")
            } else {
                "<li>該当なし</li>"
            }
            $metricRows = ""
            if ($AIDiagnosis.PSObject.Properties["InputMetrics"] -and $AIDiagnosis.InputMetrics) {
                $metricRows = @(
                    $AIDiagnosis.InputMetrics.PSObject.Properties |
                    ForEach-Object {
                        $k = [System.Web.HttpUtility]::HtmlEncode("$($_.Name)")
                        $v = if ($null -eq $_.Value -or "$($_.Value)" -eq "") { "N/A" } else { [System.Web.HttpUtility]::HtmlEncode("$($_.Value)") }
                        "<tr><td class='sys-key'>$k</td><td>$v</td></tr>"
                    }
                ) -join "`n"
            }
            $narrativeBlock = if (-not [string]::IsNullOrWhiteSpace($aiNarrative)) {
@"
  <div class="ai-narrative">$aiNarrative</div>
"@
            } else { "" }
            $fallbackBlock = if ($aiFallback) { "<p class='ai-meta'><strong>フォールバック理由:</strong> $aiFallback</p>" } else { "" }
            $summaryBlock = if ($aiSummary) { "<p class='ai-meta'><strong>サマリー:</strong> $aiSummary</p>" } else { "" }
            $metricTableBlock = if ($metricRows) { "<h3>入力指標一覧</h3><table>$metricRows</table>" } else { "" }
            $aiSection = @"
<div class="card">
  <h2>&#x1F9E0; AI 診断</h2>
  <p class="ai-meta"><strong>要約:</strong> $aiHeadline</p>
  $summaryBlock
  <p class="ai-meta"><strong>評価:</strong> $aiEval <span class="ai-source">ソース: $aiSource</span></p>
  <p class="ai-meta"><strong>Prompt Version:</strong> $aiPromptVersion</p>
  <p class="ai-meta"><strong>信頼度:</strong> $aiConfidence</p>
  <p class="ai-meta"><strong>対象データ時刻:</strong> $aiDataTimestamp</p>
  $fallbackBlock
$narrativeBlock
  <div class="ai-grid">
    <div>
      <h3>根拠</h3>
      <ul>$findingRows</ul>
    </div>
    <div>
      <h3>推奨アクション</h3>
      <ol>$actionRows</ol>
    </div>
  </div>
  $metricTableBlock
</div>
"@
        }

        $html = @"
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>PC Optimizer Report</title>
<style>
:root {
  --bg:#ffffff; --card:#f0f8ff; --accent:#5b9bd5; --ok:#2e8b57;
  --ng:#d32f2f; --text:#222222; --sub:#6b7280; --border:#b8d8f5;
}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--text);font-family:'Segoe UI',sans-serif;font-size:14px;line-height:1.6}
header{background:linear-gradient(135deg,#5b9bd5,#3a7bc8);padding:2rem;text-align:center}
header h1{font-size:1.8rem;font-weight:700;letter-spacing:.05em;color:#ffffff}
header p{color:rgba(255,255,255,.85);margin-top:.5rem}
.container{max-width:960px;margin:2rem auto;padding:0 1rem}
.card{background:var(--card);border:1px solid var(--border);border-radius:8px;padding:1.5rem;margin-bottom:1.5rem}
.card h2{font-size:1rem;text-transform:uppercase;letter-spacing:.1em;color:var(--accent);margin-bottom:1rem}
.summary-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:1rem}
.summary-item{text-align:center;padding:.8rem;background:#d6eaf8;border:1px solid var(--border);border-radius:6px}
.summary-item .val{font-size:2rem;font-weight:700;color:var(--accent)}
.summary-item .lbl{font-size:.75rem;color:var(--sub);margin-top:.2rem}
table{width:100%;border-collapse:collapse}
th{text-align:left;padding:.6rem .8rem;background:#c8dfff;color:#1a4a7a;font-size:.8rem;text-transform:uppercase;letter-spacing:.05em}
td{padding:.5rem .8rem;border-bottom:1px solid var(--border)}
.row-ok td:first-child{color:var(--ok)}
.row-ng td:first-child{color:var(--ng)}
.row-skip td:first-child{color:#d97706}
.row-ng{background:rgba(211,47,47,.05)}
.row-skip{background:rgba(217,119,6,.08)}
.err-detail{color:#d97706;font-size:.8rem}
.sys-key{color:var(--sub);width:120px}
.ai-meta{margin:.2rem 0 .5rem}
.ai-source{display:inline-block;margin-left:.8rem;font-size:.8rem;color:var(--sub)}
.ai-grid{display:grid;grid-template-columns:1fr 1fr;gap:1rem}
.ai-grid h3{font-size:.9rem;margin:0 0 .4rem;color:#1a4a7a}
.ai-grid ul,.ai-grid ol{padding-left:1.2rem}
.ai-narrative{white-space:pre-wrap;background:#ffffff;border:1px solid var(--border);padding:.75rem;border-radius:6px;margin:.6rem 0 .8rem}
@media (max-width: 860px){.ai-grid{grid-template-columns:1fr}}
footer{text-align:center;color:var(--sub);font-size:.8rem;padding:2rem}
</style>
</head>
<body>
<header>
  <h1>&#x1F4BB; PC Optimizer Report</h1>
  <p>$reportTime</p>
</header>
<div class="container">

<div class="card">
  <h2>&#x1F4CA; サマリー</h2>
  <div class="summary-grid">
    <div class="summary-item"><div class="val">$($Results.Count)</div><div class="lbl">総タスク数</div></div>
    <div class="summary-item"><div class="val" style="color:var(--ok)">$okCount</div><div class="lbl">成功</div></div>
    <div class="summary-item"><div class="val" style="color:var(--ng)">$ngCount</div><div class="lbl">失敗</div></div>
    <div class="summary-item"><div class="val" style="color:#d97706">$skipCount</div><div class="lbl">SKIP</div></div>
    <div class="summary-item"><div class="val" style="color:$diskColor">$diskSign GB</div><div class="lbl">ディスク解放量</div></div>
  </div>
</div>

$unexecutedSection

<div class="card">
  <h2>&#x1F4BE; ディスク空き容量</h2>
  <table>
    <tr><th>タイミング</th><th>空き容量</th></tr>
    <tr><td>最適化前</td><td>$([math]::Round($DiskBefore, 2)) GB</td></tr>
    <tr><td>最適化後</td><td>$([math]::Round($DiskAfter, 2)) GB</td></tr>
    <tr><td><strong>解放量</strong></td><td><strong style="color:$diskColor">$diskSign GB</strong></td></tr>
  </table>
</div>

<div class="card">
  <h2>&#x1F5A5;&#xFE0F; システム情報</h2>
  <table>$sysRows</table>
</div>

$aiSection

<div class="card">
  <h2>&#x2705; タスク実行結果</h2>
  <table>
    <tr><th></th><th>タスク名</th><th>ステータス</th><th>所要時間</th><th>エラー詳細</th></tr>
    $taskRows
  </table>
</div>

</div>
<footer>Generated by PC Optimizer &mdash; $reportTime</footer>
</body>
</html>
"@

        [System.IO.File]::WriteAllText($reportPath, $html, [System.Text.Encoding]::UTF8)
        Write-Log "[HTMLレポート] 保存完了: $reportPath"
        Show "  HTMLレポートを開いています..." Cyan
        Start-Process $reportPath
    } catch {
        Write-Log "[HTMLレポート] 生成失敗（警告のみ）: $_"
        Write-ErrorLog "New-HtmlReport: $_"
    }
}

# 再起動保留の確認
function Test-PendingReboot {
    $rebootRequired = $false
    try {
        if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired") {
            $rebootRequired = $true
        }
        if (Test-Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager") {
            $p = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" `
                                  -Name "PendingFileRenameOperations" -ErrorAction SilentlyContinue
            if ($p -and $p.PendingFileRenameOperations) {
                $rebootRequired = $true
            }
        }
    } catch {
        Write-ErrorLog "再起動保留状態の確認に失敗しました: $_"
    }
    return $rebootRequired
}

# Windows Update の利用可能な更新を一覧表示し、確認後に適用する（Windows 10/11）
function Run-WindowsUpdate {
    if (Get-Command Invoke-UpdateMaintenance -ErrorAction SilentlyContinue) {
        if ($script:IsWhatIfMode) {
            $preview = Invoke-UpdateMaintenance -WhatIfMode
            Write-Log "[Windows Update][module] WhatIf: $($preview.Status)"
            Show "Windows Update（モジュール）: WhatIf プレビューのみ実施しました。" DarkGray
            return
        }
        $choice = Get-UserChoice -Prompt "${B}Windows Update を実行しますか？${RB} (Y/N)" -Default 'N'
        Write-Log "[Windows Update][module] ユーザー選択: $choice"
        if (Test-ChoiceYes -Choice $choice) {
            $result = Invoke-UpdateMaintenance
            Write-Log "[Windows Update][module] 実行結果: $($result.Status)"
            Show "Windows Update（モジュール）を実行しました。" Green
        } else {
            Show "Windows Update（モジュール）をスキップしました。" Yellow
            Write-Log "[Windows Update][module] ユーザーによりスキップ"
        }
        return
    }

    # UsoClient.exe は Windows 10/11 でのみ利用可能
    $usoPath = Join-Path $env:SystemRoot "System32\UsoClient.exe"
    if (-not (Test-Path $usoPath)) {
        $msg = "UsoClient.exe が見つかりません。Windows Update をスキップします。"
        Show $msg DarkGray
        Write-Log "[Windows Update] $msg"
        return
    }

    # COM API で利用可能な更新を検索（追加モジュール不要）
    Show "Windows Update を確認中... (しばらくお待ちください)" Cyan
    Write-Log "[Windows Update] 利用可能な更新を検索中..."

    $updates     = $null
    $updateCount = 0
    try {
        $session     = New-Object -ComObject Microsoft.Update.Session
        $searcher    = $session.CreateUpdateSearcher()
        $result      = $searcher.Search("IsInstalled=0 and IsHidden=0")
        $updates     = $result.Updates
        $updateCount = $updates.Count
    } catch {
        # COM API が利用できない環境（稀）は件数不明のまま続行
        Write-ErrorLog "[Windows Update] 更新一覧の取得に失敗: $_"
        Show "更新一覧を取得できませんでした。件数不明のまま続行します。" Yellow
    }

    if ($updateCount -eq 0 -and $null -ne $updates) {
        $msg = "利用可能な Windows Update はありません。"
        Show $msg Green
        Write-Log "[Windows Update] $msg"
        return
    }

    # 更新一覧の表示（最大 10 件、それ以上は省略）
    if ($updateCount -gt 0) {
        Show "${B}利用可能な更新: ${updateCount} 件${R}" Yellow
        Write-Log "[Windows Update] 利用可能な更新: ${updateCount} 件"
        $showCount = [Math]::Min($updateCount, 10)
        for ($i = 0; $i -lt $showCount; $i++) {
            $line = "  [$($i+1)] $($updates.Item($i).Title)"
            Show $line White
            Write-Log "[Windows Update] $line"
        }
        if ($updateCount -gt 10) {
            $more = "  ... 他 $($updateCount - 10) 件"
            Show $more DarkGray
            Write-Log "[Windows Update] $more"
        }
    }

    # ユーザー確認
    $choice = Get-UserChoice -Prompt "${B}これらの更新を適用しますか？${RB} (Y/N)" -Default 'N'
    Write-Log "[Windows Update] ユーザー選択: $choice"
    if (Test-ChoiceYes -Choice $choice) {
        Show "Windows Update を開始します..." Cyan
        Write-Log "[Windows Update] ユーザーが更新を承認。UsoClient でインストール開始。"
        Start-Process -FilePath $usoPath -ArgumentList "StartScan"     -NoNewWindow -Wait
        Start-Process -FilePath $usoPath -ArgumentList "StartDownload" -NoNewWindow -Wait
        Start-Process -FilePath $usoPath -ArgumentList "StartInstall"  -NoNewWindow -Wait
        Write-Log "[Windows Update] 完了。"
        Show "Windows Update のインストールを開始しました（完了まで数分〜数十分かかる場合があります）。" Green
    } else {
        $msg = "Windows Update をスキップしました。"
        Show $msg Yellow
        Write-Log "[Windows Update] $msg"
    }
}

# ==========================
# 初期ディスク空き容量の記録（前後比較用）
# ==========================
$initialFreeGB = 0
try {
    $initialMaybe = Get-SystemDriveFreeGB -DriveLetter 'C:'
    if ($null -ne $initialMaybe) { $initialFreeGB = $initialMaybe }
} catch {
    Write-Host "  初期ディスク空き容量の取得に失敗しました（スキップ）" -ForegroundColor DarkGray
}

# ==========================
# メイン最適化・クリーンアップ
# ==========================
Try-Step "一時ファイルの削除" {
    if (Get-Command Invoke-CleanupMaintenance -ErrorAction SilentlyContinue) {
        Invoke-CleanupMaintenance -WhatIfMode:$script:IsWhatIfMode -Tasks @('temp') | Out-Null
    } else {
        Clear-DirContents "$env:SystemRoot\Temp"
        Clear-DirContents $env:TEMP
        Clear-DirContents "$env:USERPROFILE\AppData\Local\Temp"
    }
}

Try-Step "Prefetch・更新キャッシュの削除" {
    Clear-DirContents "$env:SystemRoot\Prefetch"

    # Windows Update サービスを停止してからキャッシュ削除（ロック回避）
    Stop-ServiceSafe -Name wuauserv
    Clear-DirContents "$env:SystemRoot\SoftwareDistribution\Download"
    Start-Service -Name wuauserv -ErrorAction SilentlyContinue

    # 配信最適化サービスを停止してからキャッシュ削除（ロック回避）
    Stop-ServiceSafe -Name DoSvc
    Clear-DirContents "$env:SystemRoot\System32\DeliveryOptimization\Cache"
    Start-Service -Name DoSvc -ErrorAction SilentlyContinue
}

Try-Step "配信最適化キャッシュの削除" {
    Stop-ServiceSafe -Name wuauserv
    Clear-DirContents "$env:SystemRoot\SoftwareDistribution\DeliveryOptimization\Cache"
    Start-Service -Name wuauserv -ErrorAction SilentlyContinue
}

Try-Step "Windows Update キャッシュの削除" {
    Stop-ServiceSafe -Name wuauserv
    Clear-DirContents "$env:SystemRoot\SoftwareDistribution\Download"
    Start-Service -Name wuauserv -ErrorAction SilentlyContinue
}

Try-Step "エラーレポート・ログ・不要キャッシュの削除" {
    Clear-DirContents "$env:ProgramData\Microsoft\Windows\WER\ReportArchive"
    Clear-DirContents "$env:ProgramData\Microsoft\Windows\WER\ReportQueue"
    Clear-DirContents "$env:SystemRoot\Logs\CBS"
}

Try-Step "OneDrive / Teams / Office キャッシュの削除" {
    Clear-DirContents "$env:LOCALAPPDATA\Microsoft\OneDrive\logs"
    Clear-DirContents "$env:APPDATA\Microsoft\Teams\Cache"
    Clear-DirContents "$env:LOCALAPPDATA\Microsoft\Office\16.0\OfficeFileCache"
}

Try-Step "ブラウザキャッシュの削除（Chrome / Edge / Firefox / Brave / Opera / Vivaldi）" {
    if (Get-Command Invoke-CleanupMaintenance -ErrorAction SilentlyContinue) {
        Invoke-CleanupMaintenance -WhatIfMode:$script:IsWhatIfMode -Tasks @('browser') | Out-Null
    } else {
        # Google Chrome
        @(
            "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache",
            "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Code Cache",
            "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\GPUCache"
        ) | ForEach-Object { Clear-DirContents $_ }
        # Microsoft Edge
        @(
            "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache",
            "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Code Cache",
            "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\GPUCache"
        ) | ForEach-Object { Clear-DirContents $_ }
        # Mozilla Firefox（全プロファイル対応）
        $ffProfileDir = "$env:APPDATA\Mozilla\Firefox\Profiles"
        if (Test-Path $ffProfileDir) {
            Get-ChildItem -Path $ffProfileDir -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                Clear-DirContents (Join-Path $_.FullName "cache2\entries")
            }
        }
        # Brave Browser（Chromium ベース）
        @(
            "$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\User Data\Default\Cache",
            "$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\User Data\Default\Code Cache",
            "$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\User Data\Default\GPUCache"
        ) | ForEach-Object { Clear-DirContents $_ }
        # Opera / Opera GX（Chromium ベース）
        @(
            "$env:APPDATA\Opera Software\Opera Stable\Cache",
            "$env:APPDATA\Opera Software\Opera Stable\Code Cache",
            "$env:APPDATA\Opera Software\Opera GX Stable\Cache",
            "$env:APPDATA\Opera Software\Opera GX Stable\Code Cache"
        ) | ForEach-Object { Clear-DirContents $_ }
        # Vivaldi（Chromium ベース）
        @(
            "$env:LOCALAPPDATA\Vivaldi\User Data\Default\Cache",
            "$env:LOCALAPPDATA\Vivaldi\User Data\Default\Code Cache",
            "$env:LOCALAPPDATA\Vivaldi\User Data\Default\GPUCache"
        ) | ForEach-Object { Clear-DirContents $_ }
    }
}

Try-Step "サムネイルキャッシュの削除" {
    $explorerDir = "$env:LOCALAPPDATA\Microsoft\Windows\Explorer"
    Add-DeletedPathCandidate "$explorerDir\thumbcache_*.db"
    if (Test-Path $explorerDir) {
        Get-ChildItem -Path $explorerDir -Filter "thumbcache_*.db" -ErrorAction SilentlyContinue |
            Remove-Item -Force -ErrorAction SilentlyContinue
    }
}

function Get-RunStatus {
    if ($script:taskResults.Count -eq 0) { return "NG" }
    $ngCount   = @($script:taskResults | Where-Object { $_.Status -eq 'NG' }).Count
    $skipCount = @($script:taskResults | Where-Object { $_.Status -eq 'SKIP' }).Count
    if ($ngCount -gt 0 -and $skipCount -gt 0) { return "PARTIAL" }
    if ($ngCount -gt 0) { return "NG" }
    if ($skipCount -gt 0) { return "PARTIAL" }
    return "OK"
}

function Resolve-ExitCode {
    if ($script:ExitCode -ne 0) { return $script:ExitCode }
    if ($script:HadTaskFailure) {
        if ($script:FailureMode -eq 'fail-fast') {
            return $script:ExitCodes.Fatal
        }
        return $script:ExitCodes.Partial
    }
    return $script:ExitCodes.Success
}

function Invoke-SafeModuleCall {
    param(
        [Parameter(Mandatory)]
        [string]$CommandName,
        [hashtable]$Parameters = @{},
        [object]$Default = $null
    )
    if (-not (Get-Command $CommandName -ErrorAction SilentlyContinue)) {
        Write-Log "[module] コマンド未検出: $CommandName"
        return $Default
    }
    try {
        return & $CommandName @Parameters
    } catch {
        Write-ErrorLog "[module] $CommandName 実行失敗: $_"
        return $Default
    }
}

function Get-ComponentScore {
    param(
        [double]$Value,
        [double]$GoodThreshold,
        [double]$WarnThreshold
    )
    if ($Value -ge $GoodThreshold) { return 100 }
    if ($Value -ge $WarnThreshold) { return 75 }
    if ($Value -ge 1) { return 50 }
    return 20
}

function Invoke-IntegratedModuleDiagnostic {
    $diag = Invoke-SafeModuleCall -CommandName 'Get-SystemDiagnostic' -Default ([PSCustomObject]@{})
    $asset = Invoke-SafeModuleCall -CommandName 'Get-AssetInventory' -Default ([PSCustomObject]@{})
    $eventSummary = Invoke-SafeModuleCall -CommandName 'Get-EventLogSummary' -Parameters @{ Hours = 24 } -Default ([PSCustomObject]@{})
    $perf = Invoke-SafeModuleCall -CommandName 'Get-PerformanceSnapshot' -Default ([PSCustomObject]@{})
    $startup = Invoke-SafeModuleCall -CommandName 'Get-StartupAnalysis' -Default ([PSCustomObject]@{ StartupCount = 0; Rating = 'Unknown'; Items = @() })
    $security = Invoke-SafeModuleCall -CommandName 'Get-SecurityDiagnostic' -Default ([PSCustomObject]@{})
    $network = Invoke-SafeModuleCall -CommandName 'Get-NetworkDiagnostic' -Default ([PSCustomObject]@{})
    $update = Invoke-SafeModuleCall -CommandName 'Get-UpdateDiagnostic' -Default ([PSCustomObject]@{})

    $cpuScore = 80
    $memoryScore = 80
    $diskScore = 80
    try {
        if ($diag.Memory -and $diag.Memory.TotalGB -and $null -ne $diag.Memory.FreeGB) {
            $freeMemPct = 0
            if ([double]$diag.Memory.TotalGB -gt 0) {
                $freeMemPct = ([double]$diag.Memory.FreeGB / [double]$diag.Memory.TotalGB) * 100
            }
            $memoryScore = Get-ComponentScore -Value $freeMemPct -GoodThreshold 30 -WarnThreshold 15
        }
        if ($diag.Disk -and @($diag.Disk).Count -gt 0) {
            $sysDisk = @($diag.Disk | Where-Object { $_.Drive -eq 'C:' } | Select-Object -First 1)
            if (-not $sysDisk) { $sysDisk = @($diag.Disk | Select-Object -First 1) }
            if ($sysDisk -and $null -ne $sysDisk[0].FreePercent) {
                $diskScore = Get-ComponentScore -Value ([double]$sysDisk[0].FreePercent) -GoodThreshold 25 -WarnThreshold 10
            }
        }
    } catch {
        Write-ErrorLog "[module] スコア計算補助データの評価失敗: $_"
    }

    $startupScore = if ($startup.Rating -eq 'Good') { 100 } elseif ($startup.Rating -eq 'Normal') { 75 } elseif ($startup.Rating -eq 'High') { 50 } else { 60 }
    $securityOkCount = 0
    foreach ($k in @('Defender','Firewall','BitLocker','Uac')) {
        if ($security.PSObject.Properties[$k] -and $security.$k -in @('Enabled','Warning')) {
            if ($security.$k -eq 'Enabled') { $securityOkCount++ }
        }
    }
    $securityScore = switch ($securityOkCount) {
        4 { 100 }
        3 { 80 }
        2 { 65 }
        1 { 45 }
        default { 25 }
    }

    $networkScore = if ($network.IpAddress -and @($network.IpAddress).Count -gt 0) { 85 } else { 40 }
    $windowsUpdateScore = if ($update.WindowsUpdate -eq 'Compliant') { 100 } elseif ($update.WindowsUpdate -eq 'PendingUpdates') { 60 } else { 50 }
    $systemHealthScore = if ($eventSummary.BsodCount -eq 0 -and $eventSummary.SystemErrors -lt 5) { 90 } elseif ($eventSummary.BsodCount -le 1) { 70 } else { 40 }

    $health = Invoke-SafeModuleCall -CommandName 'Get-HealthScore' -Parameters @{
        Cpu = [int]$cpuScore
        Memory = [int]$memoryScore
        Disk = [int]$diskScore
        Startup = [int]$startupScore
        Security = [int]$securityScore
        Network = [int]$networkScore
        WindowsUpdate = [int]$windowsUpdateScore
        SystemHealth = [int]$systemHealthScore
    } -Default ([PSCustomObject]@{ Score = 0; Status = 'Critical' })

    $script:ModuleSnapshot = [PSCustomObject]@{
        systemDiagnostics = $diag
        assetInventory = $asset
        eventLogSummary = $eventSummary
        performanceSnapshot = $perf
        startupAnalysis = $startup
        securityDiagnostics = $security
        networkDiagnostics = $network
        updateDiagnostics = $update
    }
    $script:HealthScore = $health

    $script:UpdateErrorClassification = @()
    if (Get-Command Get-UpdateErrorClassification -ErrorAction SilentlyContinue) {
        $script:UpdateErrorClassification = @(Get-UpdateErrorClassification -UpdateErrors @($update.UpdateErrors))
    }
    $script:M365Connectivity = @()
    if (Get-Command Test-M365Connectivity -ErrorAction SilentlyContinue) {
        $script:M365Connectivity = @(Test-M365Connectivity)
    }
    $script:EventAnomaly = $null
    if (Get-Command Get-EventLogAnomaly -ErrorAction SilentlyContinue) {
        $script:EventAnomaly = Get-EventLogAnomaly -Hours 24
    }
    $script:BootTrend = $null
    if (Get-Command Update-BootShutdownTrend -ErrorAction SilentlyContinue) {
        $trendDir = Join-Path $logsDir "trends"
        $script:BootTrend = Update-BootShutdownTrend -OutputDir $trendDir
    }

    Write-Log "[module] 診断データ収集完了: score=$($health.Score), status=$($health.Status)"
}

function Export-IntegratedModuleReport {
    if (-not (Get-Command Export-OptimizerReport -ErrorAction SilentlyContinue)) {
        Write-Log "[module-report] Report モジュール未読込のため統合レポートをスキップ"
        return
    }
    if (-not $script:ModuleSnapshot) {
        Write-Log "[module-report] 診断スナップショットが無いため統合レポートをスキップ"
        return
    }

    $stamp = Get-Date -Format "yyyyMMddHHmmss"
    $reportBase = [PSCustomObject]@{
        computerName  = $env:COMPUTERNAME
        userName      = $env:USERNAME
        score         = if ($script:HealthScore) { $script:HealthScore.Score } else { 0 }
        status        = if ($script:HealthScore) { (Convert-HealthStatusToJapanese -Status "$($script:HealthScore.Status)") } else { '不明' }
        cpuScore      = if ($script:HealthScore -and $script:HealthScore.ScoreInput) { $script:HealthScore.ScoreInput.Cpu } else { 0 }
        memoryScore   = if ($script:HealthScore -and $script:HealthScore.ScoreInput) { $script:HealthScore.ScoreInput.Memory } else { 0 }
        diskScore     = if ($script:HealthScore -and $script:HealthScore.ScoreInput) { $script:HealthScore.ScoreInput.Disk } else { 0 }
        startupScore  = if ($script:HealthScore -and $script:HealthScore.ScoreInput) { $script:HealthScore.ScoreInput.Startup } else { 0 }
        securityScore = if ($script:HealthScore -and $script:HealthScore.ScoreInput) { $script:HealthScore.ScoreInput.Security } else { 0 }
        networkScore  = if ($script:HealthScore -and $script:HealthScore.ScoreInput) { $script:HealthScore.ScoreInput.Network } else { 0 }
        updateScore   = if ($script:HealthScore -and $script:HealthScore.ScoreInput) { $script:HealthScore.ScoreInput.WindowsUpdate } else { 0 }
        generatedAt   = (Get-Date).ToString("s")
        diagnostics   = $script:ModuleSnapshot
        updateErrorClassification = @($script:UpdateErrorClassification)
        m365Connectivity = @($script:M365Connectivity)
        eventAnomaly = $script:EventAnomaly
        bootTrend = $script:BootTrend
        aiDiagnosis = $script:AIDiagnosis
        agentSummary = $script:AgentTeamsResult
    }

    $reportData = New-OptimizerReportData -InputObject $reportBase
    $jsonPath = Join-Path $reportsDir "PC_Health_Report_${stamp}.json"
    $csvPath  = Join-Path $reportsDir "PC_Health_Report_${stamp}.csv"
    $htmlPath = Join-Path $reportsDir "PC_Health_Report_${stamp}.html"
    $useLocalChart = [bool]$UseLocalChartJs
    if ($useLocalChart) {
        $copied = Ensure-ReportChartAsset -ReportsDir $reportsDir -ScriptRoot $PSScriptRoot -RelativePath $ChartJsLocalRelativePath
        if (-not $copied) {
            Write-Log "[report] ローカル Chart.js が見つからないためフォールバック表示を使用します: $ChartJsLocalRelativePath"
        }
    }

    Export-OptimizerReport -ReportData $reportData.Data -Format json -Path $jsonPath | Out-Null
    Export-OptimizerReport -ReportData $reportData.Data -Format csv -Path $csvPath | Out-Null
    Export-OptimizerReport -ReportData $reportData.Data -Format html -Path $htmlPath -UseLocalChartJs:$useLocalChart -ChartJsScriptPath $ChartJsLocalRelativePath | Out-Null
    Write-Log "[module-report] 保存完了: $jsonPath"
    Write-Log "[module-report] 保存完了: $csvPath"
    Write-Log "[module-report] 保存完了: $htmlPath"
}

function Export-JsonExecutionReport {
    if (-not (Get-Command Export-OptimizerReport -ErrorAction SilentlyContinue)) {
        Write-Log "[JSONレポート] modules\\Report.psm1 が未読込のためスキップ"
        return $null
    }

    $finishedAt = Get-Date
    $hostInfo = [PSCustomObject]@{
        hostname  = $env:COMPUTERNAME
        os        = if ($script:sysInfo) { $script:sysInfo.OS } else { $env:OS }
        psVersion = $PSVersionTable.PSVersion.ToString()
    }
    $tasks = @($script:taskResults | ForEach-Object {
        [PSCustomObject]@{
            id          = if ($_.Id) { [int]$_.Id } else { 0 }
            name        = $_.Name
            status      = $_.Status
            duration    = [double]$_.Duration
            errors      = if ($_.Errors) { @($_.Errors) } elseif ($_.Error) { @("$($_.Error)") } else { @() }
            previewOnly = [bool]$_.PreviewOnly
        }
    })

    $unexecuted = @($script:taskResults | Where-Object { $_.Status -eq 'SKIP' -and $_.Error -eq 'FailFast skip' } | ForEach-Object { $_.Id })
    $selectedTasks = @()
    if ($script:SelectedTaskSet) {
        $selectedTasks = @($script:SelectedTaskSet | Sort-Object)
    } else {
        $selectedTasks = @(1..$script:TaskCounter)
    }
    $skipSummary = @{}
    foreach ($t in ($script:taskResults | Where-Object { $_.Status -eq 'SKIP' })) {
        $key = if ($t.Error) { "$($t.Error)" } else { "Unknown skip" }
        if (-not $skipSummary.ContainsKey($key)) { $skipSummary[$key] = 0 }
        $skipSummary[$key]++
    }
    $jsonObject = [PSCustomObject]@{
        version         = "1.0"
        runId           = $script:RunId
        startedAt       = $script:scriptStartTime.ToUniversalTime().ToString("o")
        finishedAt      = $finishedAt.ToUniversalTime().ToString("o")
        host            = $hostInfo
        status          = Get-RunStatus
        exitCode        = Resolve-ExitCode
        failureMode     = $script:FailureMode
        durationSeconds = [math]::Round(($finishedAt - $script:scriptStartTime).TotalSeconds, 3)
        selectedTasks   = $selectedTasks
        skippedReasonSummary = [PSCustomObject]$skipSummary
        unexecutedTasks = $unexecuted
        tasks           = $tasks
        healthScore     = if ($script:HealthScore) { $script:HealthScore } else { $null }
        moduleSnapshot  = if ($script:ModuleSnapshot) { $script:ModuleSnapshot } else { $null }
        updateErrorClassification = @($script:UpdateErrorClassification)
        m365Connectivity = @($script:M365Connectivity)
        eventAnomaly = $script:EventAnomaly
        bootTrend = $script:BootTrend
        aiDiagnosis = $script:AIDiagnosis
    }

    $reportData = New-OptimizerReportData -InputObject $jsonObject
    $jsonPath = Join-Path $logsDir ("PC_Optimizer_Report_{0}.json" -f (Get-Date -Format "yyyyMMddHHmmss"))
    Export-OptimizerReport -ReportData $reportData.Data -Format json -Path $jsonPath | Out-Null
    Write-Log "[JSONレポート] 保存完了: $jsonPath"
    return $jsonPath
}

function Export-RunAuditJson {
    [CmdletBinding()]
    param()

    if (-not (Test-Path $reportsDir)) {
        New-Item -ItemType Directory -Path $reportsDir -Force | Out-Null
    }

    $stamp = Get-Date -Format "yyyyMMddHHmmss"
    $auditPath = Join-Path $reportsDir ("Audit_Run_{0}_{1}.json" -f $script:RunId, $stamp)
    $okCount = @($script:taskResults | Where-Object { $_.Status -eq 'OK' }).Count
    $ngCount = @($script:taskResults | Where-Object { $_.Status -eq 'NG' }).Count
    $skipCount = @($script:taskResults | Where-Object { $_.Status -eq 'SKIP' }).Count

    $audit = [PSCustomObject]@{
        schemaVersion = "1.0"
        runId = $script:RunId
        generatedAt = (Get-Date).ToString("s")
        execution = [PSCustomObject]@{
            mode = $script:RunMode
            whatIf = $script:IsWhatIfMode
            nonInteractive = $script:IsNonInteractive
            failureMode = $script:FailureMode
            executionProfile = $script:ExecutionProfile
            selectedTasks = if ($script:SelectedTaskSet) { @($script:SelectedTaskSet | Sort-Object) } else { @() }
        }
        summary = [PSCustomObject]@{
            total = @($script:taskResults).Count
            success = $okCount
            ng = $ngCount
            skipped = $skipCount
            hadFailure = [bool]$script:HadTaskFailure
            exitCode = [int]$script:ExitCode
        }
        ai = [PSCustomObject]@{
            enabled = [bool]$EnableAIDiagnosis
            requestedAnthropic = [bool]$UseAnthropicAI
            source = if ($script:AIDiagnosis) { $script:AIDiagnosis.Source } else { $null }
            evaluation = if ($script:AIDiagnosis) { $script:AIDiagnosis.Evaluation } else { $null }
            promptVersion = if ($script:AIDiagnosis -and $script:AIDiagnosis.PSObject.Properties["PromptVersion"]) { $script:AIDiagnosis.PromptVersion } else { $null }
            confidence = if ($script:AIDiagnosis -and $script:AIDiagnosis.PSObject.Properties["Confidence"]) { $script:AIDiagnosis.Confidence } else { $null }
            fallbackReason = if ($script:AIDiagnosis -and $script:AIDiagnosis.PSObject.Properties["FallbackReason"]) { $script:AIDiagnosis.FallbackReason } else { $null }
        }
        agentTeams = if ($script:AgentTeamsResult) { $script:AgentTeamsResult } elseif ($script:AgentTeamsSummary -and $script:AgentTeamsSummary.summary) { $script:AgentTeamsSummary.summary } else { $null }
        hooks = [PSCustomObject]@{
            count = @($script:HookHistory).Count
            events = @($script:HookHistory)
        }
        mcp = @($script:McpResults)
        config = $script:Config
        changedTargets = [PSCustomObject]@{
            deletedPathCandidates = @($script:DeletedPathList | Sort-Object -Unique)
            blockedPaths = @(
                if ($script:GuardrailState -and $script:GuardrailState.PSObject.Properties["BlockedPaths"] -and $script:GuardrailState.BlockedPaths) {
                    $script:GuardrailState.BlockedPaths | Sort-Object -Unique
                }
            )
        }
        outputs = [PSCustomObject]@{
            logsDir = $logsDir
            reportsDir = $reportsDir
        }
    }

    $audit | ConvertTo-Json -Depth 12 | Set-Content -Path $auditPath -Encoding utf8
    return $auditPath
}

function Export-AuditDiffJson {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$CurrentAuditPath
    )

    $all = @(Get-ChildItem -Path $reportsDir -Filter "Audit_Run_*.json" -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTimeUtc)
    if (@($all).Count -lt 2) { return $null }
    $currentFile = Get-Item -Path $CurrentAuditPath -ErrorAction SilentlyContinue
    if (-not $currentFile) { return $null }
    $previousFile = @($all | Where-Object { $_.FullName -ne $currentFile.FullName } | Sort-Object LastWriteTimeUtc | Select-Object -Last 1)
    if (-not $previousFile) { return $null }

    $curr = Get-Content -Path $currentFile.FullName -Raw -Encoding utf8 | ConvertFrom-Json
    $prev = Get-Content -Path $previousFile[0].FullName -Raw -Encoding utf8 | ConvertFrom-Json

    $currDeleted = @($curr.changedTargets.deletedPathCandidates | ForEach-Object { "$_" })
    $prevDeleted = @($prev.changedTargets.deletedPathCandidates | ForEach-Object { "$_" })
    $addedDeleted = @($currDeleted | Where-Object { $prevDeleted -notcontains $_ } | Sort-Object -Unique)
    $removedDeleted = @($prevDeleted | Where-Object { $currDeleted -notcontains $_ } | Sort-Object -Unique)

    $currBlocked = @($curr.changedTargets.blockedPaths | ForEach-Object { "$_" })
    $prevBlocked = @($prev.changedTargets.blockedPaths | ForEach-Object { "$_" })
    $addedBlocked = @($currBlocked | Where-Object { $prevBlocked -notcontains $_ } | Sort-Object -Unique)
    $removedBlocked = @($prevBlocked | Where-Object { $currBlocked -notcontains $_ } | Sort-Object -Unique)
    $scoreDeltaText = $null
    if ($curr.ai -and $curr.ai.PSObject.Properties["evaluation"] -and $prev.ai -and $prev.ai.PSObject.Properties["evaluation"]) {
        $scoreDeltaText = "$($prev.ai.evaluation) -> $($curr.ai.evaluation)"
    }

    $diff = [PSCustomObject]@{
        schemaVersion = "1.0"
        generatedAt = (Get-Date).ToString("s")
        currentAudit = $currentFile.FullName
        previousAudit = $previousFile[0].FullName
        summaryDelta = [PSCustomObject]@{
            success = ([int]$curr.summary.success - [int]$prev.summary.success)
            ng = ([int]$curr.summary.ng - [int]$prev.summary.ng)
            skipped = ([int]$curr.summary.skipped - [int]$prev.summary.skipped)
            score = $scoreDeltaText
        }
        aiDelta = [PSCustomObject]@{
            source = "$($prev.ai.source) -> $($curr.ai.source)"
            promptVersion = "$($prev.ai.promptVersion) -> $($curr.ai.promptVersion)"
            fallbackReason = "$($prev.ai.fallbackReason) -> $($curr.ai.fallbackReason)"
        }
        changedTargetsDelta = [PSCustomObject]@{
            addedDeletedPathCandidates = $addedDeleted
            removedDeletedPathCandidates = $removedDeleted
            addedBlockedPaths = $addedBlocked
            removedBlockedPaths = $removedBlocked
        }
    }

    $stamp = Get-Date -Format "yyyyMMddHHmmss"
    $diffPath = Join-Path $reportsDir ("Audit_Diff_{0}_{1}.json" -f $script:RunId, $stamp)
    $diff | ConvertTo-Json -Depth 12 | Set-Content -Path $diffPath -Encoding utf8
    return $diffPath
}

Try-Step "Microsoft Store キャッシュのクリア" {
    if (Get-Command Invoke-CleanupMaintenance -ErrorAction SilentlyContinue) {
        Invoke-CleanupMaintenance -WhatIfMode:$script:IsWhatIfMode -Tasks @('store') | Out-Null
    } else {
        @(
            "$env:LOCALAPPDATA\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe\LocalCache",
            "$env:LOCALAPPDATA\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe\LocalState\Cache"
        ) | ForEach-Object { Clear-DirContents $_ }
    }
}

Try-Step "ごみ箱を空にする" {
    Add-DeletedPathCandidate '$Recycle.Bin'
    Start-Process -FilePath "rundll32.exe" -ArgumentList "shell32.dll,SHEmptyRecycleBinA 0" -NoNewWindow -Wait
}

Try-Step "DNS キャッシュのクリア" {
    if (Get-Command Invoke-CleanupMaintenance -ErrorAction SilentlyContinue) {
        Invoke-CleanupMaintenance -WhatIfMode:$script:IsWhatIfMode -Tasks @('dns') | Out-Null
    } else {
        & ipconfig /flushdns | Out-Null
    }
}

Try-Step "Windows イベントログのクリア" {
    foreach ($logName in @("Application", "System")) {
        & wevtutil cl $logName 2>&1 | Out-Null
        if ($LASTEXITCODE -eq 0) {
            Write-Log "[イベントログ] ${logName}: クリア完了"
        } else {
            Write-Log "[イベントログ] ${logName}: クリアをスキップ（使用中または権限不足）"
        }
    }
}

Try-Step "ディスクの最適化（SSD: TRIM / HDD: デフラグ）" {
    $mediaType = "HDD"   # 安全なデフォルト値
    if (Get-Command -Name Get-PhysicalDisk -ErrorAction SilentlyContinue) {
        $pDisk = Get-PhysicalDisk | Where-Object { $_.DeviceID -eq 0 }
        if ($pDisk -and $pDisk.MediaType) { $mediaType = $pDisk.MediaType }
    }
    if ($mediaType -eq "SSD") {
        Optimize-Volume -DriveLetter C -ReTrim -ErrorAction SilentlyContinue
    } else {
        & defrag C: /U /V | Out-Null
        if ($LASTEXITCODE -ne 0) { Write-Log "[警告] デフラグが失敗しました（exitcode=$LASTEXITCODE）。" }
    }
}

Try-Step "SSD ヘルスチェック（SMART）" {
    if (Get-Command -Name Get-PhysicalDisk -ErrorAction SilentlyContinue) {
        $smart = Get-PhysicalDisk | Where-Object { $_.MediaType -eq "SSD" }
        if ($smart) {
            foreach ($disk in $smart) {
                $health = $disk.OperationalStatus
                $model  = $disk.FriendlyName
                $serial = $disk.SerialNumber
                $msg    = "SSD [${model}] シリアル:[${serial}] 状態: $health"
                Show $msg Yellow
                Write-Log "[SSD ヘルス] $msg"
            }
        } else {
            $msg = "SSD が検出されませんでした。"
            Show $msg DarkGray
            Write-Log "[SSD ヘルス] $msg"
        }
    } else {
        $msg = "Get-PhysicalDisk が利用できません（Storage モジュール未インストール）。SSD チェックをスキップします。"
        Show $msg DarkGray
        Write-Log "[SSD ヘルス] $msg"
    }
}

Try-Step "システムファイルの整合性チェック・修復（SFC）" {
    if ($isPS7Plus) {
        Show "🔍  ${B}sfc /scannow${RB} を起動します（15〜60 分かかる場合があります）" Yellow
    } else {
        Show "* sfc /scannow を起動します（15〜60 分かかる場合があります）" Yellow
    }
    Show "  SFC ウィンドウの「検証が XX% 完了しました」が進捗です。" Gray
    Show "  ──────────────────────────────────────────" Gray
    Write-Log "[SFC] sfc /scannow 開始"

    # SFC を別プロセスで起動し、本ウィンドウで経過時間と CBS.log をリアルタイム表示
    $sfcProc = Start-Process -FilePath "$env:SystemRoot\System32\sfc.exe" `
        -ArgumentList "/scannow" -PassThru
    $cbsLog  = "$env:SystemRoot\Logs\CBS\CBS.log"
    $start   = Get-Date
    $spinArr = @('|', '/', '-', '\')
    $spinIdx = 0

    while (-not $sfcProc.HasExited) {
        Start-Sleep -Seconds 3
        $elapsed = (Get-Date) - $start
        $ch      = $spinArr[$spinIdx % 4]; $spinIdx++
        $act     = ""
        if (Test-Path $cbsLog) {
            $line = try { (Get-Content $cbsLog -Tail 1 -ErrorAction Stop) } catch { "" }
            if ($line -match 'CBS\s+(.{5,60})$') {
                $act = $Matches[1].Trim()
                if ($act.Length -gt 55) { $act = $act.Substring(0, 55) + "..." }
            }
        }
        Write-Host ("`r  [{0}] 経過: {1}分{2:D2}秒   {3,-58}" -f `
            $ch, [int]$elapsed.TotalMinutes, $elapsed.Seconds, $act) `
            -NoNewline -ForegroundColor Cyan
    }
    Write-Host ""  # 改行確定

    $exitCode = $sfcProc.ExitCode
    if ($exitCode -eq 0) {
        if ($isPS7Plus) {
            Show "  ✅ SFC 完了（終了コード: 0）" Green
        } else {
            Show "  [完了] SFC 完了（終了コード: 0 = 問題なし/修復済み）" Green
        }
    } else {
        Show ("  [警告] SFC 終了コード: {0}  詳細: {1}" -f $exitCode, $cbsLog) Yellow
    }
    Write-Log "[SFC] sfc /scannow 完了（終了コード: $exitCode）"
}

Try-Step "Windows コンポーネントストアの診断（DISM）" {
    if ($isPS7Plus) {
        Show "🛠️  ${B}DISM /ScanHealth${RB} を起動します（5〜20 分かかる場合があります）" Yellow
    } else {
        Show "* DISM /ScanHealth を起動します（5〜20 分かかる場合があります）" Yellow
    }
    Write-Log "[DISM] /ScanHealth 開始"

    $dismProc = Start-Process -FilePath "$env:SystemRoot\System32\Dism.exe" `
        -ArgumentList "/Online", "/Cleanup-Image", "/ScanHealth" -PassThru
    $start   = Get-Date
    $spinArr = @('|', '/', '-', '\')
    $spinIdx = 0

    while (-not $dismProc.HasExited) {
        Start-Sleep -Seconds 3
        $elapsed = (Get-Date) - $start
        $ch      = $spinArr[$spinIdx % 4]; $spinIdx++
        Write-Host ("`r  [{0}] 経過: {1}分{2:D2}秒" -f `
            $ch, [int]$elapsed.TotalMinutes, $elapsed.Seconds) `
            -NoNewline -ForegroundColor Cyan
    }
    Write-Host ""  # 改行確定

    $exitCode = $dismProc.ExitCode
    if ($exitCode -eq 0) {
        if ($isPS7Plus) {
            Show "  ✅ DISM 完了（終了コード: 0 = 正常）" Green
        } else {
            Show "  [完了] DISM 完了（終了コード: 0 = 正常）" Green
        }
    } else {
        Show ("  [警告] DISM 終了コード: {0}" -f $exitCode) Yellow
    }
    Write-Log "[DISM] /ScanHealth 完了（終了コード: $exitCode）"
}

Try-Step "電源プランの最適化" {
    $battery  = Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue
    $isLaptop = ($null -ne $battery)
    $balanced = "381b4222-f694-41f0-9685-ff5bb260df2e"
    $highPerf = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"

    if ($isLaptop) {
        & powercfg /setactive $balanced 2>&1 | Out-Null
        $powercfgExit = $LASTEXITCODE
        $planName   = "バランス"
        $deviceType = "ノートPC"
    } else {
        # 高パフォーマンスプランが存在しない場合は有効化
        $exists = (& powercfg /list) | Select-String -SimpleMatch $highPerf -Quiet
        if (-not $exists) {
            & powercfg /duplicatescheme $highPerf 2>&1 | Out-Null
        }
        & powercfg /setactive $highPerf 2>&1 | Out-Null
        $powercfgExit = $LASTEXITCODE
        $planName   = "高パフォーマンス"
        $deviceType = "デスクトップ PC"
    }

    if ($powercfgExit -eq 0) {
        $msg = "電源プランを「${planName}」に設定しました（${deviceType} 判定）。"
        Show $msg Cyan
    } else {
        $msg = "電源プランの設定に失敗しました（このシステムではプランが利用できない可能性があります）。"
        Show $msg Yellow
    }
    Write-Log "[電源プラン] $msg"
}

# ==========================

# Microsoft 365 の更新確認・適用
Try-Step "Microsoft 365 の更新確認・適用" {
    # Click-to-Run クライアントのパスを検索（x64 / x86 両対応）
    $c2rExe = @(
        "$env:ProgramFiles\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe",
        "${env:ProgramFiles(x86)}\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1

    if (-not $c2rExe) {
        $msg = "Microsoft 365 (Click-to-Run) が検出されませんでした。スキップします。"
        Show $msg DarkGray
        Write-Log "[Microsoft 365] $msg"
        return
    }

    # レジストリから現在のバージョンを取得
    $verMsg = "バージョン情報を取得できませんでした"
    try {
        $c2rConfig = Get-ItemProperty `
            -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" `
            -ErrorAction SilentlyContinue
        if ($c2rConfig -and $c2rConfig.ClientVersionToReport) {
            $verMsg = "現在のバージョン: $($c2rConfig.ClientVersionToReport)"
        }
    } catch {
        Write-Host "  M365 バージョン情報の取得に失敗しました（スキップ）" -ForegroundColor DarkGray
    }

    Show "${B}Microsoft 365 を検出しました。${R}" Cyan
    Show "  $verMsg" White
    Write-Log "[Microsoft 365] 検出: $c2rExe / $verMsg"

    # ユーザー確認
    $choice = Get-UserChoice -Prompt "${B}Microsoft 365 の更新を確認・適用しますか？${RB} (Y/N)" -Default 'N'
    Write-Log "[Microsoft 365] ユーザー選択: $choice"
    if (Test-ChoiceYes -Choice $choice) {
        Show "Microsoft 365 の更新を開始します（更新 UI が表示される場合があります）..." Cyan
        Write-Log "[Microsoft 365] ユーザーが更新を承認。OfficeC2RClient で更新開始。"
        # displaylevel=True で更新 UI を表示し、forceappshutdown=false で Office を強制終了しない
        Start-Process -FilePath $c2rExe `
            -ArgumentList "/update user displaylevel=True forceappshutdown=false" -Wait
        Write-Log "[Microsoft 365] 更新コマンドを実行しました（バックグラウンドで処理されます）。"
        Show "Microsoft 365 の更新を開始しました。" Green
    } else {
        $msg = "Microsoft 365 の更新をスキップしました。"
        Show $msg Yellow
        Write-Log "[Microsoft 365] $msg"
    }
}

# 最後に Windows Update を実行
Try-Step "Windows Update の実行" {
    Run-WindowsUpdate
}

# ── タスク 20: スタートアップ・サービスレポート ──────────────────────
Try-Step "スタートアップ・サービスレポート" {
    # ── フェーズ 1: スタートアップ登録アプリ一覧 ──────────────────
    if ($isPS7Plus) {
        Show "📋  ${B}スタートアップ登録アプリ一覧${RB}" Cyan
    } else {
        Show "* スタートアップ登録アプリ一覧" Cyan
    }
    Write-Log "---- [タスク20] スタートアップ登録アプリ一覧 ----"

    try {
        if (Get-Command Get-StartupAnalysis -ErrorAction SilentlyContinue) {
            $startupModule = Get-StartupAnalysis
            $startupApps = @($startupModule.Items | Sort-Object Name)
            Write-Log ("  [module] StartupCount={0}, Rating={1}" -f $startupModule.StartupCount, $startupModule.Rating)
        } else {
            $startupApps = Get-CimInstance -ClassName Win32_StartupCommand -ErrorAction Stop |
                Select-Object Name, Command, Location, User |
                Sort-Object Location, Name
        }

        if (-not $startupApps) {
            Show "  登録なし（スタートアップアプリは見つかりませんでした）" Gray
            Write-Log "  登録なし"
        } else {
            # 表示上限 30 件（ログには全件記録）
            $displayLimit  = 30
            $displayCount  = 0
            $totalCount    = @($startupApps).Count

            foreach ($app in $startupApps) {
                Write-Log ("  名前: {0} | コマンド: {1} | 場所: {2} | ユーザー: {3}" -f `
                    $app.Name, $app.Command, $app.Location, $app.User)

                if ($displayCount -lt $displayLimit) {
                    Show ("  [{0}] {1}" -f $app.Location, $app.Name) White
                    $displayCount++
                }
            }

            if ($totalCount -gt $displayLimit) {
                Show ("  ... 他 {0} 件はログファイルを参照してください" -f ($totalCount - $displayLimit)) Gray
            }
            Write-Log ("  合計: {0} 件" -f $totalCount)
        }
    } catch {
        Show "  スタートアップ情報の取得に失敗しました（スキップ）" Yellow
        Write-ErrorLog "タスク20 スタートアップ取得失敗: $_"
    }

    Show "-----------------------------------------------------" Gray

    # ── フェーズ 2: 不要サービスチェック ──────────────────────────
    if ($isPS7Plus) {
        Show "🔍  ${B}不要サービス稼働チェック${RB}" Cyan
    } else {
        Show "* 不要サービス稼働チェック" Cyan
    }
    Write-Log "---- [タスク20] 不要サービス稼働チェック ----"

    # 照合リスト（読み取り専用チェックのみ。変更・停止は行わない）
    $unnecessaryServices = @(
        'Fax',         # FAXサービス（多くの環境で不使用）
        'RetailDemo',  # 小売デモモード（一般ユーザー不要）
        'XblGameSave', # Xbox ゲームセーブ（非ゲーマー不要）
        'XboxGipSvc',  # Xbox 周辺機器インターフェース
        'WpcMonSvc',   # 保護者機能モニタ（Windows 10 のみ）
        'DiagTrack',   # Connected User Experiences and Telemetry
        'MapsBroker',  # ダウンロード済み地図マネージャ
        'lfsvc'        # 位置情報サービス
    )

    try {
        $runningServices = @(Get-Service -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq 'Running' })

        $foundCount = 0
        foreach ($svcName in $unnecessaryServices) {
            $match = $runningServices | Where-Object { $_.Name -eq $svcName }
            if ($match) {
                $foundCount++
                Show ("  [確認推奨] {0} が実行中です" -f $svcName) Yellow
                Write-Log ("  [確認推奨] {0} ({1})" -f $svcName, $match.DisplayName)
            }
        }

        if ($foundCount -eq 0) {
            Show "  問題なし（確認推奨サービスの稼働は検出されませんでした）" Green
            Write-Log "  確認推奨サービス: 0 件"
        } else {
            Show ("  合計 {0} 件の確認推奨サービスが実行中です（変更は手動で行ってください）" -f $foundCount) Yellow
            Write-Log ("  確認推奨サービス: {0} 件" -f $foundCount)
        }
    } catch {
        Show "  サービス情報の取得に失敗しました（スキップ）" Yellow
        Write-ErrorLog "タスク20 サービスチェック失敗: $_"
    }
}

# ==========================
# 実行前後のディスク空き容量比較
# ==========================
$finalFreeGB = $initialFreeGB   # CIM 取得失敗時のフォールバック（解放量 0 GB 扱い）
try {
    $finalMaybe = Get-SystemDriveFreeGB -DriveLetter 'C:'
    if ($null -eq $finalMaybe) { throw "ディスク空き容量を取得できませんでした。" }
    $finalFreeGB = $finalMaybe
    $freedGB     = [math]::Round($finalFreeGB - $initialFreeGB, 2)
    $freedColor  = if ($freedGB -gt 0) { "Green" } else { "Yellow" }
    $freedStr    = if ($freedGB -gt 0) { "+${freedGB}" } else { "${freedGB}" }

    Show "-----------------------------------------------------" Gray
    if ($isPS7Plus) {
        Show "📊  ${B}ディスク解放結果:${R}" Cyan
        Show "    最適化前: ${initialFreeGB} GB  →  最適化後: ${finalFreeGB} GB" White
        Show "    ${B}解放容量:${RB} ${freedStr} GB" $freedColor
    } else {
        Show "* ディスク解放結果:" Cyan
        Show "  最適化前: ${initialFreeGB} GB -> 最適化後: ${finalFreeGB} GB" White
        Show "  解放容量: ${freedStr} GB" $freedColor
    }
    Write-Log "[ディスク解放] 最適化前: ${initialFreeGB} GB / 最適化後: ${finalFreeGB} GB / 解放: ${freedStr} GB"
} catch {
    Write-ErrorLog "ディスク空き容量の比較に失敗しました: $_"
}

# ── モジュール診断の統合実行 ─────────────────────────────────────────
if (Get-Command Invoke-IntegratedModuleDiagnostic -ErrorAction SilentlyContinue) {
    Invoke-IntegratedModuleDiagnostic
}
if ($EnableAIDiagnosis -and (Get-Command Invoke-AIDiagnosis -ErrorAction SilentlyContinue)) {
    $apiKeySource = "none"
    $apiKey = ""
    if (-not [string]::IsNullOrWhiteSpace($AnthropicApiKey)) {
        $apiKey = $AnthropicApiKey
        $apiKeySource = "param"
    } else {
        $envKey = [Environment]::GetEnvironmentVariable("ANTHROPIC_API_KEY", "Process")
        if (-not [string]::IsNullOrWhiteSpace($envKey)) {
            $apiKey = $envKey
            $apiKeySource = "env"
        }
    }

    $anthropicRequested = [bool]$UseAnthropicAI
    if (-not $anthropicRequested -and -not [string]::IsNullOrWhiteSpace($apiKey)) {
        $anthropicRequested = $true
        Write-Log "[AI] Anthropic自動有効化: APIキーが存在するため -UseAnthropicAI 未指定でも実行します。"
    }

    if ($anthropicRequested) {
        if ([string]::IsNullOrWhiteSpace($apiKey)) {
            Write-Log "[AI] Anthropic APIキー未検出: source=$apiKeySource"
        } else {
            Write-Log "[AI] Anthropic APIキー検出: source=$apiKeySource length=$($apiKey.Length)"
        }
    }
    $script:AIDiagnosis = Invoke-AIDiagnosis `
        -HealthScore $script:HealthScore `
        -Snapshot $script:ModuleSnapshot `
        -UpdateClassifiedErrors $script:UpdateErrorClassification `
        -M365Connectivity $script:M365Connectivity `
        -EventAnomaly $script:EventAnomaly `
        -BootTrend $script:BootTrend `
        -AnthropicApiKey $(if ($anthropicRequested) { $apiKey } else { "" }) `
        -AnthropicModel $AnthropicModel
    if ($script:AIDiagnosis) {
        Write-Log "[AI] source=$($script:AIDiagnosis.Source) eval=$($script:AIDiagnosis.Evaluation)"
        if ($script:AIDiagnosis.PSObject.Properties["FallbackReason"] -and $script:AIDiagnosis.FallbackReason) {
            Write-Log "[AI] fallback=$($script:AIDiagnosis.FallbackReason)"
            if (Get-Command Invoke-AgentHookEvent -ErrorAction SilentlyContinue) {
                $hookContext = [PSCustomObject]@{ runId = $script:RunId; source = $script:AIDiagnosis.Source; fallbackReason = $script:AIDiagnosis.FallbackReason }
                $script:HookHistory += @(Invoke-AgentHookEvent -EventName "on_fallback" -Context $hookContext -HooksConfig $(if ($script:Config) { $script:Config.hooks } else { $null }) -RunId $script:RunId -LogsDir $logsDir)
            }
        }
    }
}
if ($script:HealthScore) {
    if (-not $script:sysInfo) { $script:sysInfo = @{} }
    $statusJa = Convert-HealthStatusToJapanese -Status "$($script:HealthScore.Status)"
    $script:sysInfo['HealthScore'] = "$($script:HealthScore.Score) / 100 ($statusJa)"
}
if ($script:AIDiagnosis) {
    if (-not $script:sysInfo) { $script:sysInfo = @{} }
    $script:sysInfo['AIEvaluation'] = "$($script:AIDiagnosis.Evaluation)"
    $script:sysInfo['AIHeadline'] = "$($script:AIDiagnosis.Headline)"
    if ($script:AIDiagnosis.PSObject.Properties["PromptVersion"]) {
        $script:sysInfo['AIPromptVersion'] = "$($script:AIDiagnosis.PromptVersion)"
    }
}
if ($script:ExecutionProfile -eq 'agent-teams') {
    try {
        Write-Log "[agent-teams] 実行開始"
        $teamResult = Invoke-AgentTeamsOrchestration -RunId $script:RunId -ReportsDir $reportsDir -ModuleSnapshot $script:ModuleSnapshot -HealthScore $script:HealthScore -AIDiagnosis $script:AIDiagnosis -HooksConfig $(if ($script:Config) { $script:Config.hooks } else { $null }) -McpProviders $(if ($script:Config) { $script:Config.mcpProviders } else { $null }) -AgentTeamsConfig $(if ($script:Config) { $script:Config.agentTeams } else { $null }) -LogsDir $logsDir -RunMode $script:RunMode
        $script:AgentTeamsSummary = $teamResult
        $script:AgentTeamsResult = $teamResult.summary
        $script:McpResults = @($teamResult.mcpResults)
        if ($teamResult.PSObject.Properties["hookHistory"] -and $teamResult.hookHistory) {
            $script:HookHistory += @($teamResult.hookHistory)
        }
        Write-Log "[agent-teams] plan=$($teamResult.planPath)"
        Write-Log "[agent-teams] summary=$($teamResult.summaryPath)"
        if (-not $script:sysInfo) { $script:sysInfo = @{} }
        $script:sysInfo['ExecutionProfile'] = $script:ExecutionProfile
        if ($script:AgentTeamsResult -and $script:AgentTeamsResult.analyzer) {
            $script:sysInfo['AgentOverall'] = "$($script:AgentTeamsResult.analyzer.overall)"
        }
    } catch {
        Write-ErrorLog "[agent-teams] 実行失敗: $_"
    }
}

# ── HTMLレポート生成 ────────────────────────────────────────────────
New-HtmlReport -Results    $script:taskResults `
               -DiskBefore $initialFreeGB `
               -DiskAfter  $finalFreeGB `
               -SysInfo    $script:sysInfo `
               -AIDiagnosis $script:AIDiagnosis
$executionReportPath = Export-JsonExecutionReport
$latestHtmlReportPath = Get-ChildItem -Path $logsDir -Filter "PC_Optimizer_Report_*.html" -File -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTimeUtc |
    Select-Object -Last 1 -ExpandProperty FullName
Write-UiEvent -Type "report_generated" -Data @{
    htmlReportPath      = $latestHtmlReportPath
    executionReportPath = $executionReportPath
}
if (Get-Command Export-DeletedPathCandidate -ErrorAction SilentlyContinue) {
    Export-DeletedPathCandidate | Out-Null
}
if (Get-Command Export-IntegratedModuleReport -ErrorAction SilentlyContinue) {
    Export-IntegratedModuleReport
}
if (Get-Command Invoke-AgentHookEvent -ErrorAction SilentlyContinue) {
    $hookContext = [PSCustomObject]@{ runId = $script:RunId; reportType = "standard"; executionProfile = $script:ExecutionProfile }
    $script:HookHistory += @(Invoke-AgentHookEvent -EventName "on_report" -Context $hookContext -HooksConfig $(if ($script:Config) { $script:Config.hooks } else { $null }) -RunId $script:RunId -LogsDir $logsDir)
}
if ($ExportPowerBIJson -or (Get-Command Export-PowerBIDashboardJson -ErrorAction SilentlyContinue)) {
    if (Get-Command Export-PowerBIDashboardJson -ErrorAction SilentlyContinue) {
        $pbiPath = Join-Path $reportsDir "PowerBI_Dashboard_latest.json"
        Export-PowerBIDashboardJson -Path $pbiPath `
            -HealthScore $script:HealthScore `
            -Snapshot $script:ModuleSnapshot `
            -AIDiagnosis $script:AIDiagnosis `
            -EventAnomaly $script:EventAnomaly `
            -BootTrend $script:BootTrend `
            -AgentSummary $script:AgentTeamsResult `
            -TaskResults $script:taskResults `
            -DiskBefore $initialFreeGB `
            -DiskAfter $finalFreeGB `
            -HookHistory $script:HookHistory `
            -McpResults $script:McpResults | Out-Null
        Write-Log "[powerbi] 保存完了: $pbiPath"
    }
}
$auditPath = Export-RunAuditJson
Write-Log "[audit] 保存完了: $auditPath"
$auditDiffPath = Export-AuditDiffJson -CurrentAuditPath $auditPath
if ($auditDiffPath) {
    Write-Log "[audit-diff] 保存完了: $auditDiffPath"
}
if ($script:GuardrailState -and (Get-Command Complete-RepairGuardrails -ErrorAction SilentlyContinue)) {
    $manifestPath = Complete-RepairGuardrails -State $script:GuardrailState -LogsDir $logsDir
    Write-Log "[guardrail] 完了: $manifestPath"
}

# 再起動が必要な場合はプロンプト表示
if (Test-PendingReboot) {
    Write-Log "[再起動] 再起動が必要です。"
    if ($script:IsNoRebootPrompt) {
        $choice = 'N'
        Write-Log "[再起動] -NoRebootPrompt により確認を抑止し再起動をスキップしました。"
    } else {
        $choice = Get-UserChoice -Prompt "${B}再起動が必要です。今すぐ再起動しますか？${RB} (Y/N)" -Default 'N'
    }
    if (Test-ChoiceYes -Choice $choice) {
        Write-Log "[再起動] ユーザーが再起動を承認しました。"
        Restart-Computer -Force
    } else {
        Write-Log "[再起動] ユーザーが再起動を拒否しました。"
    }
}

# 完了ログ・出力
Write-Log "[全タスク完了] $(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')"
if ($isPS7Plus) {
    Show "${B}✅ PC 最適化が完了しました。${R}" Green
} else {
    Show "PC 最適化が完了しました。" Green
}
Write-Host ""

# 終了前に待機
if ($script:IsNonInteractive) {
    Write-Log "[NonInteractive] 終了待機をスキップしました。"
} else {
    Read-Host "Enter キーを押して終了"
}

# 終了コードの標準化
if ($script:ExitCode -eq 0) {
    if ($script:HadTaskFailure) {
        if ($script:FailureMode -eq 'fail-fast') {
            $script:ExitCode = $script:ExitCodes.Fatal
        } else {
            $script:ExitCode = $script:ExitCodes.Partial
        }
    } else {
        $script:ExitCode = $script:ExitCodes.Success
    }
}
Write-UiEvent -Type "run_complete" -Data @{
    exitCode = $script:ExitCode
    status   = Get-RunStatus
}
exit $script:ExitCode




