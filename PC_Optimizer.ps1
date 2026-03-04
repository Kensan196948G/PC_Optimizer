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
    [string]$Tasks = "all",
    [string]$FailureMode = "continue",
    [string]$ExportDeletedPaths = ""
)

# ログパスの構築
$now     = Get-Date -Format "yyyyMMddHHmm"
$logsDir = Join-Path -Path $PSScriptRoot -ChildPath "logs"
if (-not (Test-Path $logsDir)) {
    New-Item -ItemType Directory -Path $logsDir -Force | Out-Null
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
$script:ExportDeletedFmt  = $ExportDeletedPaths.ToLowerInvariant()
$script:HadTaskFailure    = $false
$script:FatalStopRequested= $false
$script:ExitCode          = 0
$script:RunId             = ([guid]::NewGuid().ToString("N"))
$script:DeletedPathSet    = New-Object 'System.Collections.Generic.HashSet[string]'
$script:DeletedPathList   = New-Object 'System.Collections.Generic.List[string]'

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

# v4.0 foundation: 共通モジュールと設定を接続
try {
    $commonModulePath = Join-Path $PSScriptRoot "modules\Common.psm1"
    $reportModulePath = Join-Path $PSScriptRoot "modules\Report.psm1"
    if (Test-Path $commonModulePath) {
        Import-Module $commonModulePath -Force -ErrorAction Stop
        if (Get-Command Get-OptimizerConfig -ErrorAction SilentlyContinue) {
            $script:Config = Get-OptimizerConfig -Path $ConfigPath
            Write-Log "[config] 読込成功: $ConfigPath"
        }
        if (Test-Path $reportModulePath) {
            Import-Module $reportModulePath -Force -ErrorAction Stop
        }
    } else {
        Write-Log "[config] modules\\Common.psm1 が見つからないため既定設定で継続"
    }
} catch {
    Write-ErrorLog "共通モジュールまたは config 読込に失敗しました: $_"
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

function Initialize-ExecutionOptions {
    if ($script:FailureMode -notin @('continue', 'fail-fast')) {
        Show "不正な -FailureMode です: $FailureMode" Red
        $script:ExitCode = $script:ExitCodes.InvalidArgs
        exit $script:ExitCode
    }
    if ($script:ExportDeletedFmt -and $script:ExportDeletedFmt -notin @('csv', 'json')) {
        Show "不正な -ExportDeletedPaths です: $ExportDeletedPaths" Red
        $script:ExitCode = $script:ExitCodes.InvalidArgs
        exit $script:ExitCode
    }

    if ($Tasks -and $Tasks -ne 'all') {
        $set = New-Object 'System.Collections.Generic.HashSet[int]'
        foreach ($token in ($Tasks -split ',')) {
            $trimmed = $token.Trim()
            if ($trimmed -notmatch '^\d+$') {
                Show "不正な -Tasks 指定です: $Tasks" Red
                $script:ExitCode = $script:ExitCodes.InvalidArgs
                exit $script:ExitCode
            }
            $id = [int]$trimmed
            if ($id -lt 1 -or $id -gt 20) {
                Show "-Tasks は 1-20 の範囲で指定してください: $Tasks" Red
                $script:ExitCode = $script:ExitCodes.InvalidArgs
                exit $script:ExitCode
            }
            [void]$set.Add($id)
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

function Register-TaskPlannedDeletedPaths {
    param([int]$TaskId)
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

function Export-DeletedPathCandidates {
    if (-not $script:ExportDeletedFmt) { return $null }
    $stamp = Get-Date -Format "yyyyMMddHHmmss"
    $sorted = @($script:DeletedPathList | Sort-Object -Unique)
    if ($script:ExportDeletedFmt -eq 'json') {
        $path = Join-Path $logsDir "DeletedPaths_${stamp}.json"
        [PSCustomObject]@{
            runId = $script:RunId
            mode  = if ($script:IsWhatIfMode) { 'whatif' } else { 'execute' }
            tasks = $sorted
        } | ConvertTo-Json -Depth 4 | Set-Content -Path $path -Encoding $logEncoding
        Write-Log "[DeletedPaths] JSON 出力: $path"
        return $path
    }
    $path = Join-Path $logsDir "DeletedPaths_${stamp}.csv"
    $sorted | ForEach-Object { [PSCustomObject]@{ path = $_ } } |
        Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
    Write-Log "[DeletedPaths] CSV 出力: $path"
    return $path
}

Initialize-ExecutionOptions
if (-not (Test-IsAdministrator) -and -not ($script:IsWhatIfMode -and $script:IsNonInteractive)) {
    Show "管理者権限で実行してください。" Red
    $script:ExitCode = $script:ExitCodes.Permission
    exit $script:ExitCode
}

# ハードウェア・システム情報の収集（24H2 / 22H2 等のマーケティングバージョンを含む）
try {
    $hostname = $env:COMPUTERNAME
    $username = $env:USERNAME

    # OS 名とマーケティングバージョン（24H2/22H2 など）
    $osinfo = Get-CimInstance Win32_OperatingSystem
    $osName = $osinfo.Caption
    $osVer  = $osinfo.Version
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

    # その他の情報
    $cpuinfo    = Get-CimInstance Win32_Processor
    $cpu        = $cpuinfo.Name
    $cpuCores   = $cpuinfo.NumberOfCores
    $cpuThreads = $cpuinfo.NumberOfLogicalProcessors
    $mem        = [Math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 2)
    $disk       = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
    $free       = [math]::Round($disk.FreeSpace / 1GB, 1)
    $total      = [math]::Round($disk.Size / 1GB, 1)
    $pwv        = $PSVersionTable.PSVersion.ToString()

    # コンソール出力 — PS7+ は絵文字、PS5.x 以前は ASCII 記号
    if ($isPS7Plus) {
        Show "🖥️  ${B}ホスト:${RB} $hostname" Green
        Show "👤  ${B}ユーザー:${RB} $username" Cyan
        Show "🏷️  ${B}OS:${RB} $os" Yellow
        Show "💻  ${B}CPU:${RB} $cpu  $cpuCores コア / $cpuThreads スレッド" Magenta
        Show "🧠  ${B}メモリ:${RB} ${mem} GB" Blue
        Show "💾  ${B}ディスク $($disk.DeviceID) 空き:${RB} ${free}GB / ${total}GB" White
    } else {
        Show "* ホスト: $hostname" Green
        Show "* ユーザー: $username" Cyan
        Show "* OS: $os" Yellow
        Show "* CPU: $cpu  $cpuCores コア / $cpuThreads スレッド" Magenta
        Show "* メモリ: ${mem} GB" Blue
        Show "* ディスク $($disk.DeviceID) 空き: ${free}GB / ${total}GB" White
    }
    Show "-----------------------------------------------------" Gray

    Write-Log "===== 高性能 PC 最適化ツール 実行ログ ====="
    Write-Log "[開始] $(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')"
    Write-Log "[ホスト] $hostname"
    Write-Log "[ユーザー] $username"
    Write-Log "[OS] $os"
    Write-Log "[CPU] $cpu  $cpuCores コア / $cpuThreads スレッド"
    Write-Log "[メモリ] $mem GB"
    Write-Log "[ディスク] $($disk.DeviceID) 空き: $free GB / $total GB"
    Write-Log "[PowerShell バージョン] $pwv"
    $script:sysInfo = @{
        Hostname  = $hostname
        Username  = $username
        OS        = $os
        CPU       = "$cpu  $cpuCores コア / $cpuThreads スレッド"
        RAM       = "${mem} GB"
        Disk      = "C: 空き ${free}GB / ${total}GB"
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
    if ($script:IsWhatIfMode) { return }
    if (-not (Test-Path $Path)) { return }
    & cmd.exe /c "rd /s /q `"$Path`"" 2>$null
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
    if ($script:IsWhatIfMode -and -not (Test-ReadOnlyTask -TaskName $desc)) {
        $sw.Stop()
        $msg = "[WhatIf] $desc はプレビュー実行のため変更をスキップしました。"
        Show $msg Yellow
        Write-Log $msg
        Register-TaskPlannedDeletedPaths -TaskId $taskId
        $script:taskResults += [PSCustomObject]@{
            Id         = $taskId
            Name       = $desc
            Status     = "SKIP"
            Duration   = [int]$sw.Elapsed.TotalSeconds
            Error      = "WhatIf skip"
            Errors     = @("WhatIf skip")
            PreviewOnly= $true
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
    }
}

# ── HTML レポート生成 ─────────────────────────────────────────────────
function New-HtmlReport {
    param(
        [PSCustomObject[]]$Results,
        [double]$DiskBefore,
        [double]$DiskAfter,
        [hashtable]$SysInfo
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
            $icon    = if ($_.Status -eq "OK") { "&#x2705;" } else { "&#x274C;" }
            $rowCls  = if ($_.Status -eq "OK") { "row-ok" } else { "row-ng" }
            $nameEsc = [System.Web.HttpUtility]::HtmlEncode($_.Name)
            $errEsc  = [System.Web.HttpUtility]::HtmlEncode($_.Error)
            $errCell = if ($errEsc) { "<span class='err-detail'>$errEsc</span>" } else { "" }
            "<tr class='$rowCls'><td>$icon</td><td>$nameEsc</td><td>$($_.Duration)s</td><td>$errCell</td></tr>"
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

        # OK / NG カウント
        $okCount = ($Results | Where-Object { $_.Status -eq "OK" }).Count
        $ngCount = ($Results | Where-Object { $_.Status -eq "NG" }).Count

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
.row-ng{background:rgba(211,47,47,.05)}
.err-detail{color:#d97706;font-size:.8rem}
.sys-key{color:var(--sub);width:120px}
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
    <div class="summary-item"><div class="val" style="color:$diskColor">$diskSign GB</div><div class="lbl">ディスク解放量</div></div>
  </div>
</div>

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

<div class="card">
  <h2>&#x2705; タスク実行結果</h2>
  <table>
    <tr><th></th><th>タスク名</th><th>所要時間</th><th>エラー詳細</th></tr>
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

    if ($updateCount -eq 0 -and $updates -ne $null) {
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
    $initialFreeGB = [math]::Round(
        (Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'").FreeSpace / 1GB, 2)
} catch {
    Write-Host "  初期ディスク空き容量の取得に失敗しました（スキップ）" -ForegroundColor DarkGray
}

# ==========================
# メイン最適化・クリーンアップ
# ==========================
Try-Step "一時ファイルの削除" {
    Clear-DirContents "$env:SystemRoot\Temp"
    Clear-DirContents $env:TEMP
    Clear-DirContents "$env:USERPROFILE\AppData\Local\Temp"
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

    $jsonObject = [PSCustomObject]@{
        version         = "1.0"
        runId           = $script:RunId
        startedAt       = $script:scriptStartTime.ToUniversalTime().ToString("o")
        finishedAt      = $finishedAt.ToUniversalTime().ToString("o")
        host            = $hostInfo
        status          = Get-RunStatus
        durationSeconds = [math]::Round(($finishedAt - $script:scriptStartTime).TotalSeconds, 3)
        tasks           = $tasks
    }

    $reportData = New-OptimizerReportData -InputObject $jsonObject
    $jsonPath = Join-Path $logsDir ("PC_Optimizer_Report_{0}.json" -f (Get-Date -Format "yyyyMMddHHmmss"))
    Export-OptimizerReport -ReportData $reportData.Data -Format json -Path $jsonPath | Out-Null
    Write-Log "[JSONレポート] 保存完了: $jsonPath"
    return $jsonPath
}

Try-Step "Microsoft Store キャッシュのクリア" {
    @(
        "$env:LOCALAPPDATA\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe\LocalCache",
        "$env:LOCALAPPDATA\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe\LocalState\Cache"
    ) | ForEach-Object { Clear-DirContents $_ }
}

Try-Step "ごみ箱を空にする" {
    Add-DeletedPathCandidate '$Recycle.Bin'
    Start-Process -FilePath "rundll32.exe" -ArgumentList "shell32.dll,SHEmptyRecycleBinA 0" -NoNewWindow -Wait
}

Try-Step "DNS キャッシュのクリア" {
    & ipconfig /flushdns | Out-Null
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
        $planName   = "バランス"
        $deviceType = "ノートPC"
    } else {
        # 高パフォーマンスプランが存在しない場合は有効化
        $exists = (& powercfg /list) | Select-String -SimpleMatch $highPerf -Quiet
        if (-not $exists) {
            & powercfg /duplicatescheme $highPerf 2>&1 | Out-Null
        }
        & powercfg /setactive $highPerf 2>&1 | Out-Null
        $planName   = "高パフォーマンス"
        $deviceType = "デスクトップ PC"
    }

    if ($LASTEXITCODE -eq 0) {
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
        $startupApps = Get-CimInstance -ClassName Win32_StartupCommand -ErrorAction Stop |
            Select-Object Name, Command, Location, User |
            Sort-Object Location, Name

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
        $runningServices = Get-Service -ErrorAction Stop |
            Where-Object { $_.Status -eq 'Running' }

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
    $finalFreeGB = [math]::Round(
        (Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'").FreeSpace / 1GB, 2)
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

# ── HTMLレポート生成 ────────────────────────────────────────────────
New-HtmlReport -Results    $script:taskResults `
               -DiskBefore $initialFreeGB `
               -DiskAfter  $finalFreeGB `
               -SysInfo    $script:sysInfo
Export-JsonExecutionReport | Out-Null
Export-DeletedPathCandidates | Out-Null

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
        $script:ExitCode = $script:ExitCodes.Partial
    } else {
        $script:ExitCode = $script:ExitCodes.Success
    }
}
exit $script:ExitCode
