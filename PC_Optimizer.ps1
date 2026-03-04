# 高性能 PC 最適化ツール - PowerShell 3.0 / 5.1 / 7.x 対応
# 追加機能:
#  - 配信最適化キャッシュの削除
#  - Windows Update キャッシュの削除
#  - エラーレポート・ログ・不要キャッシュの削除
#  - OneDrive / Teams / Office キャッシュの削除
#  - SSD ヘルスチェック
#  - ブラウザキャッシュ削除（Chrome / Edge / Firefox）
#  - サムネイルキャッシュ削除
#  - Microsoft Store キャッシュクリア
#  - Windows イベントログのクリア
#  - システムファイル整合性チェック・修復（SFC）
#  - Windows コンポーネントストア診断（DISM）
#  - 電源プランの最適化
#  - 実行前後のディスク空き容量比較
# エンコーディング: UTF-8 BOM 付き (PS5.1) / UTF-8 BOM なし (PS7+) -- docs/文字コード規約.md 参照

param()

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

# 各最適化ステップのラッパー
function Try-Step ($desc, [ScriptBlock]$action) {
    $start = Get-Date
    Write-Log "[$($start.ToString('HH:mm:ss'))] $desc 開始..."
    Progress-Bar "$desc..." 0
    try {
        & $action
        Progress-Bar "$desc 完了" 100
        Write-Log "[$((Get-Date).ToString('HH:mm:ss'))] $desc 完了"
    } catch {
        Progress-Bar "$desc 失敗" 100
        Write-Log "[$((Get-Date).ToString('HH:mm:ss'))] $desc 失敗"
        Write-ErrorLog "$desc : $_"
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
    $choice = Read-Host "${B}これらの更新を適用しますか？${RB} (Y/N)"
    Write-Log "[Windows Update] ユーザー選択: $choice"
    if ($choice -match '^[Yy]$') {
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
    Remove-Item -Path "C:\Windows\Temp\*"                              -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "$env:TEMP\*"                                    -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "$env:USERPROFILE\AppData\Local\Temp\*"         -Recurse -Force -ErrorAction SilentlyContinue
}

Try-Step "Prefetch・更新キャッシュの削除" {
    Remove-Item -Path "C:\Windows\Prefetch\*"                                         -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "C:\Windows\SoftwareDistribution\Download\*"                   -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "C:\Windows\System32\DeliveryOptimization\Cache\*"             -Recurse -Force -ErrorAction SilentlyContinue
}

Try-Step "配信最適化キャッシュの削除" {
    Remove-Item -Path "C:\Windows\SoftwareDistribution\DeliveryOptimization\Cache\*" -Recurse -Force -ErrorAction SilentlyContinue
}

Try-Step "Windows Update キャッシュの削除" {
    Remove-Item -Path "C:\Windows\SoftwareDistribution\Download\*"                   -Recurse -Force -ErrorAction SilentlyContinue
}

Try-Step "エラーレポート・ログ・不要キャッシュの削除" {
    # WER レポート
    Remove-Item -Path "C:\ProgramData\Microsoft\Windows\WER\ReportArchive\*" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "C:\ProgramData\Microsoft\Windows\WER\ReportQueue\*"   -Recurse -Force -ErrorAction SilentlyContinue
    # CBS ログ
    Remove-Item -Path "C:\Windows\Logs\CBS\*"                                -Recurse -Force -ErrorAction SilentlyContinue
}

Try-Step "OneDrive / Teams / Office キャッシュの削除" {
    $odtemp = "$env:LOCALAPPDATA\Microsoft\OneDrive\logs"
    if (Test-Path $odtemp) {
        Remove-Item -Path "$odtemp\*" -Recurse -Force -ErrorAction SilentlyContinue
    }
    $teamsCache = "$env:APPDATA\Microsoft\Teams\Cache"
    if (Test-Path $teamsCache) {
        Remove-Item -Path "$teamsCache\*" -Recurse -Force -ErrorAction SilentlyContinue
    }
    $officeCache = "$env:LOCALAPPDATA\Microsoft\Office\16.0\OfficeFileCache"
    if (Test-Path $officeCache) {
        Remove-Item -Path "$officeCache\*" -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Try-Step "ブラウザキャッシュの削除（Chrome / Edge / Firefox）" {
    # Google Chrome
    @(
        "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache",
        "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Code Cache",
        "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\GPUCache"
    ) | Where-Object { Test-Path $_ } | ForEach-Object {
        Remove-Item -Path "$_\*" -Recurse -Force -ErrorAction SilentlyContinue
    }
    # Microsoft Edge
    @(
        "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache",
        "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Code Cache",
        "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\GPUCache"
    ) | Where-Object { Test-Path $_ } | ForEach-Object {
        Remove-Item -Path "$_\*" -Recurse -Force -ErrorAction SilentlyContinue
    }
    # Mozilla Firefox（全プロファイル対応）
    $ffProfileDir = "$env:APPDATA\Mozilla\Firefox\Profiles"
    if (Test-Path $ffProfileDir) {
        Get-ChildItem -Path $ffProfileDir -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $cache = Join-Path $_.FullName "cache2\entries"
            if (Test-Path $cache) {
                Remove-Item -Path "$cache\*" -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

Try-Step "サムネイルキャッシュの削除" {
    $explorerDir = "$env:LOCALAPPDATA\Microsoft\Windows\Explorer"
    if (Test-Path $explorerDir) {
        Get-ChildItem -Path $explorerDir -Filter "thumbcache_*.db" -ErrorAction SilentlyContinue |
            Remove-Item -Force -ErrorAction SilentlyContinue
    }
}

Try-Step "Microsoft Store キャッシュのクリア" {
    @(
        "$env:LOCALAPPDATA\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe\LocalCache",
        "$env:LOCALAPPDATA\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe\LocalState\Cache"
    ) | Where-Object { Test-Path $_ } | ForEach-Object {
        Remove-Item -Path "$_\*" -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Try-Step "ごみ箱を空にする" {
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
        "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe",
        "C:\Program Files (x86)\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe"
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
    $choice = Read-Host "${B}Microsoft 365 の更新を確認・適用しますか？${RB} (Y/N)"
    Write-Log "[Microsoft 365] ユーザー選択: $choice"
    if ($choice -match '^[Yy]$') {
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

# 再起動が必要な場合はプロンプト表示
if (Test-PendingReboot) {
    Write-Log "[再起動] 再起動が必要です。"
    $choice = Read-Host "${B}再起動が必要です。今すぐ再起動しますか？${RB} (Y/N)"
    if ($choice -match '^[Yy]$') {
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
Read-Host "Enter キーを押して終了"
