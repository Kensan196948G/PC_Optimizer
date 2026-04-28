param(
    [switch]$Elevated
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$repoRoot = Split-Path $PSScriptRoot -Parent
$mainScriptPath = Join-Path $repoRoot 'PC_Optimizer.ps1'
$reportsDir = Join-Path $repoRoot 'reports'
$logsDir = Join-Path $repoRoot 'logs'
$docsGuiDir = Join-Path $repoRoot 'docs\GUI'
$configDefaultPath = Join-Path $repoRoot 'config\config.json'
$exportDefaultPath = Join-Path $repoRoot 'logs\gui-deletedpaths'
$heroImagePath = Join-Path $PSScriptRoot 'PC2026.jpg'

$theme = @{
    WindowBack    = [System.Drawing.Color]::FromArgb(248, 250, 255)
    Surface       = [System.Drawing.Color]::White
    Hero          = [System.Drawing.Color]::FromArgb(33, 150, 243)
    HeroSecondary = [System.Drawing.Color]::FromArgb(0, 188, 212)
    Accent        = [System.Drawing.Color]::FromArgb(255, 112, 67)
    Success       = [System.Drawing.Color]::FromArgb(46, 204, 113)
    Warning       = [System.Drawing.Color]::FromArgb(255, 193, 7)
    Danger        = [System.Drawing.Color]::FromArgb(231, 76, 60)
    Text          = [System.Drawing.Color]::FromArgb(41, 50, 65)
    Muted         = [System.Drawing.Color]::FromArgb(103, 116, 142)
    Border        = [System.Drawing.Color]::FromArgb(221, 228, 240)
    LogBack       = [System.Drawing.Color]::FromArgb(23, 27, 38)
    LogText       = [System.Drawing.Color]::FromArgb(225, 230, 240)
}

$taskDefinitions = @(
    [PSCustomObject]@{ Id = 1;  Name = '一時ファイルのクリーンアップ' },
    [PSCustomObject]@{ Id = 2;  Name = 'Prefetch と更新キャッシュの整理' },
    [PSCustomObject]@{ Id = 3;  Name = '配信最適化キャッシュの整理' },
    [PSCustomObject]@{ Id = 4;  Name = 'Windows Update キャッシュの整理' },
    [PSCustomObject]@{ Id = 5;  Name = 'エラーレポートとログの整理' },
    [PSCustomObject]@{ Id = 6;  Name = 'OneDrive / Teams / Office キャッシュ整理' },
    [PSCustomObject]@{ Id = 7;  Name = 'ブラウザキャッシュの整理' },
    [PSCustomObject]@{ Id = 8;  Name = 'サムネイルキャッシュの整理' },
    [PSCustomObject]@{ Id = 9;  Name = 'Microsoft Store キャッシュクリア' },
    [PSCustomObject]@{ Id = 10; Name = 'ごみ箱のクリーンアップ' },
    [PSCustomObject]@{ Id = 11; Name = 'DNS キャッシュのクリア' },
    [PSCustomObject]@{ Id = 12; Name = 'イベントログのクリア' },
    [PSCustomObject]@{ Id = 13; Name = 'ディスクの最適化' },
    [PSCustomObject]@{ Id = 14; Name = 'SSD ヘルスチェック' },
    [PSCustomObject]@{ Id = 15; Name = 'SFC 整合性チェック' },
    [PSCustomObject]@{ Id = 16; Name = 'DISM コンポーネント診断' },
    [PSCustomObject]@{ Id = 17; Name = '電源プランの最適化' },
    [PSCustomObject]@{ Id = 18; Name = 'Microsoft 365 更新確認' },
    [PSCustomObject]@{ Id = 19; Name = 'Windows Update 実行確認' },
    [PSCustomObject]@{ Id = 20; Name = 'スタートアップ / サービスレポート' }
)

function Test-IsAdministrator {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        return $false
    }
}

function Get-CurrentHostPath {
    try {
        $currentPath = (Get-Process -Id $PID).Path
        if ($currentPath -and (Test-Path -LiteralPath $currentPath)) {
            return $currentPath
        }
    } catch {
    }

    foreach ($candidate in @(
        (Join-Path $env:ProgramFiles 'PowerShell\7\pwsh.exe'),
        (Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe')
    )) {
        if ($candidate -and (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
    }

    $pwsh = Get-Command pwsh.exe -ErrorAction SilentlyContinue
    if ($pwsh -and $pwsh.Source) {
        return $pwsh.Source
    }

    return (Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe')
}

function Get-PreferredPowerShellExe {
    $pwshCandidates = @(
        (Join-Path $env:ProgramFiles 'PowerShell\7\pwsh.exe'),
        (Join-Path $env:ProgramFiles 'PowerShell\7-preview\pwsh.exe')
    )

    foreach ($candidate in $pwshCandidates) {
        if ($candidate -and (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
    }

    $pwsh = Get-Command pwsh.exe -ErrorAction SilentlyContinue
    if ($pwsh -and $pwsh.Source) {
        return $pwsh.Source
    }

    return (Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe')
}

function Convert-ToArgumentString {
    param([Parameter(Mandatory)][string[]]$Arguments)

    $escaped = foreach ($argument in $Arguments) {
        if ($null -eq $argument) {
            '""'
            continue
        }

        if ($argument -notmatch '[\s"]') {
            $argument
            continue
        }

        $builder = New-Object System.Text.StringBuilder
        [void]$builder.Append('"')
        $backslashCount = 0

        foreach ($char in $argument.ToCharArray()) {
            if ($char -eq [char]'\') {
                $backslashCount++
            } elseif ($char -eq [char]'"') {
                [void]$builder.Append('\' * ($backslashCount * 2 + 1))
                [void]$builder.Append('"')
                $backslashCount = 0
            } else {
                if ($backslashCount -gt 0) {
                    [void]$builder.Append('\' * $backslashCount)
                    $backslashCount = 0
                }
                [void]$builder.Append($char)
            }
        }

        if ($backslashCount -gt 0) {
            [void]$builder.Append('\' * ($backslashCount * 2))
        }

        [void]$builder.Append('"')
        $builder.ToString()
    }

    return ($escaped -join ' ')
}

function Convert-ToPowerShellLiteral {
    param([AllowNull()][string]$Value)

    if ($null -eq $Value) {
        return '$null'
    }

    return "'" + ($Value -replace "'", "''") + "'"
}

function ConvertTo-TaskToken {
    param([Parameter(Mandatory)][int[]]$TaskIds)

    $ordered = @($TaskIds | Sort-Object -Unique)
    if ($ordered.Count -eq 20) {
        return '1-20'
    }
    return ($ordered -join ',')
}

function Get-TaskDisplayName {
    param([Parameter(Mandatory)][int]$TaskId)

    $task = $taskDefinitions | Where-Object { $_.Id -eq $TaskId } | Select-Object -First 1
    if ($task) {
        return $task.Name
    }
    return "タスク $TaskId"
}

function Remove-AnsiEscapeSequence {
    param([AllowNull()][string]$Text)

    if ($null -eq $Text) {
        return ''
    }

    $cleaned = [regex]::Replace($Text, '\x1B\[[0-9;?]*[ -/]*[@-~]', '')
    $cleaned = $cleaned -replace '[\x00-\x08\x0B\x0C\x0E-\x1F]', ''
    return $cleaned.TrimEnd()
}

function Get-TaskIdFromLine {
    param([Parameter(Mandatory)][string]$Line)

    $taskMatch = [regex]::Match($Line, '#(?<id>\d{1,2})')
    if ($taskMatch.Success) {
        $id = [int]$taskMatch.Groups['id'].Value
        if ($id -ge 1 -and $id -le 20) {
            return $id
        }
    }

    if ($sync.CompletedCount -lt $sync.SelectedIds.Count) {
        return [int]$sync.SelectedIds[$sync.CompletedCount]
    }

    return $null
}

function Start-SelfElevated {
    param([Parameter(Mandatory)][string]$ScriptPath)

    $hostPath = Get-CurrentHostPath
    $args = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', $ScriptPath,
        '-Elevated'
    )

    try {
        Start-Process -FilePath $hostPath -ArgumentList (Convert-ToArgumentString -Arguments $args) -Verb RunAs | Out-Null
        exit
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "管理者権限での再起動に失敗しました。`r`n$_",
            'PC Optimizer GUI',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        exit 1
    }
}

function New-ThemeFont {
    param(
        [string]$Name = 'Segoe UI',
        [float]$Size = 9.0,
        [System.Drawing.FontStyle]$Style = [System.Drawing.FontStyle]::Regular
    )

    return New-Object System.Drawing.Font($Name, $Size, $Style)
}

function Set-FlatButtonStyle {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.Button]$Button,
        [Parameter(Mandatory)][System.Drawing.Color]$BackColor,
        [Parameter(Mandatory)][System.Drawing.Color]$ForeColor
    )

    $Button.FlatStyle = 'Flat'
    $Button.FlatAppearance.BorderSize = 0
    $Button.BackColor = $BackColor
    $Button.ForeColor = $ForeColor
    $Button.Font = New-ThemeFont -Size 9.5 -Style Bold
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
}

function New-ChipLabel {
    param(
        [int]$X,
        [int]$Y,
        [int]$Width,
        [string]$Text
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point($X, $Y)
    $label.Size = New-Object System.Drawing.Size($Width, 26)
    $label.Text = $Text
    $label.BackColor = [System.Drawing.Color]::FromArgb(52, 170, 255)
    $label.ForeColor = [System.Drawing.Color]::White
    $label.TextAlign = 'MiddleCenter'
    $label.Font = New-ThemeFont -Size 8.8 -Style Bold
    return $label
}

function Open-PathInShell {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        [System.Windows.Forms.MessageBox]::Show(
            "パスが見つかりません。`r`n$Path",
            'PC Optimizer GUI',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    Start-Process -FilePath 'explorer.exe' -ArgumentList ('"{0}"' -f $Path) | Out-Null
}

if (-not (Test-Path -LiteralPath $mainScriptPath)) {
    [System.Windows.Forms.MessageBox]::Show(
        "バックエンドスクリプトが見つかりません。`r`n$mainScriptPath",
        'PC Optimizer GUI',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    exit 1
}

if (-not $Elevated -and -not (Test-IsAdministrator)) {
    Start-SelfElevated -ScriptPath $PSCommandPath
}

if ($Elevated -and -not (Test-IsAdministrator)) {
    [System.Windows.Forms.MessageBox]::Show(
        '管理者権限を取得できませんでした。',
        'PC Optimizer GUI',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    exit 1
}

$sync = [hashtable]::Synchronized(@{
    Queue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    Process = $null
    Running = $false
    StartTime = $null
    ExitCode = $null
    OutputEventId = $null
    ErrorEventId = $null
    ExitEventId = $null
    SelectedIds = @()
    CompletedCount = 0
    CompletedTaskIds = New-Object 'System.Collections.Generic.HashSet[int]'
    StopRequested = $false
    LatestHtml = ''
    LatestJson = ''
    CurrentTask = ''
    WrapperPath = ''
})

$form = New-Object System.Windows.Forms.Form
$form.Text = 'PC Optimizer GUI for PowerShell 7'
$form.StartPosition = 'CenterScreen'
$form.ClientSize = New-Object System.Drawing.Size(1360, 900)
$form.MinimumSize = New-Object System.Drawing.Size(1200, 900)
$form.BackColor = $theme.WindowBack
$form.Font = New-ThemeFont -Size 9

$heroPanel = New-Object System.Windows.Forms.Panel
$heroPanel.Location = New-Object System.Drawing.Point(12, 12)
$heroPanel.Size = New-Object System.Drawing.Size(1320, 150)
$heroPanel.BackColor = $theme.Hero
$heroPanel.Anchor = 'Top,Left,Right'

$heroAccent = New-Object System.Windows.Forms.Panel
$heroAccent.Location = New-Object System.Drawing.Point(0, 118)
$heroAccent.Size = New-Object System.Drawing.Size(1320, 32)
$heroAccent.BackColor = $theme.HeroSecondary
$heroAccent.Anchor = 'Left,Right,Bottom'

$heroImage = New-Object System.Windows.Forms.PictureBox
$heroImage.Location = New-Object System.Drawing.Point(1038, 16)
$heroImage.Size = New-Object System.Drawing.Size(258, 118)
$heroImage.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
$heroImage.SizeMode = 'Zoom'
$heroImage.BorderStyle = 'FixedSingle'
if (Test-Path -LiteralPath $heroImagePath) {
    $heroImage.Image = [System.Drawing.Image]::FromFile($heroImagePath)
}

$lblHeroTitle = New-Object System.Windows.Forms.Label
$lblHeroTitle.Location = New-Object System.Drawing.Point(24, 16)
$lblHeroTitle.Size = New-Object System.Drawing.Size(700, 34)
$lblHeroTitle.Text = 'PC Optimizer GUI'
$lblHeroTitle.ForeColor = [System.Drawing.Color]::White
$lblHeroTitle.BackColor = [System.Drawing.Color]::Transparent
$lblHeroTitle.Font = New-ThemeFont -Size 22 -Style Bold

$lblHeroSubtitle = New-Object System.Windows.Forms.Label
$lblHeroSubtitle.Location = New-Object System.Drawing.Point(27, 56)
$lblHeroSubtitle.Size = New-Object System.Drawing.Size(860, 24)
$lblHeroSubtitle.Text = 'PowerShell 7 優先 / CLI バックエンド再利用 / ポップな運用向けフロントエンド'
$lblHeroSubtitle.ForeColor = [System.Drawing.Color]::White
$lblHeroSubtitle.BackColor = [System.Drawing.Color]::Transparent
$lblHeroSubtitle.Font = New-ThemeFont -Size 10.5 -Style Bold

$lblEngineBadge = New-ChipLabel -X 28 -Y 92 -Width 350 -Text ("Engine: " + (Get-PreferredPowerShellExe))
$lblAdminBadge = New-ChipLabel -X 388 -Y 92 -Width 170 -Text '権限: 管理者'
$lblModeBadge = New-ChipLabel -X 568 -Y 92 -Width 220 -Text 'Mode: repair / classic'

$heroPanel.Controls.AddRange(@(
    $heroAccent,
    $lblHeroTitle,
    $lblHeroSubtitle,
    $lblEngineBadge,
    $lblAdminBadge,
    $lblModeBadge,
    $heroImage
))

$cardSummary = New-Object System.Windows.Forms.Panel
$cardSummary.Location = New-Object System.Drawing.Point(12, 174)
$cardSummary.Size = New-Object System.Drawing.Size(1320, 72)
$cardSummary.BackColor = $theme.Surface
$cardSummary.BorderStyle = 'FixedSingle'
$cardSummary.Anchor = 'Top,Left,Right'

$lblSummaryTitle = New-Object System.Windows.Forms.Label
$lblSummaryTitle.Location = New-Object System.Drawing.Point(18, 10)
$lblSummaryTitle.Size = New-Object System.Drawing.Size(400, 22)
$lblSummaryTitle.Text = '現在の構成'
$lblSummaryTitle.ForeColor = $theme.Text
$lblSummaryTitle.Font = New-ThemeFont -Size 11 -Style Bold

$lblSummaryInfo = New-Object System.Windows.Forms.Label
$lblSummaryInfo.Location = New-Object System.Drawing.Point(18, 38)
$lblSummaryInfo.Size = New-Object System.Drawing.Size(860, 22)
$lblSummaryInfo.Text = '未実行'
$lblSummaryInfo.ForeColor = $theme.Muted
$lblSummaryInfo.Font = New-ThemeFont -Size 9.5

$lblSummaryNote = New-Object System.Windows.Forms.Label
$lblSummaryNote.Location = New-Object System.Drawing.Point(910, 16)
$lblSummaryNote.Size = New-Object System.Drawing.Size(388, 40)
$lblSummaryNote.Text = 'GUI 実行時は -NonInteractive -NoRebootPrompt を常時付与します。'
$lblSummaryNote.ForeColor = $theme.Text
$lblSummaryNote.Font = New-ThemeFont -Size 9.2 -Style Bold

$cardSummary.Controls.AddRange(@($lblSummaryTitle, $lblSummaryInfo, $lblSummaryNote))

$grpSettings = New-Object System.Windows.Forms.GroupBox
$grpSettings.Text = '実行設定'
$grpSettings.Location = New-Object System.Drawing.Point(12, 258)
$grpSettings.Size = New-Object System.Drawing.Size(430, 392)
$grpSettings.BackColor = $theme.Surface
$grpSettings.ForeColor = $theme.Text
$grpSettings.Font = New-ThemeFont -Size 10 -Style Bold
$grpSettings.Anchor = 'Top,Left'

$lblMode = New-Object System.Windows.Forms.Label
$lblMode.Location = New-Object System.Drawing.Point(18, 32)
$lblMode.Size = New-Object System.Drawing.Size(120, 20)
$lblMode.Text = 'モード'
$lblMode.Font = New-ThemeFont -Size 9.5 -Style Bold

$cmbMode = New-Object System.Windows.Forms.ComboBox
$cmbMode.Location = New-Object System.Drawing.Point(152, 28)
$cmbMode.Size = New-Object System.Drawing.Size(248, 28)
$cmbMode.DropDownStyle = 'DropDownList'
$cmbMode.Font = New-ThemeFont -Size 9.5
[void]$cmbMode.Items.AddRange(@('repair', 'diagnose'))
$cmbMode.SelectedIndex = 0

$lblProfile = New-Object System.Windows.Forms.Label
$lblProfile.Location = New-Object System.Drawing.Point(18, 70)
$lblProfile.Size = New-Object System.Drawing.Size(120, 20)
$lblProfile.Text = '実行プロファイル'
$lblProfile.Font = New-ThemeFont -Size 9.5 -Style Bold

$cmbProfile = New-Object System.Windows.Forms.ComboBox
$cmbProfile.Location = New-Object System.Drawing.Point(152, 66)
$cmbProfile.Size = New-Object System.Drawing.Size(248, 28)
$cmbProfile.DropDownStyle = 'DropDownList'
$cmbProfile.Font = New-ThemeFont -Size 9.5
[void]$cmbProfile.Items.AddRange(@('classic', 'agent-teams'))
$cmbProfile.SelectedIndex = 0

$lblFailure = New-Object System.Windows.Forms.Label
$lblFailure.Location = New-Object System.Drawing.Point(18, 108)
$lblFailure.Size = New-Object System.Drawing.Size(120, 20)
$lblFailure.Text = '失敗時動作'
$lblFailure.Font = New-ThemeFont -Size 9.5 -Style Bold

$cmbFailure = New-Object System.Windows.Forms.ComboBox
$cmbFailure.Location = New-Object System.Drawing.Point(152, 104)
$cmbFailure.Size = New-Object System.Drawing.Size(248, 28)
$cmbFailure.DropDownStyle = 'DropDownList'
$cmbFailure.Font = New-ThemeFont -Size 9.5
[void]$cmbFailure.Items.AddRange(@('continue', 'fail-fast'))
$cmbFailure.SelectedIndex = 0

$lblConfig = New-Object System.Windows.Forms.Label
$lblConfig.Location = New-Object System.Drawing.Point(18, 146)
$lblConfig.Size = New-Object System.Drawing.Size(120, 20)
$lblConfig.Text = 'ConfigPath'
$lblConfig.Font = New-ThemeFont -Size 9.5 -Style Bold

$txtConfig = New-Object System.Windows.Forms.TextBox
$txtConfig.Location = New-Object System.Drawing.Point(18, 170)
$txtConfig.Size = New-Object System.Drawing.Size(304, 24)
$txtConfig.Text = $configDefaultPath
$txtConfig.Font = New-ThemeFont -Size 9.2

$btnBrowseConfig = New-Object System.Windows.Forms.Button
$btnBrowseConfig.Location = New-Object System.Drawing.Point(332, 168)
$btnBrowseConfig.Size = New-Object System.Drawing.Size(68, 28)
$btnBrowseConfig.Text = '参照'
Set-FlatButtonStyle -Button $btnBrowseConfig -BackColor $theme.HeroSecondary -ForeColor ([System.Drawing.Color]::White)

$lblExportPath = New-Object System.Windows.Forms.Label
$lblExportPath.Location = New-Object System.Drawing.Point(18, 208)
$lblExportPath.Size = New-Object System.Drawing.Size(220, 20)
$lblExportPath.Text = 'ExportDeletedPathsPath'
$lblExportPath.Font = New-ThemeFont -Size 9.5 -Style Bold

$txtExportPath = New-Object System.Windows.Forms.TextBox
$txtExportPath.Location = New-Object System.Drawing.Point(18, 232)
$txtExportPath.Size = New-Object System.Drawing.Size(304, 24)
$txtExportPath.Text = $exportDefaultPath
$txtExportPath.Font = New-ThemeFont -Size 9.2

$btnBrowseExport = New-Object System.Windows.Forms.Button
$btnBrowseExport.Location = New-Object System.Drawing.Point(332, 230)
$btnBrowseExport.Size = New-Object System.Drawing.Size(68, 28)
$btnBrowseExport.Text = '参照'
Set-FlatButtonStyle -Button $btnBrowseExport -BackColor $theme.HeroSecondary -ForeColor ([System.Drawing.Color]::White)

$lblOptions = New-Object System.Windows.Forms.Label
$lblOptions.Location = New-Object System.Drawing.Point(18, 270)
$lblOptions.Size = New-Object System.Drawing.Size(220, 20)
$lblOptions.Text = 'オプション'
$lblOptions.Font = New-ThemeFont -Size 10 -Style Bold

$chkWhatIf = New-Object System.Windows.Forms.CheckBox
$chkWhatIf.Location = New-Object System.Drawing.Point(22, 296)
$chkWhatIf.Size = New-Object System.Drawing.Size(170, 22)
$chkWhatIf.Text = 'WhatIf で安全に確認'
$chkWhatIf.Font = New-ThemeFont -Size 9.2

$chkUseLocalChart = New-Object System.Windows.Forms.CheckBox
$chkUseLocalChart.Location = New-Object System.Drawing.Point(205, 296)
$chkUseLocalChart.Size = New-Object System.Drawing.Size(180, 22)
$chkUseLocalChart.Text = 'ローカル Chart.js を使う'
$chkUseLocalChart.Checked = $true
$chkUseLocalChart.Font = New-ThemeFont -Size 9.2

$chkExportPowerBI = New-Object System.Windows.Forms.CheckBox
$chkExportPowerBI.Location = New-Object System.Drawing.Point(22, 324)
$chkExportPowerBI.Size = New-Object System.Drawing.Size(176, 22)
$chkExportPowerBI.Text = 'Power BI JSON を出力'
$chkExportPowerBI.Font = New-ThemeFont -Size 9.2

$chkUseAnthropic = New-Object System.Windows.Forms.CheckBox
$chkUseAnthropic.Location = New-Object System.Drawing.Point(216, 324)
$chkUseAnthropic.Size = New-Object System.Drawing.Size(168, 22)
$chkUseAnthropic.Text = 'AI診断を利用'
$chkUseAnthropic.Font = New-ThemeFont -Size 9.2
$chkUseAnthropic.ForeColor = $theme.Text
$chkUseAnthropic.BackColor = $theme.Surface
$chkUseAnthropic.UseVisualStyleBackColor = $false

$chkExportDeleted = New-Object System.Windows.Forms.CheckBox
$chkExportDeleted.Location = New-Object System.Drawing.Point(22, 352)
$chkExportDeleted.Size = New-Object System.Drawing.Size(170, 22)
$chkExportDeleted.Text = '削除候補をエクスポート'
$chkExportDeleted.Font = New-ThemeFont -Size 9.2

$cmbExportFormat = New-Object System.Windows.Forms.ComboBox
$cmbExportFormat.Location = New-Object System.Drawing.Point(205, 349)
$cmbExportFormat.Size = New-Object System.Drawing.Size(88, 28)
$cmbExportFormat.DropDownStyle = 'DropDownList'
$cmbExportFormat.Font = New-ThemeFont -Size 9.2
[void]$cmbExportFormat.Items.AddRange(@('json', 'csv'))
$cmbExportFormat.SelectedIndex = 0

$grpSettings.Controls.AddRange(@(
    $lblMode, $cmbMode, $lblProfile, $cmbProfile, $lblFailure, $cmbFailure,
    $lblConfig, $txtConfig, $btnBrowseConfig,
    $lblExportPath, $txtExportPath, $btnBrowseExport,
    $lblOptions, $chkWhatIf, $chkUseLocalChart, $chkExportPowerBI,
    $chkUseAnthropic, $chkExportDeleted, $cmbExportFormat
))

$grpTasks = New-Object System.Windows.Forms.GroupBox
$grpTasks.Text = 'タスク選択'
$grpTasks.Location = New-Object System.Drawing.Point(452, 258)
$grpTasks.Size = New-Object System.Drawing.Size(434, 392)
$grpTasks.BackColor = $theme.Surface
$grpTasks.ForeColor = $theme.Text
$grpTasks.Font = New-ThemeFont -Size 10 -Style Bold
$grpTasks.Anchor = 'Top,Left'

$lblTaskHint = New-Object System.Windows.Forms.Label
$lblTaskHint.Location = New-Object System.Drawing.Point(16, 30)
$lblTaskHint.Size = New-Object System.Drawing.Size(400, 20)
$lblTaskHint.Text = 'Task 18 / 19 は GUI では実行確認のみ。NonInteractive により自動スキップされます。'
$lblTaskHint.ForeColor = $theme.Muted
$lblTaskHint.Font = New-ThemeFont -Size 8.7

$taskList = New-Object System.Windows.Forms.CheckedListBox
$taskList.Location = New-Object System.Drawing.Point(16, 58)
$taskList.Size = New-Object System.Drawing.Size(398, 244)
$taskList.CheckOnClick = $true
$taskList.Font = New-ThemeFont -Size 9.2
$taskList.BorderStyle = 'FixedSingle'

foreach ($task in $taskDefinitions) {
    [void]$taskList.Items.Add(("{0:D2}: {1}" -f $task.Id, $task.Name), $true)
}

$btnSelectAll = New-Object System.Windows.Forms.Button
$btnSelectAll.Location = New-Object System.Drawing.Point(16, 320)
$btnSelectAll.Size = New-Object System.Drawing.Size(120, 34)
$btnSelectAll.Text = '全部選択'
Set-FlatButtonStyle -Button $btnSelectAll -BackColor $theme.Success -ForeColor ([System.Drawing.Color]::White)

$btnClearAll = New-Object System.Windows.Forms.Button
$btnClearAll.Location = New-Object System.Drawing.Point(148, 320)
$btnClearAll.Size = New-Object System.Drawing.Size(120, 34)
$btnClearAll.Text = '全部解除'
Set-FlatButtonStyle -Button $btnClearAll -BackColor $theme.Warning -ForeColor $theme.Text

$btnDiagnosePreset = New-Object System.Windows.Forms.Button
$btnDiagnosePreset.Location = New-Object System.Drawing.Point(280, 320)
$btnDiagnosePreset.Size = New-Object System.Drawing.Size(134, 34)
$btnDiagnosePreset.Text = '診断向けプリセット'
Set-FlatButtonStyle -Button $btnDiagnosePreset -BackColor $theme.Accent -ForeColor ([System.Drawing.Color]::White)

$lblTaskCount = New-Object System.Windows.Forms.Label
$lblTaskCount.Location = New-Object System.Drawing.Point(16, 360)
$lblTaskCount.Size = New-Object System.Drawing.Size(300, 22)
$lblTaskCount.Text = '選択中: 20 / 20'
$lblTaskCount.ForeColor = $theme.Text
$lblTaskCount.Font = New-ThemeFont -Size 9.5 -Style Bold

$grpTasks.Controls.AddRange(@($lblTaskHint, $taskList, $btnSelectAll, $btnClearAll, $btnDiagnosePreset, $lblTaskCount))

$grpStatus = New-Object System.Windows.Forms.GroupBox
$grpStatus.Text = '実行ステータス'
$grpStatus.Location = New-Object System.Drawing.Point(896, 258)
$grpStatus.Size = New-Object System.Drawing.Size(436, 392)
$grpStatus.BackColor = $theme.Surface
$grpStatus.ForeColor = $theme.Text
$grpStatus.Font = New-ThemeFont -Size 10 -Style Bold
$grpStatus.Anchor = 'Top,Left,Right'

$statusBadge = New-Object System.Windows.Forms.Label
$statusBadge.Location = New-Object System.Drawing.Point(18, 32)
$statusBadge.Size = New-Object System.Drawing.Size(120, 32)
$statusBadge.Text = '待機中'
$statusBadge.TextAlign = 'MiddleCenter'
$statusBadge.BackColor = $theme.HeroSecondary
$statusBadge.ForeColor = [System.Drawing.Color]::White
$statusBadge.Font = New-ThemeFont -Size 10 -Style Bold

$lblCurrentTask = New-Object System.Windows.Forms.Label
$lblCurrentTask.Location = New-Object System.Drawing.Point(18, 78)
$lblCurrentTask.Size = New-Object System.Drawing.Size(396, 22)
$lblCurrentTask.Text = '現在タスク: -'
$lblCurrentTask.Font = New-ThemeFont -Size 10 -Style Bold

$lblElapsed = New-Object System.Windows.Forms.Label
$lblElapsed.Location = New-Object System.Drawing.Point(18, 106)
$lblElapsed.Size = New-Object System.Drawing.Size(188, 20)
$lblElapsed.Text = '経過時間: 00:00:00'
$lblElapsed.Font = New-ThemeFont -Size 9.2

$lblExitCode = New-Object System.Windows.Forms.Label
$lblExitCode.Location = New-Object System.Drawing.Point(220, 106)
$lblExitCode.Size = New-Object System.Drawing.Size(188, 20)
$lblExitCode.Text = '終了コード: -'
$lblExitCode.Font = New-ThemeFont -Size 9.2

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(18, 136)
$progress.Size = New-Object System.Drawing.Size(396, 26)
$progress.Minimum = 0
$progress.Maximum = 100
$progress.Value = 0
$progress.Style = 'Continuous'

$lblProgress = New-Object System.Windows.Forms.Label
$lblProgress.Location = New-Object System.Drawing.Point(18, 170)
$lblProgress.Size = New-Object System.Drawing.Size(396, 20)
$lblProgress.Text = '進捗: 0 / 0'
$lblProgress.Font = New-ThemeFont -Size 9.2

$lblLatestHtml = New-Object System.Windows.Forms.Label
$lblLatestHtml.Location = New-Object System.Drawing.Point(18, 206)
$lblLatestHtml.Size = New-Object System.Drawing.Size(396, 36)
$lblLatestHtml.Text = '最新 HTML: -'
$lblLatestHtml.Font = New-ThemeFont -Size 8.8

$lblLatestJson = New-Object System.Windows.Forms.Label
$lblLatestJson.Location = New-Object System.Drawing.Point(18, 246)
$lblLatestJson.Size = New-Object System.Drawing.Size(396, 36)
$lblLatestJson.Text = '最新 JSON: -'
$lblLatestJson.Font = New-ThemeFont -Size 8.8

$btnOpenLatestHtml = New-Object System.Windows.Forms.Button
$btnOpenLatestHtml.Location = New-Object System.Drawing.Point(18, 300)
$btnOpenLatestHtml.Size = New-Object System.Drawing.Size(190, 34)
$btnOpenLatestHtml.Text = '最新 HTML を開く'
Set-FlatButtonStyle -Button $btnOpenLatestHtml -BackColor $theme.Hero -ForeColor ([System.Drawing.Color]::White)

$btnOpenLatestJson = New-Object System.Windows.Forms.Button
$btnOpenLatestJson.Location = New-Object System.Drawing.Point(224, 300)
$btnOpenLatestJson.Size = New-Object System.Drawing.Size(190, 34)
$btnOpenLatestJson.Text = '最新 JSON を開く'
Set-FlatButtonStyle -Button $btnOpenLatestJson -BackColor $theme.HeroSecondary -ForeColor ([System.Drawing.Color]::White)

$lblStatusFoot = New-Object System.Windows.Forms.Label
$lblStatusFoot.Location = New-Object System.Drawing.Point(18, 350)
$lblStatusFoot.Size = New-Object System.Drawing.Size(396, 24)
$lblStatusFoot.Text = '実行エンジン: 初期化中'
$lblStatusFoot.ForeColor = $theme.Muted
$lblStatusFoot.Font = New-ThemeFont -Size 8.8 -Style Bold

$grpStatus.Controls.AddRange(@(
    $statusBadge, $lblCurrentTask, $lblElapsed, $lblExitCode, $progress, $lblProgress,
    $lblLatestHtml, $lblLatestJson, $btnOpenLatestHtml, $btnOpenLatestJson, $lblStatusFoot
))

$grpCommand = New-Object System.Windows.Forms.GroupBox
$grpCommand.Text = 'コマンドプレビュー'
$grpCommand.Location = New-Object System.Drawing.Point(12, 662)
$grpCommand.Size = New-Object System.Drawing.Size(640, 188)
$grpCommand.BackColor = $theme.Surface
$grpCommand.ForeColor = $theme.Text
$grpCommand.Font = New-ThemeFont -Size 10 -Style Bold
$grpCommand.Anchor = 'Top,Left,Bottom'

$txtCommand = New-Object System.Windows.Forms.TextBox
$txtCommand.Location = New-Object System.Drawing.Point(16, 30)
$txtCommand.Size = New-Object System.Drawing.Size(608, 140)
$txtCommand.Multiline = $true
$txtCommand.ReadOnly = $true
$txtCommand.ScrollBars = 'Vertical'
$txtCommand.BackColor = [System.Drawing.Color]::FromArgb(245, 248, 255)
$txtCommand.ForeColor = $theme.Text
$txtCommand.Font = New-ThemeFont -Name 'Consolas' -Size 9

$grpCommand.Controls.Add($txtCommand)

$grpLog = New-Object System.Windows.Forms.GroupBox
$grpLog.Text = '実行ログ'
$grpLog.Location = New-Object System.Drawing.Point(664, 662)
$grpLog.Size = New-Object System.Drawing.Size(668, 188)
$grpLog.BackColor = $theme.Surface
$grpLog.ForeColor = $theme.Text
$grpLog.Font = New-ThemeFont -Size 10 -Style Bold
$grpLog.Anchor = 'Top,Left,Right,Bottom'

$txtLog = New-Object System.Windows.Forms.RichTextBox
$txtLog.Location = New-Object System.Drawing.Point(16, 30)
$txtLog.Size = New-Object System.Drawing.Size(636, 140)
$txtLog.ReadOnly = $true
$txtLog.BackColor = $theme.LogBack
$txtLog.ForeColor = $theme.LogText
$txtLog.BorderStyle = 'FixedSingle'
$txtLog.Font = New-ThemeFont -Name 'Consolas' -Size 9
$txtLog.DetectUrls = $false

$grpLog.Controls.Add($txtLog)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Location = New-Object System.Drawing.Point(12, 860)
$btnRun.Size = New-Object System.Drawing.Size(150, 36)
$btnRun.Text = '▶ 実行'
Set-FlatButtonStyle -Button $btnRun -BackColor $theme.Success -ForeColor ([System.Drawing.Color]::White)
$btnRun.Anchor = 'Left,Bottom'

$btnStop = New-Object System.Windows.Forms.Button
$btnStop.Location = New-Object System.Drawing.Point(174, 860)
$btnStop.Size = New-Object System.Drawing.Size(150, 36)
$btnStop.Text = '■ 停止'
Set-FlatButtonStyle -Button $btnStop -BackColor $theme.Danger -ForeColor ([System.Drawing.Color]::White)
$btnStop.Enabled = $false
$btnStop.Anchor = 'Left,Bottom'

$btnOpenReports = New-Object System.Windows.Forms.Button
$btnOpenReports.Location = New-Object System.Drawing.Point(336, 860)
$btnOpenReports.Size = New-Object System.Drawing.Size(150, 36)
$btnOpenReports.Text = '📊 reports'
Set-FlatButtonStyle -Button $btnOpenReports -BackColor $theme.Hero -ForeColor ([System.Drawing.Color]::White)
$btnOpenReports.Anchor = 'Left,Bottom'

$btnOpenLogs = New-Object System.Windows.Forms.Button
$btnOpenLogs.Location = New-Object System.Drawing.Point(498, 860)
$btnOpenLogs.Size = New-Object System.Drawing.Size(150, 36)
$btnOpenLogs.Text = '🧾 logs'
Set-FlatButtonStyle -Button $btnOpenLogs -BackColor $theme.HeroSecondary -ForeColor ([System.Drawing.Color]::White)
$btnOpenLogs.Anchor = 'Left,Bottom'

$btnOpenDocs = New-Object System.Windows.Forms.Button
$btnOpenDocs.Location = New-Object System.Drawing.Point(660, 860)
$btnOpenDocs.Size = New-Object System.Drawing.Size(180, 36)
$btnOpenDocs.Text = '📘 docs/GUI'
Set-FlatButtonStyle -Button $btnOpenDocs -BackColor $theme.Accent -ForeColor ([System.Drawing.Color]::White)
$btnOpenDocs.Anchor = 'Left,Bottom'

$btnOpenRepo = New-Object System.Windows.Forms.Button
$btnOpenRepo.Location = New-Object System.Drawing.Point(852, 860)
$btnOpenRepo.Size = New-Object System.Drawing.Size(180, 36)
$btnOpenRepo.Text = '📁 リポジトリを開く'
Set-FlatButtonStyle -Button $btnOpenRepo -BackColor $theme.Warning -ForeColor $theme.Text
$btnOpenRepo.Anchor = 'Left,Bottom'

$form.Controls.AddRange(@(
    $heroPanel, $cardSummary, $grpSettings, $grpTasks, $grpStatus, $grpCommand, $grpLog,
    $btnRun, $btnStop, $btnOpenReports, $btnOpenLogs, $btnOpenDocs, $btnOpenRepo
))

$controlsToLock = @(
    $cmbMode, $cmbProfile, $cmbFailure, $txtConfig, $txtExportPath, $taskList,
    $chkWhatIf, $chkUseLocalChart, $chkExportPowerBI, $chkUseAnthropic, $chkExportDeleted,
    $cmbExportFormat, $btnSelectAll, $btnClearAll, $btnDiagnosePreset, $btnBrowseConfig, $btnBrowseExport
)

function Set-StatusBadge {
    param(
        [Parameter(Mandatory)][string]$Text,
        [Parameter(Mandatory)][System.Drawing.Color]$BackColor
    )

    $statusBadge.Text = $Text
    $statusBadge.BackColor = $BackColor
}

function Set-ControlsEnabled {
    param([bool]$Enabled)

    foreach ($control in $controlsToLock) {
        $control.Enabled = $Enabled
    }
    $btnRun.Enabled = $Enabled
    $btnStop.Enabled = -not $Enabled
}

function Get-SelectedTaskIds {
    $ids = New-Object 'System.Collections.Generic.List[int]'
    foreach ($item in $taskList.CheckedItems) {
        $text = [string]$item
        if ($text -match '^(\d{2}):') {
            [void]$ids.Add([int]$Matches[1])
        }
    }
    return @($ids | Sort-Object)
}

function Get-ShortPathLabel {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return '-'
    }
    if ($Path.Length -le 58) {
        return $Path
    }
    return ('...' + $Path.Substring($Path.Length - 55))
}

function Find-LatestArtifacts {
    $sync.LatestHtml = ''
    $sync.LatestJson = ''

    if (Test-Path -LiteralPath $reportsDir) {
        $latestHtml = Get-ChildItem -LiteralPath $reportsDir -Filter '*.html' -File -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTimeUtc -Descending |
            Select-Object -First 1
        $latestJson = Get-ChildItem -LiteralPath $reportsDir -Filter '*.json' -File -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTimeUtc -Descending |
            Select-Object -First 1
        if ($latestHtml) { $sync.LatestHtml = $latestHtml.FullName }
        if ($latestJson) { $sync.LatestJson = $latestJson.FullName }
    }

    if (-not $sync.LatestHtml -and (Test-Path -LiteralPath $logsDir)) {
        $fallbackHtml = Get-ChildItem -LiteralPath $logsDir -Filter 'PC_Optimizer_Report_*.html' -File -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTimeUtc -Descending |
            Select-Object -First 1
        if ($fallbackHtml) {
            $sync.LatestHtml = $fallbackHtml.FullName
        }
    }
}

function Update-ArtifactLabels {
    $lblLatestHtml.Text = '最新 HTML: ' + (Get-ShortPathLabel -Path $sync.LatestHtml)
    $lblLatestJson.Text = '最新 JSON: ' + (Get-ShortPathLabel -Path $sync.LatestJson)
    $btnOpenLatestHtml.Enabled = [bool]$sync.LatestHtml
    $btnOpenLatestJson.Enabled = [bool]$sync.LatestJson
}

function New-ExecutionArguments {
    param([int[]]$TaskIdsOverride = $null)

    $selectedIds = Get-SelectedTaskIds
    if ($TaskIdsOverride) {
        $selectedIds = @($TaskIdsOverride | Sort-Object -Unique)
    }

    $args = New-Object 'System.Collections.Generic.List[string]'
    [void]$args.Add('-NonInteractive')
    [void]$args.Add('-NoRebootPrompt')
    [void]$args.Add('-Mode')
    [void]$args.Add([string]$cmbMode.SelectedItem)
    [void]$args.Add('-ExecutionProfile')
    [void]$args.Add([string]$cmbProfile.SelectedItem)
    [void]$args.Add('-Tasks')
    [void]$args.Add((ConvertTo-TaskToken -TaskIds $selectedIds))
    [void]$args.Add('-FailureMode')
    [void]$args.Add([string]$cmbFailure.SelectedItem)

    if ($chkWhatIf.Checked) {
        [void]$args.Add('-WhatIf')
    }
    if ($chkUseLocalChart.Checked) {
        [void]$args.Add('-UseLocalChartJs')
    }
    if ($chkExportPowerBI.Checked) {
        [void]$args.Add('-ExportPowerBIJson')
    }
    if ($chkUseAnthropic.Checked) {
        [void]$args.Add('-UseAnthropicAI')
    }
    if ($txtConfig.Text.Trim()) {
        [void]$args.Add('-ConfigPath')
        [void]$args.Add($txtConfig.Text.Trim())
    }
    if ($chkExportDeleted.Checked) {
        [void]$args.Add('-ExportDeletedPaths')
        [void]$args.Add([string]$cmbExportFormat.SelectedItem)
        if ($txtExportPath.Text.Trim()) {
            [void]$args.Add('-ExportDeletedPathsPath')
            [void]$args.Add($txtExportPath.Text.Trim())
        }
    }

    return @{
        Args = @($args)
        SelectedIds = @($selectedIds)
    }
}

function New-HostInvocation {
    param(
        [Parameter(Mandatory)][string]$EnginePath,
        [Parameter(Mandatory)][string[]]$BackendArgs
    )

    $invocationTokens = foreach ($token in $BackendArgs) {
        if ($token -match '^-[A-Za-z]') {
            $token
        } else {
            Convert-ToPowerShellLiteral -Value $token
        }
    }

    $wrapperPath = Join-Path $logsDir ("gui-runner-{0}.ps1" -f ([guid]::NewGuid().ToString('N')))
    $wrapperContent = @(
        '$OutputEncoding = [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)'
        '$env:NO_COLOR = ''1'''
        'if (Get-Variable -Name PSStyle -ErrorAction SilentlyContinue) { $PSStyle.OutputRendering = ''PlainText'' }'
        ('& ' + (Convert-ToPowerShellLiteral -Value $mainScriptPath) + ' ' + ($invocationTokens -join ' '))
    ) -join [Environment]::NewLine
    $utf8Bom = [System.Text.UTF8Encoding]::new($true)
    [System.IO.File]::WriteAllText($wrapperPath, $wrapperContent, $utf8Bom)

    $hostArgs = New-Object 'System.Collections.Generic.List[string]'
    [void]$hostArgs.Add('-NoProfile')
    [void]$hostArgs.Add('-ExecutionPolicy')
    [void]$hostArgs.Add('Bypass')
    [void]$hostArgs.Add('-File')
    [void]$hostArgs.Add($wrapperPath)

    return @{
        HostArgs = @($hostArgs)
        WrapperPath = $wrapperPath
    }
}

function Update-SummaryState {
    $selectedIds = Get-SelectedTaskIds
    $engineName = Split-Path (Get-PreferredPowerShellExe) -Leaf
    $lblSummaryInfo.Text = "Engine=$engineName / Mode=$($cmbMode.SelectedItem) / Profile=$($cmbProfile.SelectedItem) / Failure=$($cmbFailure.SelectedItem) / Tasks=$($selectedIds.Count)"
    $lblTaskCount.Text = "選択中: $($selectedIds.Count) / $($taskDefinitions.Count)"
    $lblModeBadge.Text = "Mode: $($cmbMode.SelectedItem) / $($cmbProfile.SelectedItem)"
    $lblStatusFoot.Text = "実行エンジン: $(Get-PreferredPowerShellExe)"
}

function Update-CommandPreview {
    $plan = New-ExecutionArguments
    $engine = Get-PreferredPowerShellExe
    $txtCommand.Text = $engine + ' ' + (Convert-ToArgumentString -Arguments @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', '<generated gui wrapper script>'
    ))
    Update-SummaryState
}

function Write-LogLine {
    param(
        [Parameter(Mandatory)][string]$Text,
        [string]$Prefix = ''
    )

    $line = if ($Prefix) { "$Prefix$Text" } else { $Text }
    $txtLog.AppendText($line + [Environment]::NewLine)
    $txtLog.SelectionStart = $txtLog.TextLength
    $txtLog.ScrollToCaret()
}

function Update-ProgressFromLine {
    param([Parameter(Mandatory)][string]$Line)

    $isProgressLine = ($Line -match '^\s*\[')
    $isGuiMetaLine = ($Line -match '^\s*\[gui\]')
    $isTaskCompletionLine = ($Line -match '^\s*\[(WhatIf|Tasks|Diagnose)\]') -or ($Line -match '完了|失敗|スキップ|(?i:skip)')

    if ($isProgressLine -and -not $isGuiMetaLine) {
        $taskId = Get-TaskIdFromLine -Line $Line
        if ($null -ne $taskId) {
            $sync.CurrentTask = Get-TaskDisplayName -TaskId $taskId
        } else {
            $sync.CurrentTask = ([regex]::Replace($Line, '^\s*\[[^\]]+\]\s*', '')).Trim()
        }
        $lblCurrentTask.Text = '現在タスク: ' + $sync.CurrentTask
    }

    if ($isProgressLine -and $isTaskCompletionLine) {
        $taskId = Get-TaskIdFromLine -Line $Line
        if ($null -ne $taskId -and -not $sync.CompletedTaskIds.Contains($taskId)) {
            [void]$sync.CompletedTaskIds.Add($taskId)
            $sync.CompletedCount = [Math]::Min(($sync.CompletedCount + 1), $sync.SelectedIds.Count)
        }
    }

    $completed = $sync.CompletedCount
    $total = [Math]::Max($sync.SelectedIds.Count, 1)
    $progress.Value = [Math]::Min([int](($completed / $total) * 100), 100)
    $lblProgress.Text = "進捗: $completed / $($sync.SelectedIds.Count)"
}

function Stop-ActiveProcess {
    if ($sync.Process -and -not $sync.Process.HasExited) {
        try {
            $sync.Process.Kill()
            $sync.StopRequested = $true
            Set-StatusBadge -Text '停止中' -BackColor $theme.Warning
        } catch {
            Write-LogLine -Text ("[gui] stop failed: " + $_.Exception.Message)
            Set-StatusBadge -Text '停止失敗' -BackColor $theme.Danger
        }
    }
}

function Clear-EventSubscriptions {
    foreach ($id in @($sync.OutputEventId, $sync.ErrorEventId, $sync.ExitEventId)) {
        if ($id) {
            Unregister-Event -SourceIdentifier $id -ErrorAction SilentlyContinue
            Remove-Job -Name $id -Force -ErrorAction SilentlyContinue
        }
    }
    $sync.OutputEventId = $null
    $sync.ErrorEventId = $null
    $sync.ExitEventId = $null
}

function Remove-WrapperScript {
    if ($sync.WrapperPath -and (Test-Path -LiteralPath $sync.WrapperPath)) {
        Remove-Item -LiteralPath $sync.WrapperPath -Force -ErrorAction SilentlyContinue
    }
    $sync.WrapperPath = ''
}

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 250
$timer.Add_Tick({
    while ($true) {
        $item = $null
        if (-not $sync.Queue.TryDequeue([ref]$item)) { break }

        switch ($item.Kind) {
            'stdout' {
                $cleanText = Remove-AnsiEscapeSequence -Text $item.Text
                if ($cleanText) {
                    Write-LogLine -Text $cleanText
                    Update-ProgressFromLine -Line $cleanText
                }
            }
            'stderr' {
                $cleanText = Remove-AnsiEscapeSequence -Text $item.Text
                if ($cleanText) {
                    Write-LogLine -Text $cleanText -Prefix '[stderr] '
                }
            }
            'exit' {
                $sync.Running = $false
                $sync.ExitCode = [int]$item.ExitCode
                Clear-EventSubscriptions
                Remove-WrapperScript
                Find-LatestArtifacts
                Update-ArtifactLabels
                $lblExitCode.Text = "終了コード: $($sync.ExitCode)"
                if ($sync.StopRequested) {
                    Set-StatusBadge -Text '停止' -BackColor $theme.Warning
                } elseif ($sync.ExitCode -eq 0) {
                    Set-StatusBadge -Text '完了' -BackColor $theme.Success
                } else {
                    Set-StatusBadge -Text '失敗' -BackColor $theme.Danger
                }
                Set-ControlsEnabled -Enabled $true
            }
        }
    }

    if ($sync.Running -and $sync.StartTime) {
        $elapsed = (Get-Date) - $sync.StartTime
        $lblElapsed.Text = '経過時間: ' + $elapsed.ToString('hh\:mm\:ss')
    }
})
$timer.Start()

$btnBrowseConfig.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = 'JSON files (*.json)|*.json|All files (*.*)|*.*'
    if ($txtConfig.Text -and (Test-Path -LiteralPath (Split-Path $txtConfig.Text -Parent))) {
        $dialog.InitialDirectory = (Split-Path $txtConfig.Text -Parent)
    }
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtConfig.Text = $dialog.FileName
    }
    $dialog.Dispose()
})

$btnBrowseExport.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($txtExportPath.Text -and (Test-Path -LiteralPath $txtExportPath.Text)) {
        $dialog.SelectedPath = $txtExportPath.Text
    }
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtExportPath.Text = $dialog.SelectedPath
    }
    $dialog.Dispose()
})

$btnSelectAll.Add_Click({
    for ($i = 0; $i -lt $taskList.Items.Count; $i++) {
        $taskList.SetItemChecked($i, $true)
    }
    Update-CommandPreview
})

$btnClearAll.Add_Click({
    for ($i = 0; $i -lt $taskList.Items.Count; $i++) {
        $taskList.SetItemChecked($i, $false)
    }
    Update-CommandPreview
})

$btnDiagnosePreset.Add_Click({
    for ($i = 0; $i -lt $taskList.Items.Count; $i++) {
        $taskList.SetItemChecked($i, $false)
    }
    foreach ($taskId in @(14, 15, 16, 20)) {
        $taskList.SetItemChecked(($taskId - 1), $true)
    }
    $cmbMode.SelectedItem = 'diagnose'
    Update-CommandPreview
})

foreach ($control in @($cmbMode, $cmbProfile, $cmbFailure, $chkWhatIf, $chkUseLocalChart, $chkExportPowerBI, $chkUseAnthropic, $chkExportDeleted, $cmbExportFormat, $txtConfig, $txtExportPath, $taskList)) {
    if ($control -is [System.Windows.Forms.TextBox]) {
        $control.Add_TextChanged({ Update-CommandPreview })
    } elseif ($control -is [System.Windows.Forms.CheckedListBox]) {
        $control.Add_ItemCheck({
            $form.BeginInvoke([System.Action]{ Update-CommandPreview }) | Out-Null
        })
    } elseif ($control -is [System.Windows.Forms.CheckBox]) {
        $control.Add_CheckedChanged({ Update-CommandPreview })
    } elseif ($control -is [System.Windows.Forms.ComboBox]) {
        $control.Add_SelectedIndexChanged({ Update-CommandPreview })
    }
}

$btnOpenReports.Add_Click({ Open-PathInShell -Path $reportsDir })
$btnOpenLogs.Add_Click({ Open-PathInShell -Path $logsDir })
$btnOpenDocs.Add_Click({ Open-PathInShell -Path $docsGuiDir })
$btnOpenRepo.Add_Click({ Open-PathInShell -Path $repoRoot })
$btnOpenLatestHtml.Add_Click({
    if ($sync.LatestHtml) {
        Start-Process -FilePath $sync.LatestHtml | Out-Null
    }
})
$btnOpenLatestJson.Add_Click({
    if ($sync.LatestJson) {
        Start-Process -FilePath $sync.LatestJson | Out-Null
    }
})

$btnStop.Add_Click({
    Stop-ActiveProcess
})

$btnRun.Add_Click({
    if ($sync.Running) { return }

    $plan = New-ExecutionArguments
    if ($plan.SelectedIds.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            '少なくとも 1 つのタスクを選択してください。',
            'PC Optimizer GUI'
        ) | Out-Null
        return
    }

    if ($chkExportDeleted.Checked -and -not $chkWhatIf.Checked) {
        [System.Windows.Forms.MessageBox]::Show(
            'ExportDeletedPaths を使う場合は WhatIf を有効にしてください。',
            'PC Optimizer GUI'
        ) | Out-Null
        return
    }

    $engine = Get-PreferredPowerShellExe
    Remove-WrapperScript
    $hostInvocation = New-HostInvocation -EnginePath $engine -BackendArgs $plan.Args
    $hostArgs = $hostInvocation.HostArgs
    $sync.WrapperPath = $hostInvocation.WrapperPath
    $sync.Queue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    $sync.SelectedIds = $plan.SelectedIds
    $sync.CompletedCount = 0
    $sync.CompletedTaskIds = New-Object 'System.Collections.Generic.HashSet[int]'
    $sync.StopRequested = $false
    $sync.ExitCode = $null
    $sync.StartTime = Get-Date
    $sync.Running = $true
    $sync.CurrentTask = ''

    $progress.Value = 0
    $lblProgress.Text = "進捗: 0 / $($sync.SelectedIds.Count)"
    $lblCurrentTask.Text = '現在タスク: -'
    $lblElapsed.Text = '経過時間: 00:00:00'
    $lblExitCode.Text = '終了コード: -'
    $txtLog.Clear()
    Set-StatusBadge -Text '実行中' -BackColor $theme.Accent
    Write-LogLine -Text ("[gui] run started: " + (Get-Date -Format 'yyyy/MM/dd HH:mm:ss'))
    Write-LogLine -Text ("[gui] command: " + $engine + ' ' + (Convert-ToArgumentString -Arguments $hostArgs))

    if (($plan.SelectedIds -contains 18) -or ($plan.SelectedIds -contains 19)) {
        Write-LogLine -Text '[gui] note: Task 18 / 19 は NonInteractive 実行のため更新確認だけ行い、既定応答 N でスキップされます。'
    }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $engine
    $psi.Arguments = Convert-ToArgumentString -Arguments $hostArgs
    $psi.WorkingDirectory = $repoRoot
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    $psi.StandardOutputEncoding = $utf8NoBom
    $psi.StandardErrorEncoding = $utf8NoBom

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $psi
    $process.EnableRaisingEvents = $true
    $sync.Process = $process

    $runId = [guid]::NewGuid().ToString('N')
    $sync.OutputEventId = "pcopt-gui-out-$runId"
    $sync.ErrorEventId = "pcopt-gui-err-$runId"
    $sync.ExitEventId = "pcopt-gui-exit-$runId"

    $null = Register-ObjectEvent -InputObject $process -EventName OutputDataReceived -SourceIdentifier $sync.OutputEventId -Action {
        if ($EventArgs.Data) {
            $syncHash = $Event.MessageData
            $syncHash.Queue.Enqueue([PSCustomObject]@{ Kind = 'stdout'; Text = $EventArgs.Data })
        }
    } -MessageData $sync

    $null = Register-ObjectEvent -InputObject $process -EventName ErrorDataReceived -SourceIdentifier $sync.ErrorEventId -Action {
        if ($EventArgs.Data) {
            $syncHash = $Event.MessageData
            $syncHash.Queue.Enqueue([PSCustomObject]@{ Kind = 'stderr'; Text = $EventArgs.Data })
        }
    } -MessageData $sync

    $null = Register-ObjectEvent -InputObject $process -EventName Exited -SourceIdentifier $sync.ExitEventId -Action {
        $syncHash = $Event.MessageData
        $eventProcess = [System.Diagnostics.Process]$Event.Sender
        $syncHash.Queue.Enqueue([PSCustomObject]@{ Kind = 'exit'; ExitCode = $eventProcess.ExitCode })
    } -MessageData $sync

    try {
        if (-not $process.Start()) {
            throw 'Process start failed.'
        }
        $process.BeginOutputReadLine()
        $process.BeginErrorReadLine()
        Set-ControlsEnabled -Enabled $false
    } catch {
        Clear-EventSubscriptions
        Remove-WrapperScript
        $sync.Running = $false
        [System.Windows.Forms.MessageBox]::Show(
            "実行開始に失敗しました。`r`n$_",
            'PC Optimizer GUI'
        ) | Out-Null
        Set-ControlsEnabled -Enabled $true
        Set-StatusBadge -Text '失敗' -BackColor $theme.Danger
    }
})

$form.Add_FormClosing({
    Stop-ActiveProcess
    Clear-EventSubscriptions
    Remove-WrapperScript
    if ($heroImage.Image) {
        $heroImage.Image.Dispose()
    }
})

Find-LatestArtifacts
Update-ArtifactLabels
Set-StatusBadge -Text '待機中' -BackColor $theme.HeroSecondary
Update-CommandPreview
[void]$form.ShowDialog()
