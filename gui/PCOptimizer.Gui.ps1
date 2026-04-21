Set-StrictMode -Version Latest

function Test-GuiAdministrator {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        return $false
    }
}

function Get-HostExecutable {
    $pwsh = Get-Command pwsh -ErrorAction SilentlyContinue
    if ($pwsh) { return $pwsh.Source }
    return (Get-Command powershell.exe -ErrorAction Stop).Source
}

function Restart-InStaIfNeeded {
    if ([Threading.Thread]::CurrentThread.GetApartmentState() -eq [Threading.ApartmentState]::STA) {
        return
    }

    $hostExe = Get-HostExecutable
    $args = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-File', "`"$PSCommandPath`"")
    Start-Process -FilePath $hostExe -ArgumentList $args | Out-Null
    exit 0
}

function Restart-ElevatedIfNeeded {
    if (Test-GuiAdministrator) { return }

    $hostExe = Get-HostExecutable
    $args = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-File', "`"$PSCommandPath`"")
    Start-Process -FilePath $hostExe -ArgumentList $args -Verb RunAs | Out-Null
    exit 0
}

function ConvertTo-ProcessArgumentString {
    param([string[]]$Arguments)

    @(
        $Arguments | ForEach-Object {
            if ($_ -match '[\s"]') {
                '"' + ($_.Replace('"', '\"')) + '"'
            } else {
                $_
            }
        }
    ) -join ' '
}

function Get-ComboText {
    param($ComboBox)

    if ($ComboBox.SelectedItem -is [System.Windows.Controls.ComboBoxItem]) {
        return "$($ComboBox.SelectedItem.Content)"
    }
    return "$($ComboBox.Text)"
}

function Resolve-UiText {
    param([string]$Text)
    return [System.Net.WebUtility]::HtmlDecode($Text)
}

function Escape-PowerShellSingleQuoted {
    param([string]$Value)
    return "'" + $Value.Replace("'", "''") + "'"
}

function ConvertTo-PowerShellArrayLiteral {
    param([string[]]$Arguments)

    $items = New-Object 'System.Collections.Generic.List[string]'
    for ($i = 0; $i -lt @($Arguments).Count; $i++) {
        $arg = $Arguments[$i]
        if ($arg -eq '-EnableAIDiagnosis' -and ($i + 1) -lt @($Arguments).Count) {
            [void]$items.Add((Escape-PowerShellSingleQuoted -Value $arg))
            $boolValue = $Arguments[$i + 1]
            if ($boolValue -match '^(?i:true|false)$') {
                [void]$items.Add(('$' + $boolValue.ToLowerInvariant()))
            } else {
                [void]$items.Add((Escape-PowerShellSingleQuoted -Value $boolValue))
            }
            $i++
            continue
        }

        [void]$items.Add((Escape-PowerShellSingleQuoted -Value $arg))
    }

    return '@(' + ($items -join ', ') + ')'
}

function ConvertTo-PowerShellInvocationArguments {
    param([string[]]$Arguments)

    $items = New-Object 'System.Collections.Generic.List[string]'
    for ($i = 0; $i -lt @($Arguments).Count; $i++) {
        $arg = $Arguments[$i]
        if ($arg -match '^-[A-Za-z]') {
            if ($arg -eq '-EnableAIDiagnosis' -and ($i + 1) -lt @($Arguments).Count) {
                $boolValue = $Arguments[$i + 1]
                if ($boolValue -match '^(?i:true|false)$') {
                    [void]$items.Add(('{0}:${1}' -f $arg, $boolValue.ToLowerInvariant()))
                } else {
                    [void]$items.Add($arg)
                    [void]$items.Add((Escape-PowerShellSingleQuoted -Value $boolValue))
                }
                $i++
                continue
            }

            [void]$items.Add($arg)
            continue
        }

        [void]$items.Add((Escape-PowerShellSingleQuoted -Value $arg))
    }

    return ($items -join ' ')
}

function Append-OutputLine {
    param(
        [Parameter(Mandatory)][hashtable]$Sync,
        [Parameter(Mandatory)][string]$Line
    )

    if ([string]::IsNullOrWhiteSpace($Line)) { return }

    if ($Line.StartsWith($Sync.EventPrefix, [System.StringComparison]::Ordinal)) {
        $payload = $Line.Substring($Sync.EventPrefix.Length)
        try {
            $evt = $payload | ConvertFrom-Json
            switch ("$($evt.type)") {
                'run_start' {
                    $Sync.StatusText.Text = ('{0}: {1} / {2}' -f (Resolve-UiText '&#x5B9F;&#x884C;&#x958B;&#x59CB;'), $evt.mode, $evt.executionProfile)
                }
                'task_start' {
                    $Sync.StatusText.Text = ('{0}: #{1} {2}' -f (Resolve-UiText '&#x5B9F;&#x884C;&#x4E2D;'), $evt.taskId, $evt.taskName)
                }
                'task_finish' {
                    $Sync.CompletedTasks = [Math]::Min($Sync.SelectedTaskCount, ($Sync.CompletedTasks + 1))
                    $Sync.ProgressBar.Value = $Sync.CompletedTasks
                    $Sync.ProgressText.Text = "{0} / {1}" -f $Sync.CompletedTasks, $Sync.SelectedTaskCount
                    $Sync.StatusText.Text = ('{0}: #{1} {2}' -f (Resolve-UiText '&#x5B8C;&#x4E86;'), $evt.taskId, $evt.status)
                }
                'report_generated' {
                    if ($evt.htmlReportPath) {
                        $Sync.LatestReportPath = "$($evt.htmlReportPath)"
                        $Sync.LatestReportText.Text = "$($evt.htmlReportPath)"
                        $Sync.OpenReportButton.IsEnabled = $true
                    }
                }
                'run_complete' {
                    $Sync.LastExitCode = [int]$evt.exitCode
                    $Sync.StatusText.Text = ('{0}: {1}={2} {3}={4}' -f (Resolve-UiText '&#x5B9F;&#x884C;&#x5B8C;&#x4E86;'), (Resolve-UiText '&#x72B6;&#x614B;'), $evt.status, (Resolve-UiText '&#x7D42;&#x4E86;&#x30B3;&#x30FC;&#x30C9;'), $evt.exitCode)
                    $Sync.RunButton.IsEnabled = $true
                }
            }
        } catch {
            $Sync.OutputTextBox.AppendText("[" + (Resolve-UiText '&#x0055;&#x0049;&#x30A4;&#x30D9;&#x30F3;&#x30C8;&#x89E3;&#x6790;&#x30A8;&#x30E9;&#x30FC;') + "] " + $payload + [Environment]::NewLine)
            $Sync.OutputTextBox.ScrollToEnd()
        }
        return
    }

    Update-FromEngineLogLine -Sync $Sync -Line $Line
    $Sync.OutputTextBox.AppendText($Line + [Environment]::NewLine)
    $Sync.OutputTextBox.ScrollToEnd()
}

function Append-PrefixedOutputLine {
    param(
        [Parameter(Mandatory)][hashtable]$Sync,
        [Parameter(Mandatory)][string]$Prefix,
        [Parameter(Mandatory)][string]$Line
    )

    if ([string]::IsNullOrWhiteSpace($Line)) { return }
    $Sync.OutputTextBox.AppendText(('{0} {1}' -f $Prefix, $Line) + [Environment]::NewLine)
    $Sync.OutputTextBox.ScrollToEnd()
}

function Update-ProgressDisplay {
    param([hashtable]$Sync)

    $Sync.ProgressBar.Value = [Math]::Min($Sync.ProgressBar.Maximum, $Sync.CompletedTasks)
    $Sync.ProgressText.Text = "{0} / {1}" -f $Sync.CompletedTasks, $Sync.SelectedTaskCount
}

function Sync-RunStateFromEngineLog {
    param([hashtable]$Sync)

    if ([string]::IsNullOrWhiteSpace($Sync.EngineLogPath)) { return }
    if (-not (Test-Path -LiteralPath $Sync.EngineLogPath)) { return }

    $lines = @(Get-Content -LiteralPath $Sync.EngineLogPath -Encoding UTF8 -ErrorAction SilentlyContinue)
    if (@($lines).Count -eq 0) { return }

    $completedCount = @(
        $lines | Where-Object {
            $line = "$_"
            $line.Contains('[WhatIf]') -or
            $line.Contains('[Tasks]')
        }
    ).Count

    $Sync.CompletedTasks = [Math]::Min($Sync.SelectedTaskCount, $completedCount)
    Update-ProgressDisplay -Sync $Sync

    $latestReport = @(
        $lines | Where-Object { ("$_").Contains('[HTML') } | Select-Object -Last 1
    )
    if ($latestReport) {
        $reportPath = ("$latestReport" -replace '^\[[^\]]+\]\s+保存完了:\s+', '').Trim()
        if ($reportPath) {
            $Sync.LatestReportPath = $reportPath
            $Sync.LatestReportText.Text = $reportPath
            $Sync.OpenReportButton.IsEnabled = (Test-Path -LiteralPath $reportPath)
        }
    }

    if (@($lines | Where-Object { ("$_").Contains('[NonInteractive]') }).Count -gt 0) {
        $Sync.CompletedTasks = $Sync.SelectedTaskCount
        Update-ProgressDisplay -Sync $Sync
        $Sync.StatusText.Text = [System.Net.WebUtility]::HtmlDecode('&#x5B9F;&#x884C;&#x5B8C;&#x4E86;')
    }
}

function Complete-TrackedTask {
    param(
        [hashtable]$Sync,
        [string]$TaskKey
    )

    if ([string]::IsNullOrWhiteSpace($TaskKey)) { return }
    if ($Sync.CompletedTaskSet.Contains($TaskKey)) { return }
    [void]$Sync.CompletedTaskSet.Add($TaskKey)
    $Sync.CompletedTasks = [Math]::Min($Sync.SelectedTaskCount, ($Sync.CompletedTasks + 1))
    Update-ProgressDisplay -Sync $Sync
}

function Update-FromEngineLogLine {
    param(
        [hashtable]$Sync,
        [string]$Line
    )

    $textRunning = [System.Net.WebUtility]::HtmlDecode('&#x5B9F;&#x884C;&#x4E2D;')
    $textWhatIfDone = [System.Net.WebUtility]::HtmlDecode('&#x30D7;&#x30EC;&#x30D3;&#x30E5;&#x30FC;&#x5B8C;&#x4E86;')
    $textSkip = [System.Net.WebUtility]::HtmlDecode('&#x30B9;&#x30AD;&#x30C3;&#x30D7;')
    $textSkipReason = [System.Net.WebUtility]::HtmlDecode('&#x5BFE;&#x8C61;&#x5916;&#x306E;&#x305F;&#x3081;&#x30B9;&#x30AD;&#x30C3;&#x30D7;')
    $textDone = [System.Net.WebUtility]::HtmlDecode('&#x5B8C;&#x4E86;')
    $textDoneMemo = [System.Net.WebUtility]::HtmlDecode('&#x51E6;&#x7406;&#x5B8C;&#x4E86;')
    $textReportReady = [System.Net.WebUtility]::HtmlDecode('&#x6700;&#x65B0;&#x30EC;&#x30DD;&#x30FC;&#x30C8;&#x3092;&#x751F;&#x6210;&#x3057;&#x307E;&#x3057;&#x305F;&#x3002;')
    $textScore = [System.Net.WebUtility]::HtmlDecode('&#x8A3A;&#x65AD;&#x30B9;&#x30B3;&#x30A2;')
    $textReportGenerating = [System.Net.WebUtility]::HtmlDecode('&#x30EC;&#x30DD;&#x30FC;&#x30C8;&#x3092;&#x751F;&#x6210;&#x3057;&#x3066;&#x3044;&#x307E;&#x3059;&#x3002;')
    $textRunComplete = [System.Net.WebUtility]::HtmlDecode('&#x5B9F;&#x884C;&#x5B8C;&#x4E86;')
    $textPcOptimizeDone = [System.Net.WebUtility]::HtmlDecode('&#x0050;&#x0043;&#x0020;&#x6700;&#x9069;&#x5316;&#x304C;&#x5B8C;&#x4E86;&#x3057;&#x307E;&#x3057;&#x305F;&#x3002;')
    $textCurrentRunning = [System.Net.WebUtility]::HtmlDecode('&#x73FE;&#x5728;&#x5B9F;&#x884C;&#x4E2D;')
    $textAllDone = [System.Net.WebUtility]::HtmlDecode('&#x3059;&#x3079;&#x3066;&#x306E;&#x51E6;&#x7406;&#x304C;&#x5B8C;&#x4E86;&#x3057;&#x307E;&#x3057;&#x305F;&#x3002;')

    if ($Line -match '^\[(\d{2}:\d{2}:\d{2})\]\s+(.+?)\s+開始\.\.\.$') {
        $taskName = $Matches[2]
        $Sync.StatusText.Text = ('{0}: {1}' -f $textRunning, $taskName)
        $Sync.SummaryMemoText.Text = ('{0}: {1}' -f $textCurrentRunning, $taskName)
        return
    }

    if ($Line -match '^\[WhatIf\]\s+(.+?)\s+はプレビュー実行のため変更をスキップしました。$') {
        $taskName = $Matches[1]
        Complete-TrackedTask -Sync $Sync -TaskKey $taskName
        $Sync.StatusText.Text = ('WhatIf: {0}' -f $taskName)
        $Sync.SummaryMemoText.Text = ('{0}: {1}' -f $textWhatIfDone, $taskName)
        return
    }

    if ($Line -match '^\[Tasks\]\s+対象外タスクのためスキップ:\s+#\d+\s+(.+)$') {
        $taskName = $Matches[1]
        Complete-TrackedTask -Sync $Sync -TaskKey $taskName
        $Sync.StatusText.Text = ('{0}: {1}' -f $textSkip, $taskName)
        $Sync.SummaryMemoText.Text = ('{0}: {1}' -f $textSkipReason, $taskName)
        return
    }

    if ($Line -match '^\[(\d{2}:\d{2}:\d{2})\]\s+(.+?)\s+完了$') {
        $taskName = $Matches[2]
        Complete-TrackedTask -Sync $Sync -TaskKey $taskName
        $Sync.StatusText.Text = ('{0}: {1}' -f $textDone, $taskName)
        $Sync.SummaryMemoText.Text = ('{0}: {1}' -f $textDoneMemo, $taskName)
        return
    }

    if ($Line -match '^\[HTMLレポート\]\s+保存完了:\s+(.+)$') {
        $reportPath = $Matches[1].Trim()
        $Sync.LatestReportPath = $reportPath
        $Sync.LatestReportText.Text = $reportPath
        $Sync.OpenReportButton.IsEnabled = (Test-Path -LiteralPath $reportPath)
        $Sync.SummaryMemoText.Text = $textReportReady
        return
    }

    if ($Line -match '^\[module\]\s+診断データ収集完了:\s+score=(\d+),\s+status=([A-Za-z]+)$') {
        $Sync.SummaryMemoText.Text = ('{0}: score={1} status={2}' -f $textScore, $Matches[1], $Matches[2])
        return
    }

    if ($Line -match 'HTML') {
        $Sync.SummaryMemoText.Text = $textReportGenerating
        return
    }

    if ($Line -match '^\[全タスク完了\]') {
        $Sync.CompletedTasks = $Sync.SelectedTaskCount
        Update-ProgressDisplay -Sync $Sync
        $Sync.StatusText.Text = $textRunComplete
        $Sync.SummaryMemoText.Text = $textAllDone
    }
}

function Read-NewFileLines {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][ref]$State
    )

    if (-not (Test-Path -LiteralPath $Path)) { return @() }

    $fs = $null
    try {
        $fs = [System.IO.FileStream]::new($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        if ($State.Value.Position -gt $fs.Length) {
            $State.Value.Position = 0L
            $State.Value.Remaining = ""
        }
        $fs.Seek($State.Value.Position, [System.IO.SeekOrigin]::Begin) | Out-Null
        $sr = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8, $true)
        $text = $sr.ReadToEnd()
        $State.Value.Position = $fs.Position
        $sr.Dispose()
        $fs = $null

        if ([string]::IsNullOrEmpty($text)) { return @() }

        $buffer = $State.Value.Remaining + $text
        $parts = $buffer -split "`r?`n", -1
        if ($buffer -match "(`r`n|`n)$") {
            $State.Value.Remaining = ""
            return @($parts | Where-Object { $_ -ne "" })
        }

        $State.Value.Remaining = $parts[-1]
        if ($parts.Count -le 1) { return @() }
        return @($parts[0..($parts.Count - 2)] | Where-Object { $_ -ne "" })
    } finally {
        if ($null -ne $fs) { $fs.Dispose() }
    }
}

function New-ExpectedEngineLogPaths {
    param(
        [Parameter(Mandatory)][datetime]$StartedAt
    )

    $minutes = @(
        $StartedAt.ToString('yyyyMMddHHmm'),
        $StartedAt.AddMinutes(1).ToString('yyyyMMddHHmm')
    ) | Select-Object -Unique

    return [PSCustomObject]@{
        Log = @($minutes | ForEach-Object { Join-Path $logsDir ("PC_Optimizer_Log_{0}.txt" -f $_) })
        Error = @($minutes | ForEach-Object { Join-Path $logsDir ("PC_Optimizer_Error_{0}.txt" -f $_) })
    }
}

Restart-InStaIfNeeded
Restart-ElevatedIfNeeded

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

$scriptDir = Split-Path -Parent $PSCommandPath
$repoRoot = Split-Path -Parent $scriptDir
$enginePath = Join-Path $repoRoot 'PC_Optimizer.ps1'
$configPath = Join-Path $repoRoot 'config\config.json'
$logsDir = Join-Path $repoRoot 'logs'

Import-Module (Join-Path $repoRoot 'modules\TaskCatalog.psm1') -Force

[xml]$xaml = Get-Content -Path (Join-Path $scriptDir 'MainWindow.xaml') -Raw -Encoding UTF8
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

$ModeCombo = $window.FindName('ModeCombo')
$ProfileCombo = $window.FindName('ProfileCombo')
$FailureCombo = $window.FindName('FailureCombo')
$WhatIfCheck = $window.FindName('WhatIfCheck')
$AICheck = $window.FindName('AICheck')
$ConfigPathText = $window.FindName('ConfigPathText')
$TasksListView = $window.FindName('TasksListView')
$SelectAllButton = $window.FindName('SelectAllButton')
$SelectNoneButton = $window.FindName('SelectNoneButton')
$SelectReadOnlyButton = $window.FindName('SelectReadOnlyButton')
$RunButton = $window.FindName('RunButton')
$OpenLogsButton = $window.FindName('OpenLogsButton')
$OpenReportButton = $window.FindName('OpenReportButton')
$RunProgressBar = $window.FindName('RunProgressBar')
$StatusText = $window.FindName('StatusText')
$ProgressText = $window.FindName('ProgressText')
$OutputTextBox = $window.FindName('OutputTextBox')
$EngineText = $window.FindName('EngineText')
$LatestReportText = $window.FindName('LatestReportText')
$RunMethodText = $window.FindName('RunMethodText')
$SummaryMemoText = $window.FindName('SummaryMemoText')

$ConfigPathText.Text = $configPath
$EngineText.Text = ('BAT 起動 / ホスト: {0}' -f (Get-HostExecutable))
$RunMethodText.Text = [System.Net.WebUtility]::HtmlDecode('&#x0042;&#x0041;&#x0054;&#x0020;&#x30D5;&#x30A1;&#x30A4;&#x30EB;&#x304B;&#x3089;&#x0020;&#x0047;&#x0055;&#x0049;&#x0020;&#x3092;&#x8D77;&#x52D5;&#x3057;&#x3001;&#x0050;&#x0043;&#x005F;&#x004F;&#x0070;&#x0074;&#x0069;&#x006D;&#x0069;&#x007A;&#x0065;&#x0072;&#x002E;&#x0070;&#x0073;&#x0031;&#x0020;&#x3092;&#x5B50;&#x30D7;&#x30ED;&#x30BB;&#x30B9;&#x3068;&#x3057;&#x3066;&#x5B9F;&#x884C;&#x3057;&#x307E;&#x3059;&#x3002;')

$taskItems = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
foreach ($task in @(Get-PCOptimizerTaskCatalog)) {
    $taskItems.Add([PSCustomObject]@{
        Id = [int]$task.Id
        Label = "$($task.Label)"
        Category = "$($task.Category)"
        IsReadOnly = [bool]$task.IsReadOnly
        ReadOnlyLabel = if ($task.IsReadOnly) { Resolve-UiText '&#x306F;&#x3044;' } else { Resolve-UiText '&#x3044;&#x3044;&#x3048;' }
        IsSelected = $true
    })
}
$TasksListView.ItemsSource = $taskItems

$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromMilliseconds(250)

$global:PCOptimizerGuiSync = [hashtable]::Synchronized(@{
    Window = $window
    OutputTextBox = $OutputTextBox
    StatusText = $StatusText
    ProgressText = $ProgressText
    ProgressBar = $RunProgressBar
    RunButton = $RunButton
    OpenReportButton = $OpenReportButton
    LatestReportText = $LatestReportText
    SummaryMemoText = $SummaryMemoText
    LatestReportPath = ''
    CompletedTasks = 0
    SelectedTaskCount = 0
    CompletedTaskSet = (New-Object 'System.Collections.Generic.HashSet[string]')
    EventPrefix = '##PCOPT_UI##'
    LastExitCode = $null
    Process = $null
    StdOutPath = ''
    StdOutState = @{ Position = 0L; Remaining = '' }
    EngineLogPath = ''
    EngineLogState = @{ Position = 0L; Remaining = '' }
    EngineErrorLogPath = ''
    EngineErrorLogState = @{ Position = 0L; Remaining = '' }
    ExpectedEngineLogPaths = @()
    ExpectedEngineErrorLogPaths = @()
    RunStartedAt = [datetime]::MinValue
    ExitHandled = $false
    Timer = $timer
    LauncherPath = ''
})

$timer.Add_Tick({
    $sync = $global:PCOptimizerGuiSync

    foreach ($line in @(Read-NewFileLines -Path $sync.StdOutPath -State ([ref]$sync.StdOutState))) {
        Append-OutputLine -Sync $sync -Line $line
    }

    if (-not $sync.EngineLogPath) {
        $sync.EngineLogPath = ($sync.ExpectedEngineLogPaths | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1)
    }
    if ($sync.EngineLogPath) {
        Sync-RunStateFromEngineLog -Sync $sync
        foreach ($line in @(Read-NewFileLines -Path $sync.EngineLogPath -State ([ref]$sync.EngineLogState))) {
            Update-FromEngineLogLine -Sync $sync -Line $line
            Append-PrefixedOutputLine -Sync $sync -Prefix '[ログ]' -Line $line
        }
    }

    if (-not $sync.EngineErrorLogPath) {
        $sync.EngineErrorLogPath = ($sync.ExpectedEngineErrorLogPaths | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1)
    }
    if ($sync.EngineErrorLogPath) {
        foreach ($line in @(Read-NewFileLines -Path $sync.EngineErrorLogPath -State ([ref]$sync.EngineErrorLogState))) {
            Append-PrefixedOutputLine -Sync $sync -Prefix '[エラー]' -Line $line
        }
    }

    $proc = $sync.Process
    if ($proc -and $proc.HasExited -and -not $sync.ExitHandled) {
        $sync.ExitHandled = $true
        $sync.Timer.Stop()
        Sync-RunStateFromEngineLog -Sync $sync
        if ($proc.ExitCode -eq 0 -and $sync.SelectedTaskCount -gt 0) {
            $sync.CompletedTasks = $sync.SelectedTaskCount
            Update-ProgressDisplay -Sync $sync
            $sync.LastExitCode = 0
            $sync.StatusText.Text = [System.Net.WebUtility]::HtmlDecode('&#x5B9F;&#x884C;&#x5B8C;&#x4E86;')
        } elseif ($null -eq $sync.LastExitCode) {
            $sync.StatusText.Text = ('{0}: {1}={2}' -f (Resolve-UiText '&#x30D7;&#x30ED;&#x30BB;&#x30B9;&#x7D42;&#x4E86;'), (Resolve-UiText '&#x7D42;&#x4E86;&#x30B3;&#x30FC;&#x30C9;'), $proc.ExitCode)
        }
        $sync.RunButton.IsEnabled = $true
        if ($sync.LatestReportPath -and (Test-Path -LiteralPath $sync.LatestReportPath)) {
            $sync.OpenReportButton.IsEnabled = $true
        }
    }
})

function Set-SelectionState {
    param([scriptblock]$Selector)

    foreach ($item in @($taskItems)) {
        $item.IsSelected = & $Selector $item
    }
    $TasksListView.Items.Refresh()
}

$SelectAllButton.Add_Click({
    Set-SelectionState { param($item) $true }
})

$SelectNoneButton.Add_Click({
    Set-SelectionState { param($item) $false }
})

$SelectReadOnlyButton.Add_Click({
    Set-SelectionState { param($item) [bool]$item.IsReadOnly }
})

$OpenLogsButton.Add_Click({
    if (-not (Test-Path $logsDir)) {
        New-Item -ItemType Directory -Path $logsDir -Force | Out-Null
    }
    Start-Process explorer.exe $logsDir | Out-Null
})

$OpenReportButton.Add_Click({
    $latest = $global:PCOptimizerGuiSync.LatestReportPath
    if ($latest -and (Test-Path $latest)) {
        Start-Process $latest | Out-Null
    } else {
        Start-Process explorer.exe $logsDir | Out-Null
    }
})

$window.Add_Closing({
    $sync = $global:PCOptimizerGuiSync
    if ($sync.Timer) {
        $sync.Timer.Stop()
    }
    $proc = if ($sync.ContainsKey('Process')) { $sync.Process } else { $null }
    if ($proc -and -not $proc.HasExited) {
        try { $proc.Kill() } catch {}
    }
})

$RunButton.Add_Click({
    $sync = $global:PCOptimizerGuiSync
    $currentProcess = if ($sync.ContainsKey('Process')) { $sync.Process } else { $null }
    if ($currentProcess -and -not $currentProcess.HasExited) {
        return
    }

    $selectedIds = @(
        $taskItems |
            Where-Object { $_.IsSelected } |
            ForEach-Object { [int]$_.Id }
    )
    if (@($selectedIds).Count -eq 0) {
        $StatusText.Text = Resolve-UiText '&#x5C11;&#x306A;&#x304F;&#x3068;&#x3082;&#x0020;&#x0031;&#x0020;&#x4EF6;&#x306E;&#x30BF;&#x30B9;&#x30AF;&#x3092;&#x9078;&#x629E;&#x3057;&#x3066;&#x304F;&#x3060;&#x3055;&#x3044;&#x3002;'
        return
    }

    if (-not (Test-Path $logsDir)) {
        New-Item -ItemType Directory -Path $logsDir -Force | Out-Null
    }

    $stamp = Get-Date -Format 'yyyyMMddHHmmss'
    $stdOutPath = Join-Path $logsDir ("gui_stdout_{0}.log" -f $stamp)
    New-Item -ItemType File -Path $stdOutPath -Force | Out-Null

    $sync.CompletedTasks = 0
    $sync.SelectedTaskCount = @($selectedIds).Count
    $sync.CompletedTaskSet = (New-Object 'System.Collections.Generic.HashSet[string]')
    $sync.LatestReportPath = ''
    $sync.LastExitCode = $null
    $sync.ExitHandled = $false
    $sync.StdOutPath = $stdOutPath
    $sync.StdOutState = @{ Position = 0L; Remaining = '' }
    $sync.EngineLogPath = ''
    $sync.EngineLogState = @{ Position = 0L; Remaining = '' }
    $sync.EngineErrorLogPath = ''
    $sync.EngineErrorLogState = @{ Position = 0L; Remaining = '' }
    $sync.RunStartedAt = Get-Date
    $expectedPaths = New-ExpectedEngineLogPaths -StartedAt $sync.RunStartedAt
    $sync.ExpectedEngineLogPaths = @($expectedPaths.Log)
    $sync.ExpectedEngineErrorLogPaths = @($expectedPaths.Error)
    $sync.LauncherPath = Join-Path $logsDir ("gui_launch_{0}.ps1" -f $stamp)
    $OpenReportButton.IsEnabled = $false
    $LatestReportText.Text = Resolve-UiText '&#x30EC;&#x30DD;&#x30FC;&#x30C8;&#x751F;&#x6210;&#x5F85;&#x3061;'
    $RunProgressBar.Minimum = 0
    $RunProgressBar.Maximum = [Math]::Max(1, @($selectedIds).Count)
    $RunProgressBar.Value = 0
    $ProgressText.Text = "0 / $(@($selectedIds).Count)"
    $StatusText.Text = Resolve-UiText '&#x5B9F;&#x884C;&#x6E96;&#x5099;&#x4E2D;'
    $SummaryMemoText.Text = Resolve-UiText '&#x5B9F;&#x884C;&#x958B;&#x59CB;&#x3092;&#x6E96;&#x5099;&#x4E2D;&#x3067;&#x3059;&#x3002;'
    $OutputTextBox.Clear()
    Append-PrefixedOutputLine -Sync $sync -Prefix '[GUI]' -Line (Resolve-UiText '&#x5B50;&#x30D7;&#x30ED;&#x30BB;&#x30B9;&#x3092;&#x8D77;&#x52D5;&#x3057;&#x307E;&#x3059;&#x3002;&#x30ED;&#x30B0;&#x30D5;&#x30A1;&#x30A4;&#x30EB;&#x3092;&#x76E3;&#x8996;&#x3057;&#x3066;&#x8868;&#x793A;&#x3057;&#x307E;&#x3059;&#x3002;')
    $RunButton.IsEnabled = $false

    $engineArgs = @(
        New-PCOptimizerArgumentList `
            -Mode (Get-ComboText -ComboBox $ModeCombo) `
            -ExecutionProfile (Get-ComboText -ComboBox $ProfileCombo) `
            -FailureMode (Get-ComboText -ComboBox $FailureCombo) `
            -TaskIds $selectedIds `
            -ConfigPath $ConfigPathText.Text `
            -EnableAIDiagnosis ([bool]$AICheck.IsChecked) `
            -WhatIfMode:([bool]$WhatIfCheck.IsChecked) `
            -NonInteractive `
            -NoRebootPrompt `
            -EmitUiEvents
    )

    $hostExe = Get-HostExecutable
    $quotedEnginePath = Escape-PowerShellSingleQuoted -Value $enginePath
    $quotedOutputPath = Escape-PowerShellSingleQuoted -Value $stdOutPath
    $quotedArgs = ConvertTo-PowerShellInvocationArguments -Arguments $engineArgs
    $launcherScript = @"
`$ErrorActionPreference = 'Continue'
`$OutputEncoding = [System.Text.UTF8Encoding]::new(`$false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new(`$false)
trap {
    (`$_ | Out-String) | Add-Content -LiteralPath $quotedOutputPath -Encoding UTF8
    continue
}
& $quotedEnginePath $quotedArgs *>> $quotedOutputPath
"@
    Set-Content -LiteralPath $sync.LauncherPath -Value $launcherScript -Encoding UTF8

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $hostExe
    $psi.Arguments = '-NoProfile -ExecutionPolicy Bypass -File ' + (ConvertTo-ProcessArgumentString -Arguments @($sync.LauncherPath))
    $psi.WorkingDirectory = $repoRoot
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    [void]$proc.Start()
    $sync.Process = $proc
    $sync.Timer.Start()
})

$window.ShowDialog() | Out-Null
