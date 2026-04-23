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

$taskDefinitions = @(
    [PSCustomObject]@{ Id = 1;  Name = 'Temp cleanup' },
    [PSCustomObject]@{ Id = 2;  Name = 'Prefetch and update cache cleanup' },
    [PSCustomObject]@{ Id = 3;  Name = 'Delivery optimization cleanup' },
    [PSCustomObject]@{ Id = 4;  Name = 'Windows Update cache cleanup' },
    [PSCustomObject]@{ Id = 5;  Name = 'Error report and log cleanup' },
    [PSCustomObject]@{ Id = 6;  Name = 'OneDrive Teams Office cache cleanup' },
    [PSCustomObject]@{ Id = 7;  Name = 'Browser cache cleanup' },
    [PSCustomObject]@{ Id = 8;  Name = 'Thumbnail cache cleanup' },
    [PSCustomObject]@{ Id = 9;  Name = 'Microsoft Store cache cleanup' },
    [PSCustomObject]@{ Id = 10; Name = 'Recycle Bin cleanup' },
    [PSCustomObject]@{ Id = 11; Name = 'DNS cache clear' },
    [PSCustomObject]@{ Id = 12; Name = 'Event log clear' },
    [PSCustomObject]@{ Id = 13; Name = 'Disk optimize' },
    [PSCustomObject]@{ Id = 14; Name = 'SSD health check' },
    [PSCustomObject]@{ Id = 15; Name = 'SFC repair scan' },
    [PSCustomObject]@{ Id = 16; Name = 'DISM scan' },
    [PSCustomObject]@{ Id = 17; Name = 'Power plan optimize' },
    [PSCustomObject]@{ Id = 18; Name = 'Microsoft 365 update check' },
    [PSCustomObject]@{ Id = 19; Name = 'Windows Update check' },
    [PSCustomObject]@{ Id = 20; Name = 'Startup service report' }
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
    return "Task $TaskId"
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

    $hostPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $elevationArgs = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', $ScriptPath,
        '-Elevated'
    )

    try {
        Start-Process -FilePath $hostPath -ArgumentList (Convert-ToArgumentString -Arguments $elevationArgs) -Verb RunAs | Out-Null
        exit
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to restart with administrator privileges.`r`n$_",
            'PC Optimizer GUI',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        exit 1
    }
}

function Get-PreferredPowerShellExe {
    return (Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe')
}

if (-not (Test-Path -LiteralPath $mainScriptPath)) {
    [System.Windows.Forms.MessageBox]::Show(
        "Backend script was not found.`r`n$mainScriptPath",
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
        'Unable to acquire administrator privileges.',
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
})

$form = New-Object System.Windows.Forms.Form
$form.Text = 'PC Optimizer GUI'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(1200, 780)
$form.MinimumSize = New-Object System.Drawing.Size(1024, 700)

$lblEngine = New-Object System.Windows.Forms.Label
$lblEngine.Location = New-Object System.Drawing.Point(12, 12)
$lblEngine.Size = New-Object System.Drawing.Size(560, 20)
$lblEngine.Text = "Engine: $(Get-PreferredPowerShellExe)"

$lblScript = New-Object System.Windows.Forms.Label
$lblScript.Location = New-Object System.Drawing.Point(12, 34)
$lblScript.Size = New-Object System.Drawing.Size(880, 20)
$lblScript.Text = "Backend: $mainScriptPath"

$grpSettings = New-Object System.Windows.Forms.GroupBox
$grpSettings.Text = 'Execution Settings'
$grpSettings.Location = New-Object System.Drawing.Point(12, 62)
$grpSettings.Size = New-Object System.Drawing.Size(560, 205)

$lblMode = New-Object System.Windows.Forms.Label
$lblMode.Location = New-Object System.Drawing.Point(16, 28)
$lblMode.Size = New-Object System.Drawing.Size(110, 20)
$lblMode.Text = 'Mode'

$cmbMode = New-Object System.Windows.Forms.ComboBox
$cmbMode.Location = New-Object System.Drawing.Point(140, 24)
$cmbMode.Size = New-Object System.Drawing.Size(140, 24)
$cmbMode.DropDownStyle = 'DropDownList'
[void]$cmbMode.Items.AddRange(@('repair', 'diagnose'))
$cmbMode.SelectedIndex = 0

$lblProfile = New-Object System.Windows.Forms.Label
$lblProfile.Location = New-Object System.Drawing.Point(300, 28)
$lblProfile.Size = New-Object System.Drawing.Size(110, 20)
$lblProfile.Text = 'ExecutionProfile'

$cmbProfile = New-Object System.Windows.Forms.ComboBox
$cmbProfile.Location = New-Object System.Drawing.Point(420, 24)
$cmbProfile.Size = New-Object System.Drawing.Size(120, 24)
$cmbProfile.DropDownStyle = 'DropDownList'
[void]$cmbProfile.Items.AddRange(@('classic', 'agent-teams'))
$cmbProfile.SelectedIndex = 0

$lblFailure = New-Object System.Windows.Forms.Label
$lblFailure.Location = New-Object System.Drawing.Point(16, 60)
$lblFailure.Size = New-Object System.Drawing.Size(110, 20)
$lblFailure.Text = 'FailureMode'

$cmbFailure = New-Object System.Windows.Forms.ComboBox
$cmbFailure.Location = New-Object System.Drawing.Point(140, 56)
$cmbFailure.Size = New-Object System.Drawing.Size(140, 24)
$cmbFailure.DropDownStyle = 'DropDownList'
[void]$cmbFailure.Items.AddRange(@('continue', 'fail-fast'))
$cmbFailure.SelectedIndex = 0

$chkWhatIf = New-Object System.Windows.Forms.CheckBox
$chkWhatIf.Location = New-Object System.Drawing.Point(300, 58)
$chkWhatIf.Size = New-Object System.Drawing.Size(100, 20)
$chkWhatIf.Text = 'WhatIf'

$chkUseLocalChart = New-Object System.Windows.Forms.CheckBox
$chkUseLocalChart.Location = New-Object System.Drawing.Point(420, 58)
$chkUseLocalChart.Size = New-Object System.Drawing.Size(120, 20)
$chkUseLocalChart.Text = 'UseLocalChartJs'
$chkUseLocalChart.Checked = $true

$chkExportPowerBI = New-Object System.Windows.Forms.CheckBox
$chkExportPowerBI.Location = New-Object System.Drawing.Point(16, 88)
$chkExportPowerBI.Size = New-Object System.Drawing.Size(140, 20)
$chkExportPowerBI.Text = 'ExportPowerBIJson'

$chkUseAnthropic = New-Object System.Windows.Forms.CheckBox
$chkUseAnthropic.Location = New-Object System.Drawing.Point(170, 88)
$chkUseAnthropic.Size = New-Object System.Drawing.Size(120, 20)
$chkUseAnthropic.Text = 'UseAnthropicAI'

$chkExportDeleted = New-Object System.Windows.Forms.CheckBox
$chkExportDeleted.Location = New-Object System.Drawing.Point(300, 88)
$chkExportDeleted.Size = New-Object System.Drawing.Size(140, 20)
$chkExportDeleted.Text = 'ExportDeletedPaths'

$cmbExportFormat = New-Object System.Windows.Forms.ComboBox
$cmbExportFormat.Location = New-Object System.Drawing.Point(450, 84)
$cmbExportFormat.Size = New-Object System.Drawing.Size(90, 24)
$cmbExportFormat.DropDownStyle = 'DropDownList'
[void]$cmbExportFormat.Items.AddRange(@('json', 'csv'))
$cmbExportFormat.SelectedIndex = 0

$lblConfig = New-Object System.Windows.Forms.Label
$lblConfig.Location = New-Object System.Drawing.Point(16, 120)
$lblConfig.Size = New-Object System.Drawing.Size(110, 20)
$lblConfig.Text = 'ConfigPath'

$txtConfig = New-Object System.Windows.Forms.TextBox
$txtConfig.Location = New-Object System.Drawing.Point(140, 116)
$txtConfig.Size = New-Object System.Drawing.Size(400, 24)
$txtConfig.Text = (Join-Path $repoRoot 'config\config.json')

$lblExportPath = New-Object System.Windows.Forms.Label
$lblExportPath.Location = New-Object System.Drawing.Point(16, 152)
$lblExportPath.Size = New-Object System.Drawing.Size(110, 20)
$lblExportPath.Text = 'ExportPath'

$txtExportPath = New-Object System.Windows.Forms.TextBox
$txtExportPath.Location = New-Object System.Drawing.Point(140, 148)
$txtExportPath.Size = New-Object System.Drawing.Size(400, 24)
$txtExportPath.Text = (Join-Path $repoRoot 'logs\gui-deletedpaths')

$lblNote = New-Object System.Windows.Forms.Label
$lblNote.Location = New-Object System.Drawing.Point(16, 178)
$lblNote.Size = New-Object System.Drawing.Size(520, 20)
$lblNote.Text = 'Note: GUI runs backend with -NonInteractive -NoRebootPrompt.'

$grpSettings.Controls.AddRange(@(
    $lblMode, $cmbMode, $lblProfile, $cmbProfile, $lblFailure, $cmbFailure,
    $chkWhatIf, $chkUseLocalChart, $chkExportPowerBI, $chkUseAnthropic,
    $chkExportDeleted, $cmbExportFormat, $lblConfig, $txtConfig,
    $lblExportPath, $txtExportPath, $lblNote
))

$grpTasks = New-Object System.Windows.Forms.GroupBox
$grpTasks.Text = 'Task Selection'
$grpTasks.Location = New-Object System.Drawing.Point(12, 276)
$grpTasks.Size = New-Object System.Drawing.Size(560, 430)

$taskList = New-Object System.Windows.Forms.CheckedListBox
$taskList.Location = New-Object System.Drawing.Point(16, 24)
$taskList.Size = New-Object System.Drawing.Size(525, 344)
$taskList.CheckOnClick = $true

foreach ($task in $taskDefinitions) {
    [void]$taskList.Items.Add(("{0:D2}: {1}" -f $task.Id, $task.Name), $true)
}

$btnSelectAll = New-Object System.Windows.Forms.Button
$btnSelectAll.Location = New-Object System.Drawing.Point(16, 380)
$btnSelectAll.Size = New-Object System.Drawing.Size(100, 30)
$btnSelectAll.Text = 'Select All'

$btnClearAll = New-Object System.Windows.Forms.Button
$btnClearAll.Location = New-Object System.Drawing.Point(126, 380)
$btnClearAll.Size = New-Object System.Drawing.Size(100, 30)
$btnClearAll.Text = 'Clear All'

$btnDiagnosePreset = New-Object System.Windows.Forms.Button
$btnDiagnosePreset.Location = New-Object System.Drawing.Point(236, 380)
$btnDiagnosePreset.Size = New-Object System.Drawing.Size(130, 30)
$btnDiagnosePreset.Text = 'Diagnose Preset'

$grpTasks.Controls.AddRange(@($taskList, $btnSelectAll, $btnClearAll, $btnDiagnosePreset))

$grpRun = New-Object System.Windows.Forms.GroupBox
$grpRun.Text = 'Run Status'
$grpRun.Location = New-Object System.Drawing.Point(586, 62)
$grpRun.Size = New-Object System.Drawing.Size(590, 205)

$lblState = New-Object System.Windows.Forms.Label
$lblState.Location = New-Object System.Drawing.Point(16, 28)
$lblState.Size = New-Object System.Drawing.Size(300, 20)
$lblState.Text = 'Status: Idle'

$lblCurrentTask = New-Object System.Windows.Forms.Label
$lblCurrentTask.Location = New-Object System.Drawing.Point(16, 54)
$lblCurrentTask.Size = New-Object System.Drawing.Size(550, 20)
$lblCurrentTask.Text = 'Current Task: -'

$lblElapsed = New-Object System.Windows.Forms.Label
$lblElapsed.Location = New-Object System.Drawing.Point(16, 80)
$lblElapsed.Size = New-Object System.Drawing.Size(220, 20)
$lblElapsed.Text = 'Elapsed: 00:00:00'

$lblExitCode = New-Object System.Windows.Forms.Label
$lblExitCode.Location = New-Object System.Drawing.Point(250, 80)
$lblExitCode.Size = New-Object System.Drawing.Size(160, 20)
$lblExitCode.Text = 'Exit Code: -'

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(16, 108)
$progress.Size = New-Object System.Drawing.Size(556, 24)
$progress.Minimum = 0
$progress.Maximum = 100
$progress.Value = 0

$lblProgress = New-Object System.Windows.Forms.Label
$lblProgress.Location = New-Object System.Drawing.Point(16, 138)
$lblProgress.Size = New-Object System.Drawing.Size(300, 20)
$lblProgress.Text = 'Progress: 0 / 0'

$lblLatestHtml = New-Object System.Windows.Forms.Label
$lblLatestHtml.Location = New-Object System.Drawing.Point(16, 164)
$lblLatestHtml.Size = New-Object System.Drawing.Size(556, 16)
$lblLatestHtml.Text = 'Latest HTML: -'

$lblLatestJson = New-Object System.Windows.Forms.Label
$lblLatestJson.Location = New-Object System.Drawing.Point(16, 182)
$lblLatestJson.Size = New-Object System.Drawing.Size(556, 16)
$lblLatestJson.Text = 'Latest JSON: -'

$grpRun.Controls.AddRange(@(
    $lblState, $lblCurrentTask, $lblElapsed, $lblExitCode,
    $progress, $lblProgress, $lblLatestHtml, $lblLatestJson
))

$grpCommand = New-Object System.Windows.Forms.GroupBox
$grpCommand.Text = 'Command Preview'
$grpCommand.Location = New-Object System.Drawing.Point(586, 276)
$grpCommand.Size = New-Object System.Drawing.Size(590, 120)

$txtCommand = New-Object System.Windows.Forms.TextBox
$txtCommand.Location = New-Object System.Drawing.Point(16, 24)
$txtCommand.Size = New-Object System.Drawing.Size(556, 80)
$txtCommand.Multiline = $true
$txtCommand.ScrollBars = 'Vertical'
$txtCommand.ReadOnly = $true

$grpCommand.Controls.Add($txtCommand)

$grpLog = New-Object System.Windows.Forms.GroupBox
$grpLog.Text = 'Execution Log'
$grpLog.Location = New-Object System.Drawing.Point(586, 404)
$grpLog.Size = New-Object System.Drawing.Size(590, 302)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(16, 24)
$txtLog.Size = New-Object System.Drawing.Size(556, 262)
$txtLog.Multiline = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.ReadOnly = $true
$txtLog.Font = New-Object System.Drawing.Font('MS Gothic', 9)

$grpLog.Controls.Add($txtLog)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Location = New-Object System.Drawing.Point(586, 12)
$btnRun.Size = New-Object System.Drawing.Size(100, 36)
$btnRun.Text = 'Run'

$btnStop = New-Object System.Windows.Forms.Button
$btnStop.Location = New-Object System.Drawing.Point(696, 12)
$btnStop.Size = New-Object System.Drawing.Size(100, 36)
$btnStop.Text = 'Stop'
$btnStop.Enabled = $false

$btnOpenReports = New-Object System.Windows.Forms.Button
$btnOpenReports.Location = New-Object System.Drawing.Point(806, 12)
$btnOpenReports.Size = New-Object System.Drawing.Size(110, 36)
$btnOpenReports.Text = 'Reports'

$btnOpenLogs = New-Object System.Windows.Forms.Button
$btnOpenLogs.Location = New-Object System.Drawing.Point(926, 12)
$btnOpenLogs.Size = New-Object System.Drawing.Size(110, 36)
$btnOpenLogs.Text = 'Logs'

$btnOpenDocs = New-Object System.Windows.Forms.Button
$btnOpenDocs.Location = New-Object System.Drawing.Point(1046, 12)
$btnOpenDocs.Size = New-Object System.Drawing.Size(130, 36)
$btnOpenDocs.Text = 'GUI Docs'

$form.Controls.AddRange(@(
    $lblEngine, $lblScript, $grpSettings, $grpTasks, $grpRun,
    $grpCommand, $grpLog, $btnRun, $btnStop, $btnOpenReports, $btnOpenLogs, $btnOpenDocs
))

$controlsToLock = @(
    $cmbMode, $cmbProfile, $cmbFailure, $chkWhatIf, $chkUseLocalChart,
    $chkExportPowerBI, $chkUseAnthropic, $chkExportDeleted, $cmbExportFormat,
    $txtConfig, $txtExportPath, $taskList, $btnSelectAll, $btnClearAll, $btnDiagnosePreset
)

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
        if ($fallbackHtml) { $sync.LatestHtml = $fallbackHtml.FullName }
    }
}

function Update-ArtifactLabels {
    $lblLatestHtml.Text = 'Latest HTML: ' + $(if ($sync.LatestHtml) { $sync.LatestHtml } else { '-' })
    $lblLatestJson.Text = 'Latest JSON: ' + $(if ($sync.LatestJson) { $sync.LatestJson } else { '-' })
}

function New-ExecutionArguments {
    param([int[]]$TaskIdsOverride = $null)

    $argumentList = New-Object 'System.Collections.Generic.List[string]'
    $selectedIds = Get-SelectedTaskIds
    if ($TaskIdsOverride) {
        $selectedIds = @($TaskIdsOverride | Sort-Object -Unique)
    }
    $taskToken = ConvertTo-TaskToken -TaskIds $selectedIds

    [void]$argumentList.Add('-NoProfile')
    [void]$argumentList.Add('-ExecutionPolicy')
    [void]$argumentList.Add('Bypass')
    [void]$argumentList.Add('-File')
    [void]$argumentList.Add($mainScriptPath)
    [void]$argumentList.Add('-NonInteractive')
    [void]$argumentList.Add('-NoRebootPrompt')
    [void]$argumentList.Add('-Mode')
    [void]$argumentList.Add([string]$cmbMode.SelectedItem)
    [void]$argumentList.Add('-ExecutionProfile')
    [void]$argumentList.Add([string]$cmbProfile.SelectedItem)
    [void]$argumentList.Add('-Tasks')
    [void]$argumentList.Add($taskToken)
    [void]$argumentList.Add('-FailureMode')
    [void]$argumentList.Add([string]$cmbFailure.SelectedItem)

    if ($chkWhatIf.Checked) {
        [void]$argumentList.Add('-WhatIf')
    }
    if ($chkUseLocalChart.Checked) {
        [void]$argumentList.Add('-UseLocalChartJs')
    }
    if ($chkExportPowerBI.Checked) {
        [void]$argumentList.Add('-ExportPowerBIJson')
    }
    if ($chkUseAnthropic.Checked) {
        [void]$argumentList.Add('-UseAnthropicAI')
    }
    if ($txtConfig.Text.Trim()) {
        [void]$argumentList.Add('-ConfigPath')
        [void]$argumentList.Add($txtConfig.Text.Trim())
    }
    if ($chkExportDeleted.Checked) {
        [void]$argumentList.Add('-ExportDeletedPaths')
        [void]$argumentList.Add([string]$cmbExportFormat.SelectedItem)
        if ($txtExportPath.Text.Trim()) {
            [void]$argumentList.Add('-ExportDeletedPathsPath')
            [void]$argumentList.Add($txtExportPath.Text.Trim())
        }
    }

    return @{
        Args = @($argumentList)
        SelectedIds = $selectedIds
    }
}

function Update-CommandPreview {
    $engine = Get-PreferredPowerShellExe
    $plan = New-ExecutionArguments
    $txtCommand.Text = $engine + ' ' + (Convert-ToArgumentString -Arguments $plan.Args)
}

function Append-LogLine {
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

    if ($Line -match 'start|開始|完了|失敗|スキップ|(?i:skip)') {
        $taskId = Get-TaskIdFromLine -Line $Line
        if ($null -ne $taskId) {
            $lblCurrentTask.Text = 'Current Task: ' + (Get-TaskDisplayName -TaskId $taskId)
        } else {
            $lblCurrentTask.Text = 'Current Task: ' + $Line
        }
    }

    if ($Line -match '完了|失敗|スキップ|(?i:skip)') {
        $taskId = Get-TaskIdFromLine -Line $Line
        if ($null -ne $taskId -and -not $sync.CompletedTaskIds.Contains($taskId)) {
            [void]$sync.CompletedTaskIds.Add($taskId)
            $sync.CompletedCount = [Math]::Min(($sync.CompletedCount + 1), $sync.SelectedIds.Count)
        }
    }

    $completed = $sync.CompletedCount
    $total = [Math]::Max($sync.SelectedIds.Count, 1)
    $progress.Value = [Math]::Min([int](($completed / $total) * 100), 100)
    $lblProgress.Text = "Progress: $completed / $($sync.SelectedIds.Count)"
}

function Stop-ActiveProcess {
    if ($sync.Process -and -not $sync.Process.HasExited) {
        try {
            $sync.Process.Kill()
            $sync.StopRequested = $true
            $lblState.Text = 'Status: Stopping'
        } catch {
            Append-LogLine -Text ("[gui] stop failed: " + $_.Exception.Message)
            $lblState.Text = 'Status: Stop failed'
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

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 250
$timer.Add_Tick({
    while ($true) {
        $item = $null
        if (-not $sync.Queue.TryDequeue([ref]$item)) { break }

        switch ($item.Kind) {
            'stdout' {
                Append-LogLine -Text $item.Text
                Update-ProgressFromLine -Line $item.Text
            }
            'stderr' {
                Append-LogLine -Text $item.Text -Prefix '[stderr] '
            }
            'exit' {
                $sync.Running = $false
                $sync.ExitCode = [int]$item.ExitCode
                Clear-EventSubscriptions
                Find-LatestArtifacts
                Update-ArtifactLabels
                $lblExitCode.Text = "Exit Code: $($sync.ExitCode)"
                if ($sync.StopRequested) {
                    $lblState.Text = 'Status: Stopped'
                } elseif ($sync.ExitCode -eq 0) {
                    $lblState.Text = 'Status: Completed'
                } else {
                    $lblState.Text = 'Status: Failed'
                }
                Set-ControlsEnabled -Enabled $true
            }
        }
    }

    if ($sync.Running -and $sync.StartTime) {
        $elapsed = (Get-Date) - $sync.StartTime
        $lblElapsed.Text = 'Elapsed: ' + $elapsed.ToString('hh\:mm\:ss')
    }
})
$timer.Start()

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
    $taskList.SetItemChecked(19, $true)
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
    } else {
        Update-CommandPreview
    }
}

$btnOpenReports.Add_Click({
    if (-not (Test-Path -LiteralPath $reportsDir)) {
        [System.Windows.Forms.MessageBox]::Show('reports folder was not found.', 'PC Optimizer GUI') | Out-Null
        return
    }
    Start-Process explorer.exe $reportsDir
})

$btnOpenLogs.Add_Click({
    if (-not (Test-Path -LiteralPath $logsDir)) {
        [System.Windows.Forms.MessageBox]::Show('logs folder was not found.', 'PC Optimizer GUI') | Out-Null
        return
    }
    Start-Process explorer.exe $logsDir
})

$btnOpenDocs.Add_Click({
    if (-not (Test-Path -LiteralPath $docsGuiDir)) {
        [System.Windows.Forms.MessageBox]::Show('docs\\GUI folder was not found.', 'PC Optimizer GUI') | Out-Null
        return
    }
    Start-Process explorer.exe $docsGuiDir
})

$btnStop.Add_Click({
    Stop-ActiveProcess
})

$btnRun.Add_Click({
    if ($sync.Running) { return }

    $plan = New-ExecutionArguments
    if ($plan.SelectedIds.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Select at least one task.', 'PC Optimizer GUI') | Out-Null
        return
    }
    if ($chkExportDeleted.Checked -and -not $chkWhatIf.Checked) {
        [System.Windows.Forms.MessageBox]::Show('ExportDeletedPaths requires WhatIf.', 'PC Optimizer GUI') | Out-Null
        return
    }
    if (($plan.SelectedIds -contains 18) -or ($plan.SelectedIds -contains 19)) {
        [System.Windows.Forms.MessageBox]::Show(
            'Tasks 18 and 19 are not supported in GUI v1. Deselect them before running.',
            'PC Optimizer GUI'
        ) | Out-Null
        return
    }

    $engine = Get-PreferredPowerShellExe
    $sync.Queue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    $sync.SelectedIds = $plan.SelectedIds
    $sync.CompletedCount = 0
    $sync.CompletedTaskIds = New-Object 'System.Collections.Generic.HashSet[int]'
    $sync.StopRequested = $false
    $sync.ExitCode = $null
    $sync.StartTime = Get-Date
    $sync.Running = $true

    $progress.Value = 0
    $lblProgress.Text = "Progress: 0 / $($sync.SelectedIds.Count)"
    $lblState.Text = 'Status: Running'
    $lblCurrentTask.Text = 'Current Task: -'
    $lblElapsed.Text = 'Elapsed: 00:00:00'
    $lblExitCode.Text = 'Exit Code: -'
    $txtLog.Clear()
    Append-LogLine -Text ("[gui] run started: " + (Get-Date -Format 'yyyy/MM/dd HH:mm:ss'))
    Append-LogLine -Text ("[gui] command: " + $engine + ' ' + (Convert-ToArgumentString -Arguments $plan.Args))

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $engine
    $psi.Arguments = Convert-ToArgumentString -Arguments $plan.Args
    $psi.WorkingDirectory = $repoRoot
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true

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
        $sync.Running = $false
        [System.Windows.Forms.MessageBox]::Show("Failed to start execution.`r`n$_", 'PC Optimizer GUI') | Out-Null
        Set-ControlsEnabled -Enabled $true
    }
})

$form.Add_FormClosing({
    Stop-ActiveProcess
    Clear-EventSubscriptions
})

Find-LatestArtifacts
Update-ArtifactLabels
Update-CommandPreview
[void]$form.ShowDialog()
