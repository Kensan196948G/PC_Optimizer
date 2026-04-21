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

$ConfigPathText.Text = $configPath
$EngineText.Text = Get-HostExecutable

$taskItems = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
foreach ($task in @(Get-PCOptimizerTaskCatalog)) {
    $taskItems.Add([PSCustomObject]@{
        Id = [int]$task.Id
        Label = "$($task.Label)"
        Category = "$($task.Category)"
        IsReadOnly = [bool]$task.IsReadOnly
        ReadOnlyLabel = if ($task.IsReadOnly) { 'Yes' } else { 'No' }
        IsSelected = $true
    })
}
$TasksListView.ItemsSource = $taskItems

$global:PCOptimizerGuiSync = [hashtable]::Synchronized(@{
    Window = $window
    OutputTextBox = $OutputTextBox
    StatusText = $StatusText
    ProgressText = $ProgressText
    ProgressBar = $RunProgressBar
    RunButton = $RunButton
    OpenReportButton = $OpenReportButton
    LatestReportText = $LatestReportText
    LatestReportPath = ''
    CompletedTasks = 0
    SelectedTaskCount = 0
    EventPrefix = '##PCOPT_UI##'
    LastExitCode = $null
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
    $proc = $global:PCOptimizerGuiSync.Process
    if ($proc -and -not $proc.HasExited) {
        try { $proc.Kill() } catch {}
    }
})

$RunButton.Add_Click({
    if ($global:PCOptimizerGuiSync.Process -and -not $global:PCOptimizerGuiSync.Process.HasExited) {
        return
    }

    $selectedIds = @(
        $taskItems |
            Where-Object { $_.IsSelected } |
            ForEach-Object { [int]$_.Id }
    )
    if (@($selectedIds).Count -eq 0) {
        $StatusText.Text = "Select at least one task."
        return
    }

    $global:PCOptimizerGuiSync.CompletedTasks = 0
    $global:PCOptimizerGuiSync.SelectedTaskCount = @($selectedIds).Count
    $global:PCOptimizerGuiSync.LatestReportPath = ''
    $global:PCOptimizerGuiSync.LastExitCode = $null
    $OpenReportButton.IsEnabled = $false
    $LatestReportText.Text = 'Waiting for report'
    $RunProgressBar.Minimum = 0
    $RunProgressBar.Maximum = [Math]::Max(1, @($selectedIds).Count)
    $RunProgressBar.Value = 0
    $ProgressText.Text = "0 / $(@($selectedIds).Count)"
    $StatusText.Text = 'Preparing run'
    $OutputTextBox.Clear()
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
    $engineArgString = ConvertTo-ProcessArgumentString -Arguments $engineArgs
    $fullArgs = '-NoProfile -ExecutionPolicy Bypass -File "{0}" {1}' -f $enginePath, $engineArgString

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $hostExe
    $psi.Arguments = $fullArgs
    $psi.WorkingDirectory = $repoRoot
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    $proc.EnableRaisingEvents = $true
    $global:PCOptimizerGuiSync.Process = $proc

    $proc.add_OutputDataReceived({
        param($sender, $eventArgs)
        $line = $eventArgs.Data
        if ([string]::IsNullOrWhiteSpace($line)) { return }

        $sync = $global:PCOptimizerGuiSync
        $sync.Window.Dispatcher.Invoke([action]{
            if ($line.StartsWith($sync.EventPrefix, [System.StringComparison]::Ordinal)) {
                $payload = $line.Substring($sync.EventPrefix.Length)
                try {
                    $evt = $payload | ConvertFrom-Json
                    switch ("$($evt.type)") {
                        'run_start' {
                            $sync.StatusText.Text = "Started: $($evt.mode) / $($evt.executionProfile)"
                        }
                        'task_start' {
                            $sync.StatusText.Text = "Running: #$($evt.taskId) $($evt.taskName)"
                        }
                        'task_finish' {
                            $sync.CompletedTasks = [Math]::Min($sync.SelectedTaskCount, ($sync.CompletedTasks + 1))
                            $sync.ProgressBar.Value = $sync.CompletedTasks
                            $sync.ProgressText.Text = "{0} / {1}" -f $sync.CompletedTasks, $sync.SelectedTaskCount
                            $sync.StatusText.Text = "Done: #$($evt.taskId) $($evt.status)"
                        }
                        'report_generated' {
                            if ($evt.htmlReportPath) {
                                $sync.LatestReportPath = "$($evt.htmlReportPath)"
                                $sync.LatestReportText.Text = "$($evt.htmlReportPath)"
                                $sync.OpenReportButton.IsEnabled = $true
                            }
                        }
                        'run_complete' {
                            $sync.LastExitCode = [int]$evt.exitCode
                            $sync.StatusText.Text = "Completed: status=$($evt.status) exit=$($evt.exitCode)"
                            $sync.RunButton.IsEnabled = $true
                        }
                    }
                } catch {
                    $sync.OutputTextBox.AppendText("[ui-event parse error] " + $payload + [Environment]::NewLine)
                    $sync.OutputTextBox.ScrollToEnd()
                }
            } else {
                $sync.OutputTextBox.AppendText($line + [Environment]::NewLine)
                $sync.OutputTextBox.ScrollToEnd()
            }
        })
    })

    $proc.add_ErrorDataReceived({
        param($sender, $eventArgs)
        $line = $eventArgs.Data
        if ([string]::IsNullOrWhiteSpace($line)) { return }

        $sync = $global:PCOptimizerGuiSync
        $sync.Window.Dispatcher.Invoke([action]{
            $sync.OutputTextBox.AppendText("[stderr] " + $line + [Environment]::NewLine)
            $sync.OutputTextBox.ScrollToEnd()
        })
    })

    $proc.add_Exited({
        $sync = $global:PCOptimizerGuiSync
        $exitCode = $sync.Process.ExitCode
        $sync.Window.Dispatcher.Invoke([action]{
            $sync.RunButton.IsEnabled = $true
            if ($null -eq $sync.LastExitCode) {
                $sync.StatusText.Text = "Process exited: exit=$exitCode"
            }
        })
    })

    [void]$proc.Start()
    $proc.BeginOutputReadLine()
    $proc.BeginErrorReadLine()
})

$window.ShowDialog() | Out-Null
