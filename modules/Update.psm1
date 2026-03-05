Set-StrictMode -Version Latest

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.EXAMPLE
An example

.NOTES
General notes
#>
function Get-UpdateDiagnostic {
    [CmdletBinding()]
    param()

    $pendingCount = $null
    $wuStatus = 'Unknown'
    $history = @()

    try {
        $session = New-Object -ComObject Microsoft.Update.Session
        $searcher = $session.CreateUpdateSearcher()
        $result = $searcher.Search("IsInstalled=0 and Type='Software'")
        $pendingCount = [int]$result.Updates.Count
        $wuStatus = if ($pendingCount -eq 0) { 'Compliant' } else { 'PendingUpdates' }
        $historyCount = [math]::Min(30, $searcher.GetTotalHistoryCount())
        if ($historyCount -gt 0) {
            $history = @(
                $searcher.QueryHistory(0, $historyCount) |
                    Select-Object Date, Title, ResultCode, HResult
            )
        }
    } catch {
        $wuStatus = 'Unknown'
    }

    $updateErrors = @()
    try {
        $updateErrors = @(
            Get-WinEvent -FilterHashtable @{
                LogName      = 'System'
                Level        = 2
                ProviderName = 'Microsoft-Windows-WindowsUpdateClient'
                StartTime    = (Get-Date).AddDays(-14)
            } -ErrorAction Ignore |
            Select-Object -First 20 TimeCreated, Id, Message
        )
    } catch { $updateErrors = @() }

    $oneDriveRunning = @((Get-Process -Name OneDrive -ErrorAction SilentlyContinue)).Count -gt 0
    $outlookRunning = @((Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)).Count -gt 0
    $teamsClassicCache = Join-Path $env:APPDATA 'Microsoft\Teams\Cache'
    $teamsNewCache = Join-Path $env:LOCALAPPDATA 'Packages\MSTeams_8wekyb3d8bbwe\LocalCache'
    $teamsCachePath = if (Test-Path $teamsClassicCache) { $teamsClassicCache } elseif (Test-Path $teamsNewCache) { $teamsNewCache } else { $null }
    $teamsCacheSizeMb = $null
    if ($teamsCachePath) {
        try {
            $teamsCacheSizeMb = [math]::Round(((Get-ChildItem -Path $teamsCachePath -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum / 1MB), 2)
        } catch {
            $teamsCacheSizeMb = $null
        }
    }

    [PSCustomObject]@{
        WindowsUpdate = $wuStatus
        PendingCount  = $pendingCount
        UpdateHistory = $history
        UpdateErrors  = $updateErrors
        M365          = [PSCustomObject]@{
            OneDriveSyncStatus = if ($oneDriveRunning) { 'Running' } else { 'NotRunning' }
            TeamsCachePath     = $teamsCachePath
            TeamsCacheSizeMB   = $teamsCacheSizeMb
            OutlookStatus      = if ($outlookRunning) { 'Running' } else { 'NotRunning' }
        }
    }
}

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER WhatIfMode
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function Invoke-UpdateMaintenance {
    [CmdletBinding()]
    param([switch]$WhatIfMode)

    if ($WhatIfMode) {
        return [PSCustomObject]@{
            Component = 'Update'
            WhatIf    = $true
            Status    = 'PreviewOnly'
            Actions   = @('Windows Update Scan', 'Windows Update Download', 'Windows Update Install')
        }
    }

    $usoPath = Join-Path $env:SystemRoot 'System32\UsoClient.exe'
    if (-not (Test-Path $usoPath)) {
        return [PSCustomObject]@{
            Component = 'Update'
            WhatIf    = $false
            Status    = 'Skipped'
            Reason    = 'UsoClientNotFound'
        }
    }

    Start-Process -FilePath $usoPath -ArgumentList 'StartScan' -NoNewWindow -Wait
    Start-Process -FilePath $usoPath -ArgumentList 'StartDownload' -NoNewWindow -Wait
    Start-Process -FilePath $usoPath -ArgumentList 'StartInstall' -NoNewWindow -Wait

    [PSCustomObject]@{
        Component = 'Update'
        WhatIf    = $false
        Status    = 'Completed'
    }
}

Export-ModuleMember -Function Get-UpdateDiagnostic,Invoke-UpdateMaintenance
