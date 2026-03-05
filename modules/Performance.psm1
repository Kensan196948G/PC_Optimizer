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
function Get-PerformanceSnapshot {
    [CmdletBinding()]
    param()

    $processes = @(Get-Process -ErrorAction SilentlyContinue)
    $cpuRows = @(
        $processes | ForEach-Object {
            $cpuVal = $null
            try { $cpuVal = [double]$_.CPU } catch { $cpuVal = 0 }
            [PSCustomObject]@{
                Name = $_.ProcessName
                Id = $_.Id
                CPU = [math]::Round($cpuVal, 2)
            }
        }
    )
    $cpuTop = @(
        $cpuRows |
            Sort-Object CPU -Descending |
            Select-Object -First 10 Name, Id, CPU
    )
    $memoryTop = @(
        $processes |
            Sort-Object WorkingSet64 -Descending |
            Select-Object -First 10 Name, Id, @{Name='MemoryMB';Expression={[math]::Round($_.WorkingSet64 / 1MB, 2)}}
    )
    $diskIoTop = @(
        $processes |
            Select-Object Name, Id, @{
                Name='DiskIOMB'
                Expression={ [math]::Round(($_.IOReadBytes + $_.IOWriteBytes) / 1MB, 2) }
            } |
            Sort-Object DiskIOMB -Descending |
            Select-Object -First 10
    )

    [PSCustomObject]@{
        CpuTopProcess    = $cpuTop
        MemoryTopProcess = $memoryTop
        DiskIoTopProcess = $diskIoTop
        Status           = 'OK'
    }
}

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
function Get-StartupAnalysis {
    [CmdletBinding()]
    param()

    $startupApps = Get-CimInstance -ClassName Win32_StartupCommand -ErrorAction SilentlyContinue
    $count = @($startupApps).Count
    $rating = if ($count -le 5) {
        'Good'
    } elseif ($count -le 10) {
        'Normal'
    } else {
        'High'
    }

    [PSCustomObject]@{
        StartupCount = $count
        Rating       = $rating
        Items        = @($startupApps | Select-Object Name, Command, User)
    }
}

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER Cpu
Parameter description

.PARAMETER Memory
Parameter description

.PARAMETER Disk
Parameter description

.PARAMETER Startup
Parameter description

.PARAMETER Security
Parameter description

.PARAMETER Network
Parameter description

.PARAMETER WindowsUpdate
Parameter description

.PARAMETER SystemHealth
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function Get-HealthScore {
    [CmdletBinding()]
    param(
        [int]$Cpu = 0,
        [int]$Memory = 0,
        [int]$Disk = 0,
        [int]$Startup = 0,
        [int]$Security = 0,
        [int]$Network = 0,
        [int]$WindowsUpdate = 0,
        [int]$SystemHealth = 0
    )

    $weights = [ordered]@{
        Cpu           = 15
        Memory        = 15
        Disk          = 20
        Startup       = 10
        Security      = 15
        Network       = 10
        WindowsUpdate = 10
        SystemHealth  = 5
    }

    $scores = [ordered]@{
        Cpu           = [math]::Max([math]::Min($Cpu, 100), 0)
        Memory        = [math]::Max([math]::Min($Memory, 100), 0)
        Disk          = [math]::Max([math]::Min($Disk, 100), 0)
        Startup       = [math]::Max([math]::Min($Startup, 100), 0)
        Security      = [math]::Max([math]::Min($Security, 100), 0)
        Network       = [math]::Max([math]::Min($Network, 100), 0)
        WindowsUpdate = [math]::Max([math]::Min($WindowsUpdate, 100), 0)
        SystemHealth  = [math]::Max([math]::Min($SystemHealth, 100), 0)
    }

    $weighted = [ordered]@{}
    $total = 0.0
    foreach ($k in $weights.Keys) {
        $componentScore = [double]$scores[$k]
        $componentWeight = [double]$weights[$k]
        $points = ($componentScore / 100.0) * $componentWeight
        $weighted[$k] = [math]::Round($points, 2)
        $total += $points
    }
    $totalRounded = [math]::Round($total, 0)

    [PSCustomObject]@{
        Score          = [int][Math]::Max([Math]::Min($totalRounded, 100), 0)
        Status         = if ($totalRounded -ge 90) { 'Excellent' } elseif ($totalRounded -ge 75) { 'Good' } elseif ($totalRounded -ge 60) { 'Warning' } else { 'Critical' }
        Weight         = $weights
        ScoreInput     = $scores
        WeightedPoints = [PSCustomObject]$weighted
    }
}

Export-ModuleMember -Function Get-PerformanceSnapshot,Get-StartupAnalysis,Get-HealthScore
