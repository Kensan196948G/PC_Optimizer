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
function Get-SystemDiagnostic {
    [CmdletBinding()]
    param()

    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
    $cpu = Get-CimInstance -ClassName Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
    $gpuList = Get-CimInstance -ClassName Win32_VideoController -ErrorAction SilentlyContinue
    $disks = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction SilentlyContinue

    $diskInfo = @(
        $disks | ForEach-Object {
            $sizeGb = [math]::Round(($_.Size / 1GB), 2)
            $freeGb = [math]::Round(($_.FreeSpace / 1GB), 2)
            $freePct = if ($_.Size -gt 0) {
                [math]::Round((($_.FreeSpace / $_.Size) * 100), 2)
            } else {
                0
            }
            [PSCustomObject]@{
                Drive       = $_.DeviceID
                FileSystem  = $_.FileSystem
                SizeGB      = $sizeGb
                FreeGB      = $freeGb
                FreePercent = $freePct
            }
        }
    )

    [PSCustomObject]@{
        ComputerName = $env:COMPUTERNAME
        UserName     = $env:USERNAME
        CollectedAt  = (Get-Date).ToString('s')
        Os           = [PSCustomObject]@{
            Caption       = if ($os) { $os.Caption } else { $null }
            Version       = if ($os) { $os.Version } else { $null }
            BuildNumber   = if ($os) { $os.BuildNumber } else { $null }
            LastBootUp    = if ($os) { $os.LastBootUpTime } else { $null }
        }
        Cpu          = [PSCustomObject]@{
            Name           = if ($cpu) { $cpu.Name } else { $null }
            LogicalCores   = if ($cpu) { $cpu.NumberOfLogicalProcessors } else { $null }
            PhysicalCores  = if ($cpu) { $cpu.NumberOfCores } else { $null }
        }
        Memory       = [PSCustomObject]@{
            TotalGB = if ($os) { [math]::Round(($os.TotalVisibleMemorySize * 1KB) / 1GB, 2) } else { $null }
            FreeGB  = if ($os) { [math]::Round(($os.FreePhysicalMemory * 1KB) / 1GB, 2) } else { $null }
        }
        Disk         = $diskInfo
        Gpu          = @($gpuList | ForEach-Object { $_.Name })
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
function Get-AssetInventory {
    [CmdletBinding()]
    param()

    $diag = Get-SystemDiagnostics
    $nics = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled=TRUE" -ErrorAction SilentlyContinue

    [PSCustomObject]@{
        PcName = $env:COMPUTERNAME
        User   = $env:USERNAME
        OS     = if ($diag.Os) { $diag.Os.Caption } else { $null }
        IP     = @($nics | ForEach-Object { @($_.IPAddress) } | Where-Object { $_ -match '\.' })
        MAC    = @($nics | ForEach-Object { $_.MACAddress } | Where-Object { $_ })
        CPU    = if ($diag.Cpu) { $diag.Cpu.Name } else { $null }
        RAMGB  = if ($diag.Memory) { $diag.Memory.TotalGB } else { $null }
        Disk   = $diag.Disk
    }
}

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER Hours
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function Get-EventLogSummary {
    [CmdletBinding()]
    param([int]$Hours = 24)

    $since = (Get-Date).AddHours(-1 * [math]::Abs($Hours))

    $appErrors = Get-WinEvent -FilterHashtable @{
        LogName   = 'Application'
        Level     = 2
        StartTime = $since
    } -ErrorAction Ignore

    $sysErrors = Get-WinEvent -FilterHashtable @{
        LogName   = 'System'
        Level     = 2
        StartTime = $since
    } -ErrorAction Ignore

    $bsod = Get-WinEvent -FilterHashtable @{
        LogName      = 'System'
        Id           = 1001
        ProviderName = 'Microsoft-Windows-WER-SystemErrorReporting'
        StartTime    = $since
    } -ErrorAction Ignore

    [PSCustomObject]@{
        Hours             = $Hours
        ApplicationErrors = @($appErrors).Count
        SystemErrors      = @($sysErrors).Count
        BsodCount         = @($bsod).Count
        RecentErrors      = @(
            @($appErrors + $sysErrors) |
                Sort-Object TimeCreated -Descending |
                Select-Object -First 20 TimeCreated, LogName, Id, ProviderName, LevelDisplayName, Message
        )
    }
}

Export-ModuleMember -Function Get-SystemDiagnostics,Get-AssetInventory,Get-EventLogSummary
