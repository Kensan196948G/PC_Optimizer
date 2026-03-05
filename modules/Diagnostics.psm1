Set-StrictMode -Version Latest

function Get-SystemDiagnostics {
    [CmdletBinding()]
    param()

    [PSCustomObject]@{
        ComputerName = $env:COMPUTERNAME
        UserName     = $env:USERNAME
        PsVersion    = $PSVersionTable.PSVersion.ToString()
        CollectedAt  = (Get-Date).ToString('s')
    }
}

function Get-SystemDiagnostic {
    [CmdletBinding()]
    param()
    Get-SystemDiagnostics
}

function Get-AssetInventory {
    [CmdletBinding()]
    param()

    [PSCustomObject]@{
        PcName = $env:COMPUTERNAME
        User   = $env:USERNAME
    }
}

function Get-EventLogSummary {
    [CmdletBinding()]
    param([int]$Hours = 24)

    [PSCustomObject]@{
        Hours  = $Hours
        Status = 'NotImplemented'
    }
}

Export-ModuleMember -Function Get-SystemDiagnostics,Get-SystemDiagnostic,Get-AssetInventory,Get-EventLogSummary
