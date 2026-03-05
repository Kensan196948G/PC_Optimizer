Set-StrictMode -Version Latest

function Get-NetworkDiagnostics {
    [CmdletBinding()]
    param()

    [PSCustomObject]@{
        IpAddress = @()
        Dns       = @()
        Gateway   = @()
        NicSpeed  = @()
        Status    = 'NotImplemented'
    }
}

function Get-NetworkDiagnostic {
    [CmdletBinding()]
    param()
    Get-NetworkDiagnostics
}

Export-ModuleMember -Function Get-NetworkDiagnostics,Get-NetworkDiagnostic
