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

Export-ModuleMember -Function Get-NetworkDiagnostics
