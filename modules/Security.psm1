Set-StrictMode -Version Latest

function Get-SecurityDiagnostics {
    [CmdletBinding()]
    param()

    [PSCustomObject]@{
        Defender = 'Unknown'
        Firewall = 'Unknown'
        BitLocker = 'Unknown'
        Uac       = 'Unknown'
    }
}

Export-ModuleMember -Function Get-SecurityDiagnostics
