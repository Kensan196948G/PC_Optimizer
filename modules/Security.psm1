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

function Get-SecurityDiagnostic {
    [CmdletBinding()]
    param()
    Get-SecurityDiagnostics
}

Export-ModuleMember -Function Get-SecurityDiagnostics,Get-SecurityDiagnostic
