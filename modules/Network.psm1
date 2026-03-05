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
function Get-NetworkDiagnostic {
    [CmdletBinding()]
    param()

    $nics = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled=TRUE" -ErrorAction SilentlyContinue
    $adapters = Get-CimInstance -ClassName Win32_NetworkAdapter -Filter "NetEnabled=TRUE" -ErrorAction SilentlyContinue

    $ipAddresses = @($nics | ForEach-Object { @($_.IPAddress) } | Where-Object { $_ } | Select-Object -Unique)
    $dnsServers = @($nics | ForEach-Object { @($_.DNSServerSearchOrder) } | Where-Object { $_ } | Select-Object -Unique)
    $gateways = @($nics | ForEach-Object { @($_.DefaultIPGateway) } | Where-Object { $_ } | Select-Object -Unique)
    $nicSpeed = @(
        $adapters | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                Status = if ($_.NetEnabled) { 'Up' } else { 'Down' }
                LinkSpeed = if ($_.Speed) { "{0} Mbps" -f [math]::Round(([double]$_.Speed / 1MB), 2) } else { $null }
                MacAddress = $_.MACAddress
            }
        }
    )

    [PSCustomObject]@{
        IpAddress = $ipAddresses
        Dns       = $dnsServers
        Gateway   = $gateways
        NicSpeed  = $nicSpeed
        Status    = 'OK'
    }
}

Export-ModuleMember -Function Get-NetworkDiagnostic
