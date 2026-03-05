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
function Get-SecurityDiagnostic {
    [CmdletBinding()]
    param()

    $defender = 'Unknown'
    if (Get-Command -Name Get-MpComputerStatus -ErrorAction SilentlyContinue) {
        $mp = Get-MpComputerStatus -ErrorAction SilentlyContinue
        if ($mp) {
            $defender = if ($mp.AntivirusEnabled -and $mp.RealTimeProtectionEnabled) { 'Enabled' } else { 'Warning' }
        }
    }

    $firewall = 'Unknown'
    if (Get-Command -Name Get-NetFirewallProfile -ErrorAction SilentlyContinue) {
        $profiles = Get-NetFirewallProfile -ErrorAction SilentlyContinue
        if ($profiles) {
            $allEnabled = @($profiles | Where-Object { -not $_.Enabled }).Count -eq 0
            $firewall = if ($allEnabled) { 'Enabled' } else { 'Warning' }
        }
    }

    $bitLocker = 'Unknown'
    if (Get-Command -Name Get-BitLockerVolume -ErrorAction SilentlyContinue) {
        try {
            $sysDrive = "$($env:SystemDrive)\"
            $bl = Get-BitLockerVolume -MountPoint $sysDrive -ErrorAction Stop
            if ($bl) {
                $bitLocker = if ($bl.ProtectionStatus -eq 'On' -or $bl.ProtectionStatus -eq 1) { 'Enabled' } else { 'Disabled' }
            }
        } catch {
            $bitLocker = 'Unknown'
        }
    }

    $uac = 'Unknown'
    $uacReg = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Name EnableLUA -ErrorAction SilentlyContinue
    if ($null -ne $uacReg) {
        $uac = if ($uacReg.EnableLUA -eq 1) { 'Enabled' } else { 'Disabled' }
    }

    [PSCustomObject]@{
        Defender = $defender
        Firewall = $firewall
        BitLocker = $bitLocker
        Uac       = $uac
    }
}

Export-ModuleMember -Function Get-SecurityDiagnostics
