Set-StrictMode -Version Latest
$sdkPath = Join-Path $PSScriptRoot "AgentSDK.psm1"
Import-Module $sdkPath -Force -ErrorAction Stop

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER Role
Parameter description

.PARAMETER AgentId
Parameter description

.PARAMETER Node
Parameter description

.PARAMETER Context
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function Invoke-AgentPlugin {
    [CmdletBinding()]
    param([string]$Role,[string]$AgentId,[pscustomobject]$Node,[hashtable]$Context)
    if ("$AgentId" -ne "NetworkAgent") { return $null }
    if (-not (Test-AgentPluginContext -Context $Context)) {
        return New-AgentPluginResult -Status Failed -Risk Unknown -Message "network-invalid-context"
    }

    if (("$Role").ToLowerInvariant() -eq "collector") {
        $d = Get-AgentSnapshot -Context $Context -Key "networkDiagnostics"
        $risk = if ($d -and ($null -eq $d.PrimaryIPv4 -or $d.PrimaryIPv4 -eq "")) { "Medium" } else { "Low" }
        return New-AgentPluginResult -Status Success -Risk $risk -Payload $d -Message "network-plugin-collector"
    }
    if (("$Role").ToLowerInvariant() -eq "analyzer" -and "$($Node.id)" -ne "analyzer.aggregate") {
        $col = Get-AgentCollectorResult -Context $Context -AgentId "NetworkAgent"
        if (-not $col) { return New-AgentPluginResult -Status Failed -Risk Unknown -Message "network-plugin-missing-collector" }
        return New-AgentPluginResult -Status Success -Risk "$($col.risk)" -Payload ([PSCustomObject]@{ source = "plugin"; sourceRisk = "$($col.risk)" }) -Message "network-plugin-analyzer"
    }
    return $null
}

Export-ModuleMember -Function Invoke-AgentPlugin
