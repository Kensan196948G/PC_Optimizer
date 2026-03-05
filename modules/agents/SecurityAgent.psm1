Set-StrictMode -Version Latest
$sdkPath = Join-Path $PSScriptRoot "AgentSDK.psm1"
Import-Module $sdkPath -Force -ErrorAction Stop

function Invoke-AgentPlugin {
    [CmdletBinding()]
    param([string]$Role,[string]$AgentId,[pscustomobject]$Node,[hashtable]$Context)
    if ("$AgentId" -ne "SecurityAgent") { return $null }
    if (-not (Test-AgentPluginContext -Context $Context)) {
        return New-AgentPluginResult -Status Failed -Risk Unknown -Message "security-invalid-context"
    }

    if (("$Role").ToLowerInvariant() -eq "collector") {
        $d = Get-AgentSnapshot -Context $Context -Key "securityDiagnostics"
        $risk = "Low"
        if ($d -and (($d.PSObject.Properties["Defender"] -and $d.Defender -and $d.Defender -ne "Enabled") -or ($d.PSObject.Properties["Firewall"] -and $d.Firewall -and $d.Firewall -ne "Enabled"))) { $risk = "High" }
        return New-AgentPluginResult -Status Success -Risk $risk -Payload $d -Message "security-plugin-collector"
    }
    if (("$Role").ToLowerInvariant() -eq "analyzer" -and "$($Node.id)" -ne "analyzer.aggregate") {
        $col = Get-AgentCollectorResult -Context $Context -AgentId "SecurityAgent"
        if (-not $col) { return New-AgentPluginResult -Status Failed -Risk Unknown -Message "security-plugin-missing-collector" }
        return New-AgentPluginResult -Status Success -Risk "$($col.risk)" -Payload ([PSCustomObject]@{ source = "plugin"; sourceRisk = "$($col.risk)" }) -Message "security-plugin-analyzer"
    }
    return $null
}

Export-ModuleMember -Function Invoke-AgentPlugin
