Set-StrictMode -Version Latest
$sdkPath = Join-Path $PSScriptRoot "AgentSDK.psm1"
Import-Module $sdkPath -Force -ErrorAction Stop

function Invoke-AgentPlugin {
    [CmdletBinding()]
    param([string]$Role,[string]$AgentId,[pscustomobject]$Node,[hashtable]$Context)
    if ("$AgentId" -ne "UpdateAgent") { return $null }
    if (-not (Test-AgentPluginContext -Context $Context)) {
        return New-AgentPluginResult -Status Failed -Risk Unknown -Message "update-invalid-context"
    }

    if (("$Role").ToLowerInvariant() -eq "collector") {
        $d = Get-AgentSnapshot -Context $Context -Key "updateDiagnostics"
        $risk = if ($d -and $d.WindowsUpdate -and $d.WindowsUpdate -ne "Compliant") { "Medium" } else { "Low" }
        return New-AgentPluginResult -Status Success -Risk $risk -Payload $d -Message "update-plugin-collector"
    }
    if (("$Role").ToLowerInvariant() -eq "analyzer" -and "$($Node.id)" -ne "analyzer.aggregate") {
        $col = Get-AgentCollectorResult -Context $Context -AgentId "UpdateAgent"
        if (-not $col) { return New-AgentPluginResult -Status Failed -Risk Unknown -Message "update-plugin-missing-collector" }
        return New-AgentPluginResult -Status Success -Risk "$($col.risk)" -Payload ([PSCustomObject]@{ source = "plugin"; sourceRisk = "$($col.risk)" }) -Message "update-plugin-analyzer"
    }
    return $null
}

Export-ModuleMember -Function Invoke-AgentPlugin
