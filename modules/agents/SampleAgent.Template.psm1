Set-StrictMode -Version Latest
$sdkPath = Join-Path $PSScriptRoot "AgentSDK.psm1"
Import-Module $sdkPath -Force -ErrorAction Stop

function Invoke-AgentPlugin {
    [CmdletBinding()]
    param(
        [string]$Role,
        [string]$AgentId,
        [pscustomobject]$Node,
        [hashtable]$Context
    )

    if (-not (Test-AgentPluginContext -Context $Context)) {
        return New-AgentPluginResult -Status Failed -Risk Unknown -Message "invalid-context"
    }
    if ("$AgentId" -ne "SampleAgent") { return $null }

    if ("$Role".ToLowerInvariant() -eq "collector") {
        $payload = [PSCustomObject]@{
            sample = $true
            runId = $Context.RunId
            collectedAt = (Get-Date).ToString("s")
        }
        return New-AgentPluginResult -Status Success -Risk Low -Payload $payload -Message "sample-collector"
    }

    if ("$Role".ToLowerInvariant() -eq "analyzer") {
        $col = Get-AgentCollectorResult -Context $Context -AgentId "SampleAgent"
        if (-not $col) {
            return New-AgentPluginResult -Status Failed -Risk Unknown -Message "sample-missing-collector"
        }
        return New-AgentPluginResult -Status Success -Risk "$($col.risk)" -Payload ([PSCustomObject]@{ source = "sample"; sourceRisk = $col.risk }) -Message "sample-analyzer"
    }

    return $null
}

Export-ModuleMember -Function Invoke-AgentPlugin
