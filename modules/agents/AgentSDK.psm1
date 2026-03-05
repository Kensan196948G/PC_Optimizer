Set-StrictMode -Version Latest

function Test-AgentPluginContext {
    [CmdletBinding()]
    param(
        [hashtable]$Context
    )
    if (-not $Context) { return $false }
    return $Context.ContainsKey("RunId") -and $Context.ContainsKey("ModuleSnapshot")
}

function Get-AgentSnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Context,
        [Parameter(Mandatory)][string]$Key
    )
    if (-not $Context.ContainsKey("ModuleSnapshot")) { return $null }
    $snapshot = $Context.ModuleSnapshot
    if (-not $snapshot) { return $null }
    if ($snapshot.PSObject.Properties[$Key]) { return $snapshot.$Key }
    return $null
}

function Get-AgentCollectorResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Context,
        [Parameter(Mandatory)][string]$AgentId
    )
    if (-not $Context.ContainsKey("CollectorByAgent")) { return $null }
    $map = $Context.CollectorByAgent
    if (-not $map) { return $null }
    if ($map.ContainsKey($AgentId)) { return $map[$AgentId] }
    return $null
}

function New-AgentPluginResult {
    [CmdletBinding()]
    param(
        [ValidateSet("Success", "Failed", "Skipped")][string]$Status = "Success",
        [ValidateSet("Low", "Medium", "High", "Unknown")][string]$Risk = "Low",
        [object]$Payload = $null,
        [string]$Message = ""
    )
    return [PSCustomObject]@{
        status = $Status
        risk = $Risk
        payload = $Payload
        message = $Message
    }
}

function Test-AgentPluginResultSchema {
    [CmdletBinding()]
    param(
        [AllowNull()][object]$Result
    )
    if (-not $Result) { return $false }
    if (-not $Result.PSObject.Properties["status"]) { return $false }
    if (-not $Result.PSObject.Properties["risk"]) { return $false }
    if (-not $Result.PSObject.Properties["message"]) { return $false }
    $status = "$($Result.status)"
    $risk = "$($Result.risk)"
    if ($status -notin @("Success", "Failed", "Skipped")) { return $false }
    if ($risk -notin @("Low", "Medium", "High", "Unknown")) { return $false }
    return $true
}

function ConvertTo-ValidatedAgentPluginResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Result,
        [string]$FallbackMessage = "plugin-result-schema-invalid"
    )
    if (Test-AgentPluginResultSchema -Result $Result) {
        return [PSCustomObject]@{
            status = "$($Result.status)"
            risk = "$($Result.risk)"
            payload = if ($Result.PSObject.Properties["payload"]) { $Result.payload } else { $null }
            message = "$($Result.message)"
        }
    }
    return [PSCustomObject]@{
        status = "Failed"
        risk = "Unknown"
        payload = $null
        message = $FallbackMessage
    }
}

function Get-AgentPluginInfo {
    [CmdletBinding()]
    param()
    return [PSCustomObject]@{
        sdkVersion = "1.0"
        schemaVersion = "1.1"
        requiredFunctions = @("Invoke-AgentPlugin")
    }
}

Export-ModuleMember -Function `
    Test-AgentPluginContext, `
    Get-AgentSnapshot, `
    Get-AgentCollectorResult, `
    New-AgentPluginResult, `
    Test-AgentPluginResultSchema, `
    ConvertTo-ValidatedAgentPluginResult, `
    Get-AgentPluginInfo
