Set-StrictMode -Version Latest
$script:_enc = if ($PSVersionTable.PSVersion.Major -ge 7) { 'utf8NoBOM' } else { 'UTF8' }

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER Path
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function Get-OptimizerConfig {
    [CmdletBinding()]
    param(
        [string]$Path = (Join-Path (Split-Path $PSScriptRoot -Parent) 'config\config.json')
    )

    if (-not (Test-Path -Path $Path)) {
        throw "Config file not found: $Path"
    }

    return Get-Content -Path $Path -Raw -Encoding utf8 | ConvertFrom-Json
}

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER Message
Parameter description

.PARAMETER Level
Parameter description

.PARAMETER Path
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function Write-StructuredLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','DEBUG')]
        [string]$Level = 'INFO',
        [string]$Path
    )

    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    if ($Path) {
        Add-Content -Path $Path -Value $line -Encoding $script:_enc
    }
    $line
}

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER Name
Parameter description

.PARAMETER Action
Parameter description

.PARAMETER ErrorLogPath
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function Invoke-GuardedStep {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name,
        [Parameter(Mandatory)]
        [scriptblock]$Action,
        [string]$ErrorLogPath
    )

    try {
        & $Action
        return [PSCustomObject]@{ Name = $Name; Status = 'OK'; Error = '' }
    } catch {
        $msg = "{0}: {1}" -f $Name, $_
        if ($ErrorLogPath) {
            Add-Content -Path $ErrorLogPath -Value $msg -Encoding $script:_enc
        }
        return [PSCustomObject]@{ Name = $Name; Status = 'NG'; Error = $msg }
    }
}

Export-ModuleMember -Function Get-OptimizerConfig,Write-StructuredLog,Invoke-GuardedStep
