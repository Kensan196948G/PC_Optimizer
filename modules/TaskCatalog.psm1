Set-StrictMode -Version Latest

function Get-PCOptimizerTaskCatalog {
    [CmdletBinding()]
    param()

    @(
        [PSCustomObject]@{ Id = 1; Label = "Task 1: Temp files"; Category = "Cleanup"; IsReadOnly = $false }
        [PSCustomObject]@{ Id = 2; Label = "Task 2: Prefetch and update cache"; Category = "Cleanup"; IsReadOnly = $false }
        [PSCustomObject]@{ Id = 3; Label = "Task 3: Delivery optimization cache"; Category = "Cleanup"; IsReadOnly = $false }
        [PSCustomObject]@{ Id = 4; Label = "Task 4: Windows Update cache"; Category = "Cleanup"; IsReadOnly = $false }
        [PSCustomObject]@{ Id = 5; Label = "Task 5: Error reports and logs"; Category = "Cleanup"; IsReadOnly = $false }
        [PSCustomObject]@{ Id = 6; Label = "Task 6: OneDrive Teams Office cache"; Category = "Cleanup"; IsReadOnly = $false }
        [PSCustomObject]@{ Id = 7; Label = "Task 7: Browser cache"; Category = "Cleanup"; IsReadOnly = $false }
        [PSCustomObject]@{ Id = 8; Label = "Task 8: Thumbnail cache"; Category = "Cleanup"; IsReadOnly = $false }
        [PSCustomObject]@{ Id = 9; Label = "Task 9: Microsoft Store cache"; Category = "Cleanup"; IsReadOnly = $false }
        [PSCustomObject]@{ Id = 10; Label = "Task 10: Recycle Bin"; Category = "Cleanup"; IsReadOnly = $false }
        [PSCustomObject]@{ Id = 11; Label = "Task 11: DNS cache"; Category = "Cleanup"; IsReadOnly = $false }
        [PSCustomObject]@{ Id = 12; Label = "Task 12: Windows event logs"; Category = "Cleanup"; IsReadOnly = $false }
        [PSCustomObject]@{ Id = 13; Label = "Task 13: Disk optimize"; Category = "Performance"; IsReadOnly = $false }
        [PSCustomObject]@{ Id = 14; Label = "Task 14: SSD SMART check"; Category = "Diagnostics"; IsReadOnly = $false }
        [PSCustomObject]@{ Id = 15; Label = "Task 15: SFC scan"; Category = "Diagnostics"; IsReadOnly = $false }
        [PSCustomObject]@{ Id = 16; Label = "Task 16: DISM scan"; Category = "Diagnostics"; IsReadOnly = $false }
        [PSCustomObject]@{ Id = 17; Label = "Task 17: Power plan optimize"; Category = "Performance"; IsReadOnly = $false }
        [PSCustomObject]@{ Id = 18; Label = "Task 18: Microsoft 365 update"; Category = "Update"; IsReadOnly = $false }
        [PSCustomObject]@{ Id = 19; Label = "Task 19: Windows Update"; Category = "Update"; IsReadOnly = $false }
        [PSCustomObject]@{ Id = 20; Label = "Task 20: Startup and service report"; Category = "Report"; IsReadOnly = $true }
    )
}

function ConvertTo-PCOptimizerTaskSelection {
    [CmdletBinding()]
    param(
        [int[]]$TaskIds
    )

    $catalog = @(Get-PCOptimizerTaskCatalog)
    if (-not $TaskIds -or @($TaskIds).Count -eq 0) {
        return "all"
    }

    $normalized = @($TaskIds | Sort-Object -Unique)
    if (@($normalized).Count -eq @($catalog).Count) {
        $allIds = @($catalog | ForEach-Object { [int]$_.Id } | Sort-Object)
        if (@(Compare-Object -ReferenceObject $allIds -DifferenceObject $normalized).Count -eq 0) {
            return "all"
        }
    }

    return ($normalized -join ",")
}

function New-PCOptimizerArgumentList {
    [CmdletBinding()]
    param(
        [ValidateSet("repair", "diagnose")]
        [string]$Mode = "repair",
        [ValidateSet("classic", "agent-teams")]
        [string]$ExecutionProfile = "classic",
        [ValidateSet("continue", "fail-fast")]
        [string]$FailureMode = "continue",
        [int[]]$TaskIds,
        [string]$ConfigPath = "",
        [bool]$EnableAIDiagnosis = $true,
        [switch]$WhatIfMode,
        [switch]$NonInteractive,
        [switch]$NoRebootPrompt,
        [switch]$EmitUiEvents
    )

    $args = New-Object 'System.Collections.Generic.List[string]'
    [void]$args.Add("-Mode")
    [void]$args.Add($Mode)
    [void]$args.Add("-ExecutionProfile")
    [void]$args.Add($ExecutionProfile)
    [void]$args.Add("-FailureMode")
    [void]$args.Add($FailureMode)
    [void]$args.Add("-EnableAIDiagnosis:$EnableAIDiagnosis")

    $taskSelection = ConvertTo-PCOptimizerTaskSelection -TaskIds $TaskIds
    if ($taskSelection -ne "all") {
        [void]$args.Add("-Tasks")
        [void]$args.Add($taskSelection)
    }
    if ($ConfigPath) {
        [void]$args.Add("-ConfigPath")
        [void]$args.Add($ConfigPath)
    }
    if ($WhatIfMode) {
        [void]$args.Add("-WhatIf")
    }
    if ($NonInteractive) {
        [void]$args.Add("-NonInteractive")
    }
    if ($NoRebootPrompt) {
        [void]$args.Add("-NoRebootPrompt")
    }
    if ($EmitUiEvents) {
        [void]$args.Add("-EmitUiEvents")
    }

    return @($args)
}

Export-ModuleMember -Function Get-PCOptimizerTaskCatalog, ConvertTo-PCOptimizerTaskSelection, New-PCOptimizerArgumentList
