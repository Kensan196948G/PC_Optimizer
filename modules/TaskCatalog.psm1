Set-StrictMode -Version Latest

function Resolve-UiText {
    param([string]$Text)
    return [System.Net.WebUtility]::HtmlDecode($Text)
}

function Get-PCOptimizerTaskCatalog {
    [CmdletBinding()]
    param()

    @(
        [PSCustomObject]@{ Id = 1; Label = (Resolve-UiText '&#x30BF;&#x30B9;&#x30AF;&#x0020;&#x0031;&#x003A;&#x0020;&#x4E00;&#x6642;&#x30D5;&#x30A1;&#x30A4;&#x30EB;&#x306E;&#x524A;&#x9664;'); Category = (Resolve-UiText '&#x30AF;&#x30EA;&#x30FC;&#x30F3;&#x30A2;&#x30C3;&#x30D7;'); IsReadOnly = $false }
        [PSCustomObject]@{ Id = 2; Label = (Resolve-UiText '&#x30BF;&#x30B9;&#x30AF;&#x0020;&#x0032;&#x003A;&#x0020;&#x0050;&#x0072;&#x0065;&#x0066;&#x0065;&#x0074;&#x0063;&#x0068;&#x30FB;&#x66F4;&#x65B0;&#x30AD;&#x30E3;&#x30C3;&#x30B7;&#x30E5;&#x306E;&#x524A;&#x9664;'); Category = (Resolve-UiText '&#x30AF;&#x30EA;&#x30FC;&#x30F3;&#x30A2;&#x30C3;&#x30D7;'); IsReadOnly = $false }
        [PSCustomObject]@{ Id = 3; Label = (Resolve-UiText '&#x30BF;&#x30B9;&#x30AF;&#x0020;&#x0033;&#x003A;&#x0020;&#x914D;&#x4FE1;&#x6700;&#x9069;&#x5316;&#x30AD;&#x30E3;&#x30C3;&#x30B7;&#x30E5;&#x306E;&#x524A;&#x9664;'); Category = (Resolve-UiText '&#x30AF;&#x30EA;&#x30FC;&#x30F3;&#x30A2;&#x30C3;&#x30D7;'); IsReadOnly = $false }
        [PSCustomObject]@{ Id = 4; Label = (Resolve-UiText '&#x30BF;&#x30B9;&#x30AF;&#x0020;&#x0034;&#x003A;&#x0020;&#x0057;&#x0069;&#x006E;&#x0064;&#x006F;&#x0077;&#x0073;&#x0020;&#x0055;&#x0070;&#x0064;&#x0061;&#x0074;&#x0065;&#x0020;&#x30AD;&#x30E3;&#x30C3;&#x30B7;&#x30E5;&#x306E;&#x524A;&#x9664;'); Category = (Resolve-UiText '&#x30AF;&#x30EA;&#x30FC;&#x30F3;&#x30A2;&#x30C3;&#x30D7;'); IsReadOnly = $false }
        [PSCustomObject]@{ Id = 5; Label = (Resolve-UiText '&#x30BF;&#x30B9;&#x30AF;&#x0020;&#x0035;&#x003A;&#x0020;&#x30A8;&#x30E9;&#x30FC;&#x30EC;&#x30DD;&#x30FC;&#x30C8;&#x30FB;&#x30ED;&#x30B0;&#x30FB;&#x4E0D;&#x8981;&#x30AD;&#x30E3;&#x30C3;&#x30B7;&#x30E5;&#x306E;&#x524A;&#x9664;'); Category = (Resolve-UiText '&#x30AF;&#x30EA;&#x30FC;&#x30F3;&#x30A2;&#x30C3;&#x30D7;'); IsReadOnly = $false }
        [PSCustomObject]@{ Id = 6; Label = (Resolve-UiText '&#x30BF;&#x30B9;&#x30AF;&#x0020;&#x0036;&#x003A;&#x0020;&#x004F;&#x006E;&#x0065;&#x0044;&#x0072;&#x0069;&#x0076;&#x0065;&#x0020;&#x002F;&#x0020;&#x0054;&#x0065;&#x0061;&#x006D;&#x0073;&#x0020;&#x002F;&#x0020;&#x004F;&#x0066;&#x0066;&#x0069;&#x0063;&#x0065;&#x0020;&#x30AD;&#x30E3;&#x30C3;&#x30B7;&#x30E5;&#x306E;&#x524A;&#x9664;'); Category = (Resolve-UiText '&#x30AF;&#x30EA;&#x30FC;&#x30F3;&#x30A2;&#x30C3;&#x30D7;'); IsReadOnly = $false }
        [PSCustomObject]@{ Id = 7; Label = (Resolve-UiText '&#x30BF;&#x30B9;&#x30AF;&#x0020;&#x0037;&#x003A;&#x0020;&#x30D6;&#x30E9;&#x30A6;&#x30B6;&#x30AD;&#x30E3;&#x30C3;&#x30B7;&#x30E5;&#x306E;&#x524A;&#x9664;'); Category = (Resolve-UiText '&#x30AF;&#x30EA;&#x30FC;&#x30F3;&#x30A2;&#x30C3;&#x30D7;'); IsReadOnly = $false }
        [PSCustomObject]@{ Id = 8; Label = (Resolve-UiText '&#x30BF;&#x30B9;&#x30AF;&#x0020;&#x0038;&#x003A;&#x0020;&#x30B5;&#x30E0;&#x30CD;&#x30A4;&#x30EB;&#x30AD;&#x30E3;&#x30C3;&#x30B7;&#x30E5;&#x306E;&#x524A;&#x9664;'); Category = (Resolve-UiText '&#x30AF;&#x30EA;&#x30FC;&#x30F3;&#x30A2;&#x30C3;&#x30D7;'); IsReadOnly = $false }
        [PSCustomObject]@{ Id = 9; Label = (Resolve-UiText '&#x30BF;&#x30B9;&#x30AF;&#x0020;&#x0039;&#x003A;&#x0020;&#x004D;&#x0069;&#x0063;&#x0072;&#x006F;&#x0073;&#x006F;&#x0066;&#x0074;&#x0020;&#x0053;&#x0074;&#x006F;&#x0072;&#x0065;&#x0020;&#x30AD;&#x30E3;&#x30C3;&#x30B7;&#x30E5;&#x306E;&#x30AF;&#x30EA;&#x30A2;'); Category = (Resolve-UiText '&#x30AF;&#x30EA;&#x30FC;&#x30F3;&#x30A2;&#x30C3;&#x30D7;'); IsReadOnly = $false }
        [PSCustomObject]@{ Id = 10; Label = (Resolve-UiText '&#x30BF;&#x30B9;&#x30AF;&#x0020;&#x0031;&#x0030;&#x003A;&#x0020;&#x3054;&#x307F;&#x7BB1;&#x3092;&#x7A7A;&#x306B;&#x3059;&#x308B;'); Category = (Resolve-UiText '&#x30AF;&#x30EA;&#x30FC;&#x30F3;&#x30A2;&#x30C3;&#x30D7;'); IsReadOnly = $false }
        [PSCustomObject]@{ Id = 11; Label = (Resolve-UiText '&#x30BF;&#x30B9;&#x30AF;&#x0020;&#x0031;&#x0031;&#x003A;&#x0020;&#x0044;&#x004E;&#x0053;&#x0020;&#x30AD;&#x30E3;&#x30C3;&#x30B7;&#x30E5;&#x306E;&#x30AF;&#x30EA;&#x30A2;'); Category = (Resolve-UiText '&#x30AF;&#x30EA;&#x30FC;&#x30F3;&#x30A2;&#x30C3;&#x30D7;'); IsReadOnly = $false }
        [PSCustomObject]@{ Id = 12; Label = (Resolve-UiText '&#x30BF;&#x30B9;&#x30AF;&#x0020;&#x0031;&#x0032;&#x003A;&#x0020;&#x0057;&#x0069;&#x006E;&#x0064;&#x006F;&#x0077;&#x0073;&#x0020;&#x30A4;&#x30D9;&#x30F3;&#x30C8;&#x30ED;&#x30B0;&#x306E;&#x30AF;&#x30EA;&#x30A2;'); Category = (Resolve-UiText '&#x30AF;&#x30EA;&#x30FC;&#x30F3;&#x30A2;&#x30C3;&#x30D7;'); IsReadOnly = $false }
        [PSCustomObject]@{ Id = 13; Label = (Resolve-UiText '&#x30BF;&#x30B9;&#x30AF;&#x0020;&#x0031;&#x0033;&#x003A;&#x0020;&#x30C7;&#x30A3;&#x30B9;&#x30AF;&#x306E;&#x6700;&#x9069;&#x5316;'); Category = (Resolve-UiText '&#x30D1;&#x30D5;&#x30A9;&#x30FC;&#x30DE;&#x30F3;&#x30B9;'); IsReadOnly = $false }
        [PSCustomObject]@{ Id = 14; Label = (Resolve-UiText '&#x30BF;&#x30B9;&#x30AF;&#x0020;&#x0031;&#x0034;&#x003A;&#x0020;&#x0053;&#x0053;&#x0044;&#x0020;&#x30D8;&#x30EB;&#x30B9;&#x30C1;&#x30A7;&#x30C3;&#x30AF;'); Category = (Resolve-UiText '&#x8A3A;&#x65AD;'); IsReadOnly = $false }
        [PSCustomObject]@{ Id = 15; Label = (Resolve-UiText '&#x30BF;&#x30B9;&#x30AF;&#x0020;&#x0031;&#x0035;&#x003A;&#x0020;&#x30B7;&#x30B9;&#x30C6;&#x30E0;&#x30D5;&#x30A1;&#x30A4;&#x30EB;&#x306E;&#x6574;&#x5408;&#x6027;&#x30C1;&#x30A7;&#x30C3;&#x30AF;&#x30FB;&#x4FEE;&#x5FA9;'); Category = (Resolve-UiText '&#x8A3A;&#x65AD;'); IsReadOnly = $false }
        [PSCustomObject]@{ Id = 16; Label = (Resolve-UiText '&#x30BF;&#x30B9;&#x30AF;&#x0020;&#x0031;&#x0036;&#x003A;&#x0020;&#x0057;&#x0069;&#x006E;&#x0064;&#x006F;&#x0077;&#x0073;&#x0020;&#x30B3;&#x30F3;&#x30DD;&#x30FC;&#x30CD;&#x30F3;&#x30C8;&#x30B9;&#x30C8;&#x30A2;&#x306E;&#x8A3A;&#x65AD;'); Category = (Resolve-UiText '&#x8A3A;&#x65AD;'); IsReadOnly = $false }
        [PSCustomObject]@{ Id = 17; Label = (Resolve-UiText '&#x30BF;&#x30B9;&#x30AF;&#x0020;&#x0031;&#x0037;&#x003A;&#x0020;&#x96FB;&#x6E90;&#x30D7;&#x30E9;&#x30F3;&#x306E;&#x6700;&#x9069;&#x5316;'); Category = (Resolve-UiText '&#x30D1;&#x30D5;&#x30A9;&#x30FC;&#x30DE;&#x30F3;&#x30B9;'); IsReadOnly = $false }
        [PSCustomObject]@{ Id = 18; Label = (Resolve-UiText '&#x30BF;&#x30B9;&#x30AF;&#x0020;&#x0031;&#x0038;&#x003A;&#x0020;&#x004D;&#x0069;&#x0063;&#x0072;&#x006F;&#x0073;&#x006F;&#x0066;&#x0074;&#x0020;&#x0033;&#x0036;&#x0035;&#x0020;&#x306E;&#x66F4;&#x65B0;&#x78BA;&#x8A8D;&#x30FB;&#x9069;&#x7528;'); Category = (Resolve-UiText '&#x66F4;&#x65B0;'); IsReadOnly = $false }
        [PSCustomObject]@{ Id = 19; Label = (Resolve-UiText '&#x30BF;&#x30B9;&#x30AF;&#x0020;&#x0031;&#x0039;&#x003A;&#x0020;&#x0057;&#x0069;&#x006E;&#x0064;&#x006F;&#x0077;&#x0073;&#x0020;&#x0055;&#x0070;&#x0064;&#x0061;&#x0074;&#x0065;&#x0020;&#x306E;&#x5B9F;&#x884C;'); Category = (Resolve-UiText '&#x66F4;&#x65B0;'); IsReadOnly = $false }
        [PSCustomObject]@{ Id = 20; Label = (Resolve-UiText '&#x30BF;&#x30B9;&#x30AF;&#x0020;&#x0032;&#x0030;&#x003A;&#x0020;&#x30B9;&#x30BF;&#x30FC;&#x30C8;&#x30A2;&#x30C3;&#x30D7;&#x30FB;&#x30B5;&#x30FC;&#x30D3;&#x30B9;&#x30EC;&#x30DD;&#x30FC;&#x30C8;'); Category = (Resolve-UiText '&#x30EC;&#x30DD;&#x30FC;&#x30C8;'); IsReadOnly = $true }
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
    [void]$args.Add("-EnableAIDiagnosis")
    [void]$args.Add($EnableAIDiagnosis.ToString())

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
