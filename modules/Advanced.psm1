Set-StrictMode -Version Latest
$script:_enc = if ($PSVersionTable.PSVersion.Major -ge 7) { 'utf8NoBOM' } else { 'UTF8' }

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER Value
Parameter description

.PARAMETER Default
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function Normalize-AiEvaluation {
    [CmdletBinding()]
    param(
        [string]$Value,
        [string]$Default = "注意"
    )

    $v = if ($null -eq $Value) { "" } else { "$Value" }
    $norm = $v.Trim().ToLowerInvariant()
    switch ($norm) {
        "excellent" { return "優秀" }
        "good" { return "良好" }
        "warning" { return "注意" }
        "critical" { return "重大" }
        "優秀" { return "優秀" }
        "良好" { return "良好" }
        "注意" { return "注意" }
        "重大" { return "重大" }
        default { return $Default }
    }
}

function Get-PathSummary {
    [CmdletBinding()]
    param([string]$Path)

    $resolved = $Path
    try { $resolved = [IO.Path]::GetFullPath($Path) } catch {}
    $exists = Test-Path -LiteralPath $resolved
    if (-not $exists) {
        return [PSCustomObject]@{ Path = $resolved; Exists = $false; FileCount = 0; TotalSizeMB = 0 }
    }
    $files = Get-ChildItem -LiteralPath $resolved -File -Recurse -ErrorAction SilentlyContinue
    $sum = ($files | Measure-Object -Property Length -Sum).Sum
    [PSCustomObject]@{
        Path = $resolved
        Exists = $true
        FileCount = @($files).Count
        TotalSizeMB = [math]::Round(($sum / 1MB), 2)
    }
}

function New-RepairAllowList {
    [CmdletBinding()]
    param()

    @(
        "$env:SystemRoot\\Temp*",
        "$env:TEMP*",
        "$env:USERPROFILE\\AppData\\Local\\Temp*",
        "$env:SystemRoot\\Prefetch*",
        "$env:SystemRoot\\SoftwareDistribution*",
        "$env:SystemRoot\\System32\\DeliveryOptimization\\Cache*",
        "$env:ProgramData\\Microsoft\\Windows\\WER*",
        "$env:SystemRoot\\Logs\\CBS*",
        "$env:LOCALAPPDATA\\Microsoft\\OneDrive\\logs*",
        "$env:APPDATA\\Microsoft\\Teams\\Cache*",
        "$env:LOCALAPPDATA\\Microsoft\\Office\\16.0\\OfficeFileCache*",
        "$env:LOCALAPPDATA\\Google\\Chrome\\User Data\\Default\\Cache*",
        "$env:LOCALAPPDATA\\Microsoft\\Edge\\User Data\\Default\\Cache*",
        "$env:APPDATA\\Mozilla\\Firefox\\Profiles\\*",
        "$env:LOCALAPPDATA\\BraveSoftware\\Brave-Browser\\User Data\\Default\\Cache*",
        "$env:APPDATA\\Opera Software\\Opera Stable\\Cache*",
        "$env:LOCALAPPDATA\\Vivaldi\\User Data\\Default\\Cache*",
        "$env:LOCALAPPDATA\\Microsoft\\Windows\\Explorer*",
        "$env:LOCALAPPDATA\\Packages\\Microsoft.WindowsStore_8wekyb3d8bbwe\\LocalCache*",
        "$env:LOCALAPPDATA\\Packages\\Microsoft.WindowsStore_8wekyb3d8bbwe\\LocalState\\Cache*",
        "$env:LOCALAPPDATA\\Temp*"
    )
}

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER Mode
Parameter description

.PARAMETER RootPath
Parameter description

.PARAMETER LogsDir
Parameter description

.PARAMETER WhatIfMode
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function Start-RepairGuardrail {
    [CmdletBinding()]
    param(
        [ValidateSet('repair','diagnose')]
        [string]$Mode = 'repair',
        [Parameter(Mandatory)]
        [string]$RootPath,
        [Parameter(Mandatory)]
        [string]$LogsDir,
        [switch]$WhatIfMode
    )

    $stamp = Get-Date -Format "yyyyMMddHHmmss"
    $snapshotTargets = @(
        (Join-Path $RootPath "logs"),
        (Join-Path $RootPath "reports"),
        "$env:TEMP",
        "$env:SystemRoot\\Temp",
        "$env:SystemRoot\\SoftwareDistribution\\Download"
    )
    $before = @($snapshotTargets | ForEach-Object { Get-PathSummary -Path $_ })
    $beforePath = Join-Path $LogsDir "Guardrail_Before_${stamp}.json"
    $before | ConvertTo-Json -Depth 5 | Set-Content -Path $beforePath -Encoding $script:_enc

    $restorePointCreated = $false
    $restorePointReason = "Skipped"
    if ($Mode -eq 'repair' -and -not $WhatIfMode) {
        if (Get-Command Checkpoint-Computer -ErrorAction SilentlyContinue) {
            try {
                Checkpoint-Computer -Description ("PC_Optimizer_{0}" -f $stamp) -RestorePointType "MODIFY_SETTINGS" -ErrorAction Stop | Out-Null
                $restorePointCreated = $true
                $restorePointReason = "Created"
            } catch {
                $restorePointReason = "Failed: $($_.Exception.Message)"
            }
        } else {
            $restorePointReason = "Checkpoint-Computer unavailable"
        }
    }

    [PSCustomObject]@{
        Mode = $Mode
        StartedAt = (Get-Date).ToString("s")
        SnapshotTargets = $snapshotTargets
        SnapshotBeforePath = $beforePath
        SnapshotAfterPath = $null
        AllowList = New-RepairAllowList
        RestorePointCreated = $restorePointCreated
        RestorePointStatus = $restorePointReason
        BlockedPaths = @()
    }
}

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER State
Parameter description

.PARAMETER Path
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function Test-RepairAllowListPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$State,
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not $State -or -not $State.AllowList) { return $false }
    $resolved = $Path
    try { $resolved = [IO.Path]::GetFullPath($Path) } catch {}
    foreach ($p in @($State.AllowList)) {
        if ($resolved -like $p) { return $true }
    }
    return $false
}

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER State
Parameter description

.PARAMETER LogsDir
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function Complete-RepairGuardrail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$State,
        [Parameter(Mandatory)]
        [string]$LogsDir
    )

    $stamp = Get-Date -Format "yyyyMMddHHmmss"
    $after = @($State.SnapshotTargets | ForEach-Object { Get-PathSummary -Path $_ })
    $afterPath = Join-Path $LogsDir "Guardrail_After_${stamp}.json"
    $after | ConvertTo-Json -Depth 5 | Set-Content -Path $afterPath -Encoding $script:_enc

    $manifest = [PSCustomObject]@{
        mode = $State.Mode
        startedAt = $State.StartedAt
        finishedAt = (Get-Date).ToString("s")
        restorePointCreated = $State.RestorePointCreated
        restorePointStatus = $State.RestorePointStatus
        snapshotBeforePath = $State.SnapshotBeforePath
        snapshotAfterPath = $afterPath
        blockedPaths = @($State.BlockedPaths)
    }
    $manifestPath = Join-Path $LogsDir "Guardrail_Manifest_${stamp}.json"
    $manifest | ConvertTo-Json -Depth 5 | Set-Content -Path $manifestPath -Encoding $script:_enc
    return $manifestPath
}

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER UpdateErrors
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function Get-UpdateErrorClassification {
    [CmdletBinding()]
    param([object[]]$UpdateErrors)

    $map = @{
        "0x8024402C" = @{ Category = "NetworkOrProxy"; Recommendation = "Check proxy and DNS settings and allow Windows Update endpoints in firewall." }
        "0x80070005" = @{ Category = "Permission"; Recommendation = "Check administrator privileges and WSUS/GPO policy settings." }
        "0x8024A105" = @{ Category = "WUServiceState"; Recommendation = "Restart wuauserv and BITS services, then retry." }
        "0x80240022" = @{ Category = "PolicyDisabled"; Recommendation = "Review policy settings that disable Windows Update." }
    }

    $classified = @()
    foreach ($e in @($UpdateErrors)) {
        $msg = ""
        try { $msg = "$($e.Message)" } catch {}
        if ([string]::IsNullOrWhiteSpace($msg)) { $msg = "$e" }
        $code = $null
        $m = [regex]::Match($msg, '0x[0-9A-Fa-f]{8}')
        if ($m.Success) { $code = $m.Value }
        if (-not $code) {
            try {
                if ($null -ne $e.HResult) { $code = ('0x{0:X8}' -f ([int]$e.HResult)) }
            } catch {}
        }
        $cat = "Unknown"
        $rec = "Check event details and triage network, policy, and service state in order."
        if ($code -and $map.ContainsKey($code)) {
            $cat = $map[$code].Category
            $rec = $map[$code].Recommendation
        } elseif ($msg -match 'timeout|timed out|connect') {
            $cat = "Connectivity"
            $rec = "Check enterprise network connectivity, name resolution, and proxy reachability."
        }
        $classified += [PSCustomObject]@{ Code = if ($code) { $code } else { "N/A" }; Category = $cat; Message = $msg; Recommendation = $rec }
    }
    return $classified
}

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
function Test-M365Connectivity {
    [CmdletBinding()]
    param()

    $targets = @("login.microsoftonline.com","graph.microsoft.com","outlook.office365.com","teams.microsoft.com","onedrive.live.com")
    $results = @()
    foreach ($t in $targets) {
        $ok = $false
        $latency = $null
        try {
            $tnc = Test-NetConnection -ComputerName $t -Port 443 -WarningAction SilentlyContinue
            $ok = [bool]$tnc.TcpTestSucceeded
            $latency = if ($tnc.PingSucceeded) { [int]$tnc.PingReplyDetails.RoundtripTime } else { $null }
        } catch {}
        $results += [PSCustomObject]@{ Target = $t; Port = 443; Reachable = $ok; LatencyMs = $latency }
    }
    return $results
}

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER Hours
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function Get-EventLogAnomaly {
    [CmdletBinding()]
    param([int]$Hours = 24)

    $now = Get-Date
    $since = $now.AddHours(-1 * [math]::Abs($Hours))
    $curr = @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; Level = 2; StartTime = $since } -ErrorAction Ignore)
    $histSince = $now.AddDays(-7)
    $hist = @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; Level = 2; StartTime = $histSince; EndTime = $since } -ErrorAction Ignore)
    $currentCount = @($curr).Count
    $histCount = @($hist).Count
    $baselinePerDay = if ($histCount -gt 0) { [math]::Round($histCount / 6.0, 2) } else { 0 }
    $currentPerDayEq = [math]::Round($currentCount * (24.0 / [math]::Abs($Hours)), 2)
    $ratio = if ($baselinePerDay -gt 0) { [math]::Round($currentPerDayEq / $baselinePerDay, 2) } else { 0 }
    $status = if ($ratio -ge 2 -and $currentCount -ge 10) { "Anomaly" } else { "Normal" }
    [PSCustomObject]@{ Hours = $Hours; CurrentErrorCount = $currentCount; BaselinePerDay = $baselinePerDay; CurrentEquivalentPerDay = $currentPerDayEq; SpikeRatio = $ratio; Status = $status }
}

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER OutputDir
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function Update-BootShutdownTrend {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$OutputDir)

    if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }
    $trendPath = Join-Path $OutputDir "BootShutdownTrend.json"
    $history = @()
    if (Test-Path $trendPath) {
        try { $history = @(Get-Content -Path $trendPath -Raw -Encoding utf8 | ConvertFrom-Json) } catch { $history = @() }
    }

    $dayStart = (Get-Date).Date
    $dayEnd = $dayStart.AddDays(1)
    $sys = @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; StartTime = $dayStart; EndTime = $dayEnd; Id = 6005,6006,6008,41 } -ErrorAction Ignore)
    $startupEvents = @(Get-WinEvent -FilterHashtable @{ LogName = 'Microsoft-Windows-Diagnostics-Performance/Operational'; StartTime = $dayStart; EndTime = $dayEnd; Id = 100 } -ErrorAction Ignore)
    $avgBootMs = $null
    if (@($startupEvents).Count -gt 0) {
    $vals = @(
        $startupEvents | ForEach-Object {
            try {
                if ($_.Properties -and $_.Properties.Count -gt 6) { $_.Properties[6].Value }
            } catch {}
        } | Where-Object { $_ -is [int] -or $_ -is [long] }
    )
        if (@($vals).Count -gt 0) { $avgBootMs = [int][math]::Round((($vals | Measure-Object -Average).Average), 0) }
    }

    $today = [PSCustomObject]@{
        Date = (Get-Date -Format "yyyy-MM-dd")
        BootEvents = @($sys | Where-Object { $_.Id -eq 6005 }).Count
        ShutdownEvents = @($sys | Where-Object { $_.Id -eq 6006 }).Count
        UnexpectedShutdown = @($sys | Where-Object { $_.Id -eq 6008 -or $_.Id -eq 41 }).Count
        AvgBootDurationMs = $avgBootMs
    }
    $newHistory = @($history | Where-Object { $_.Date -ne $today.Date }) + @($today)
    $newHistory = @($newHistory | Sort-Object Date | Select-Object -Last 90)
    $newHistory | ConvertTo-Json -Depth 5 | Set-Content -Path $trendPath -Encoding $script:_enc
    return $today
}

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER TaskName
Parameter description

.PARAMETER ScriptPath
Parameter description

.PARAMETER ScriptArguments
Parameter description

.PARAMETER DayOfMonth
Parameter description

.PARAMETER At
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function Set-MonthlyMaintenanceTask {
    [CmdletBinding()]
    param(
        [string]$TaskName = "PC_Optimizer_Monthly",
        [Parameter(Mandatory)][string]$ScriptPath,
        [string]$ScriptArguments = "-NonInteractive -Mode repair",
        [int]$DayOfMonth = 1,
        [string]$At = "03:00"
    )

    if (Get-Command Register-ScheduledTask -ErrorAction SilentlyContinue) {
        $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument ("-NoProfile -ExecutionPolicy Bypass -File `"{0}`" {1}" -f $ScriptPath, $ScriptArguments)
        $trigger = New-ScheduledTaskTrigger -Monthly -DaysOfMonth $DayOfMonth -At $At
        $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
        $task = New-ScheduledTask -Action $action -Trigger $trigger -Principal $principal
        Register-ScheduledTask -TaskName $TaskName -InputObject $task -Force | Out-Null
        return "RegisteredWithScheduledTasksModule"
    }
    $cmd = "schtasks /Create /TN `"$TaskName`" /TR `"powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`" $ScriptArguments`" /SC MONTHLY /MO 1 /D $DayOfMonth /ST $At /RL HIGHEST /F"
    cmd.exe /c $cmd | Out-Null
    return "RegisteredWithSchtasks"
}

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER TaskName
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function Remove-MonthlyMaintenanceTask {
    [CmdletBinding()]
    param([string]$TaskName = "PC_Optimizer_Monthly")

    if (Get-Command Unregister-ScheduledTask -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        return "RemovedWithScheduledTasksModule"
    }
    cmd.exe /c ("schtasks /Delete /TN `"{0}`" /F" -f $TaskName) | Out-Null
    return "RemovedWithSchtasks"
}

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER ComputerName
Parameter description

.PARAMETER OutputDir
Parameter description

.PARAMETER RetryCount
Parameter description

.PARAMETER ThrottleLimit
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function Invoke-WinRMRemoteDiagnosticsBatch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$ComputerName,
        [Parameter(Mandatory)][string]$OutputDir,
        [int]$RetryCount = 1,
        [int]$ThrottleLimit = 8
    )

    if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }
    $pending = @($ComputerName | Where-Object { $_ } | Select-Object -Unique)
    $success = New-Object 'System.Collections.Generic.List[string]'
    $failed = New-Object 'System.Collections.Generic.List[string]'

    for ($attempt = 1; $attempt -le ($RetryCount + 1); $attempt++) {
        if (@($pending).Count -eq 0) { break }
        $job = Invoke-Command -ComputerName $pending -ThrottleLimit $ThrottleLimit -AsJob -ErrorAction SilentlyContinue -ScriptBlock {
            $os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
            $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
            [PSCustomObject]@{
                ComputerName = $env:COMPUTERNAME
                UserName = $env:USERNAME
                Os = if ($os) { $os.Caption } else { $null }
                Version = if ($os) { $os.Version } else { $null }
                TotalMemoryGB = if ($cs) { [math]::Round(($cs.TotalPhysicalMemory / 1GB), 2) } else { $null }
                CollectedAt = (Get-Date).ToString("s")
            }
        }
        Wait-Job $job | Out-Null
        $result = @()
        try { $result = @(Receive-Job -Job $job -ErrorAction SilentlyContinue) } catch { $result = @() }
        Remove-Job $job -Force -ErrorAction SilentlyContinue | Out-Null

        $hit = @($result | Select-Object -ExpandProperty PSComputerName -Unique)
        foreach ($row in $result) {
            $path = Join-Path $OutputDir ("RemoteDiag_{0}.json" -f $row.PSComputerName)
            $row | ConvertTo-Json -Depth 5 | Set-Content -Path $path -Encoding $script:_enc
            if (-not $success.Contains($row.PSComputerName)) { [void]$success.Add($row.PSComputerName) }
        }
        $pending = @($pending | Where-Object { $hit -notcontains $_ })
    }
    foreach ($p in $pending) { [void]$failed.Add($p) }

    $summary = [PSCustomObject]@{ Requested = @($ComputerName); Succeeded = @($success); Failed = @($failed); GeneratedAt = (Get-Date).ToString("s") }
    $summaryPath = Join-Path $OutputDir "RemoteDiag_Summary.json"
    $summary | ConvertTo-Json -Depth 5 | Set-Content -Path $summaryPath -Encoding $script:_enc
    return $summary
}

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER InputDir
Parameter description

.PARAMETER OutputDir
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function Invoke-AssetCentralAggregation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$InputDir,
        [Parameter(Mandatory)][string]$OutputDir
    )

    if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }
    $files = @(Get-ChildItem -Path $InputDir -File -Filter *.json -Recurse -ErrorAction SilentlyContinue)
    $rows = @()
    foreach ($f in $files) {
        try {
            $obj = Get-Content -Path $f.FullName -Raw -Encoding utf8 | ConvertFrom-Json
            $asset = $null
            if ($obj.PSObject.Properties["moduleSnapshot"] -and $obj.moduleSnapshot.assetInventory) { $asset = $obj.moduleSnapshot.assetInventory }
            elseif ($obj.PSObject.Properties["assetInventory"]) { $asset = $obj.assetInventory }
            elseif ($obj.PSObject.Properties["PcName"]) { $asset = $obj }
            if ($asset) {
                $rows += [PSCustomObject]@{ PcName = $asset.PcName; User = $asset.User; OS = $asset.OS; CPU = $asset.CPU; RAMGB = $asset.RAMGB; SourceFile = $f.FullName }
            }
        } catch {}
    }

    $stamp = Get-Date -Format "yyyyMMddHHmmss"
    $csvPath = Join-Path $OutputDir "AssetInventory_Aggregated_${stamp}.csv"
    $jsonPath = Join-Path $OutputDir "AssetInventory_Aggregated_${stamp}.json"
    $rows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    $rows | ConvertTo-Json -Depth 6 | Set-Content -Path $jsonPath -Encoding $script:_enc

    $baselinePath = Join-Path $OutputDir "AssetInventory_Baseline.csv"
    $diffPath = Join-Path $OutputDir "AssetInventory_Diff_${stamp}.csv"
    if (Test-Path $baselinePath) {
        $base = @(Import-Csv -Path $baselinePath)
        $curr = @($rows)
        $cmp = Compare-Object -ReferenceObject $base -DifferenceObject $curr -Property PcName,User,OS,CPU,RAMGB -PassThru
        $cmp | Select-Object PcName,User,OS,CPU,RAMGB,SideIndicator | Export-Csv -Path $diffPath -NoTypeInformation -Encoding UTF8
    } else {
        @() | Export-Csv -Path $diffPath -NoTypeInformation -Encoding UTF8
    }
    $rows | Export-Csv -Path $baselinePath -NoTypeInformation -Encoding UTF8

    [PSCustomObject]@{ Count = @($rows).Count; AggregatedCsv = $csvPath; AggregatedJson = $jsonPath; DiffCsv = $diffPath }
}

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER HealthScore
Parameter description

.PARAMETER Snapshot
Parameter description

.PARAMETER UpdateClassifiedErrors
Parameter description

.PARAMETER M365Connectivity
Parameter description

.PARAMETER EventAnomaly
Parameter description

.PARAMETER BootTrend
Parameter description

.PARAMETER AnthropicApiKey
Parameter description

.PARAMETER AnthropicModel
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function Invoke-AIDiagnosis {
    [CmdletBinding()]
    param(
        [pscustomobject]$HealthScore,
        [pscustomobject]$Snapshot,
        [object[]]$UpdateClassifiedErrors,
        [object[]]$M365Connectivity,
        [pscustomobject]$EventAnomaly,
        [pscustomobject]$BootTrend,
        [string]$AnthropicApiKey = "",
        [string]$AnthropicModel = "claude-sonnet-4-6"
    )

    $findings = @()
    $recs = @()
    $score = if ($HealthScore) { [int]$HealthScore.Score } else { 0 }
    $status = if ($HealthScore) { "$($HealthScore.Status)" } else { "Unknown" }
    $promptVersion = "2026-03-05.1"
    $statusJa = switch ($status.ToLowerInvariant()) {
        "excellent" { "優秀" }
        "good" { "良好" }
        "warning" { "注意" }
        "critical" { "重大" }
        default { if ($score -ge 90) { "優秀" } elseif ($score -ge 75) { "良好" } elseif ($score -ge 60) { "注意" } else { "重大" } }
    }
    $dataTimestamp = (Get-Date).ToString("s")
    $inputMetrics = [PSCustomObject]@{
        score = $score
        scoreStatus = $status
        cpuScore = if ($HealthScore -and $HealthScore.ScoreInput) { $HealthScore.ScoreInput.Cpu } else { $null }
        memoryScore = if ($HealthScore -and $HealthScore.ScoreInput) { $HealthScore.ScoreInput.Memory } else { $null }
        diskScore = if ($HealthScore -and $HealthScore.ScoreInput) { $HealthScore.ScoreInput.Disk } else { $null }
        startupScore = if ($HealthScore -and $HealthScore.ScoreInput) { $HealthScore.ScoreInput.Startup } else { $null }
        securityScore = if ($HealthScore -and $HealthScore.ScoreInput) { $HealthScore.ScoreInput.Security } else { $null }
        networkScore = if ($HealthScore -and $HealthScore.ScoreInput) { $HealthScore.ScoreInput.Network } else { $null }
        windowsUpdateScore = if ($HealthScore -and $HealthScore.ScoreInput) { $HealthScore.ScoreInput.WindowsUpdate } else { $null }
        systemHealthScore = if ($HealthScore -and $HealthScore.ScoreInput) { $HealthScore.ScoreInput.SystemHealth } else { $null }
        updateErrorCount = @($UpdateClassifiedErrors).Count
        m365UnreachableCount = if ($M365Connectivity) { @($M365Connectivity | Where-Object { -not $_.Reachable }).Count } else { 0 }
        eventAnomalyStatus = if ($EventAnomaly) { $EventAnomaly.Status } else { $null }
        bootAvgDurationMs = if ($BootTrend) { $BootTrend.AvgBootDurationMs } else { $null }
    }

    if ($score -lt 60) {
        $findings += "総合スコアが低く、業務影響リスクが高い状態です。"
        $recs += "Windows Update適用状況、ディスク空き容量、スタートアップ項目の順で是正してください。"
    } elseif ($score -lt 75) {
        $findings += "中程度の劣化が見られます。運用可能ですが改善余地があります。"
        $recs += "月次メンテナンスを定期実行し、常駐アプリの整理を推奨します。"
    } else {
        $findings += "全体的に健全です。予防保全中心の運用が可能です。"
        $recs += "現行運用を維持し、更新失敗イベントの監視を継続してください。"
    }

    if ($EventAnomaly -and $EventAnomaly.Status -eq "Anomaly") {
        $findings += ("直近{0}時間のSystem Error件数がベースライン比 {1} 倍です。" -f $EventAnomaly.Hours, $EventAnomaly.SpikeRatio)
        $recs += "上位イベントIDを最近の更新・ドライバ変更と照合して原因を特定してください。"
    }

    if ($M365Connectivity) {
        $ng = @($M365Connectivity | Where-Object { -not $_.Reachable }).Count
        if ($ng -gt 0) {
            $findings += "Microsoft 365接続テストで到達不可エンドポイントが検出されました。"
            $recs += "Microsoft 365向けのプロキシ/ファイアウォール許可設定を確認してください。"
        }
    }

    if ($BootTrend -and $BootTrend.AvgBootDurationMs -and $BootTrend.AvgBootDurationMs -gt 120000) {
        $findings += ("起動時間が長い傾向です（平均 {0} ms）。" -f $BootTrend.AvgBootDurationMs)
        $recs += "スタートアップアプリと自動起動サービスを見直してください。"
    }

    if ($UpdateClassifiedErrors -and @($UpdateClassifiedErrors).Count -gt 0) {
        $topCat = @($UpdateClassifiedErrors | Group-Object Category | Sort-Object Count -Descending | Select-Object -First 1)
        if ($topCat) {
            $findings += ("Windows Updateエラーは {0} 系が最多です。" -f $topCat[0].Name)
            $topRec = @($UpdateClassifiedErrors | Where-Object { $_.Category -eq $topCat[0].Name } | Select-Object -First 1).Recommendation
            if ($topRec) { $recs += "$topRec" }
        }
    }

    $localResult = [PSCustomObject]@{
        Source = "LocalRuleEngine"
        Headline = ("PC評価: {0} ({1}/100)" -f $statusJa, $score)
        Evaluation = if ($score -ge 90) { "優秀" } elseif ($score -ge 75) { "良好" } elseif ($score -ge 60) { "注意" } else { "重大" }
        EvaluationCode = if ($score -ge 90) { "Excellent" } elseif ($score -ge 75) { "Good" } elseif ($score -ge 60) { "Warning" } else { "Critical" }
        Summary = ("総合スコア {0}/100（{1}）です。" -f $score, $statusJa)
        Confidence = 0.55
        DataTimestamp = $dataTimestamp
        PromptVersion = $promptVersion
        InputMetrics = $inputMetrics
        Findings = @($findings)
        Recommendations = @($recs | Select-Object -Unique)
    }

    function Get-JsonFromText {
        param([string]$Text)
        if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
        $candidate = $Text.Trim()
        if ($candidate -match '```json\s*(?<j>\{[\s\S]*\})\s*```') {
            $candidate = $matches['j']
        } elseif ($candidate -match '```(?<j>\{[\s\S]*\})\s*```') {
            $candidate = $matches['j']
        } elseif ($candidate -match '(?<j>\{[\s\S]*\})') {
            $candidate = $matches['j']
        }
        try { return ($candidate | ConvertFrom-Json -ErrorAction Stop) } catch { return $null }
    }

    $fallbackReason = ""
    if (-not [string]::IsNullOrWhiteSpace($AnthropicApiKey)) {
        try {
            $prompt = @"
あなたは企業IT運用の診断アシスタントです。必ずJSONのみを返してください。
キーは次を厳守してください:
{
  "summary": "要約（日本語）",
  "evaluation": "優秀|良好|注意|重大 のいずれか",
  "confidence": 0.0-1.0 の数値,
  "evidence": ["根拠1","根拠2","根拠3"],
  "recommended_actions": ["推奨1","推奨2","推奨3"]
}
説明文やMarkdownは不要です。JSON以外は返さないでください。
promptVersion: $promptVersion
入力データ:
$(($localResult | ConvertTo-Json -Depth 8))
"@
            $body = @{ model = $AnthropicModel; max_tokens = 700; messages = @(@{ role = "user"; content = $prompt }) } | ConvertTo-Json -Depth 6
            $headers = @{ "x-api-key" = $AnthropicApiKey; "anthropic-version" = "2023-06-01"; "content-type" = "application/json" }
            # PS5.1 では Invoke-RestMethod が UTF-8 レスポンスをシステムデフォルト
            # エンコーディング (ANSI/CP932) で解釈して文字化けが発生することがある。
            # Invoke-WebRequest + 明示的 UTF-8 デコードで回避する。
            $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($body)
            $text = ""
            if ($PSVersionTable.PSVersion.Major -ge 7) {
                $resp = Invoke-RestMethod -Uri "https://api.anthropic.com/v1/messages" `
                    -Method Post -Headers $headers -Body $bodyBytes -TimeoutSec 45
                if ($resp.content -and @($resp.content).Count -gt 0) { $text = "$($resp.content[0].text)" }
            } else {
                $wresp = Invoke-WebRequest -Uri "https://api.anthropic.com/v1/messages" `
                    -Method Post -Headers $headers -Body $bodyBytes -TimeoutSec 45
                $respObj = [System.Text.Encoding]::UTF8.GetString($wresp.RawContentStream.ToArray()) |
                    ConvertFrom-Json -ErrorAction Stop
                if ($respObj.content -and @($respObj.content).Count -gt 0) { $text = "$($respObj.content[0].text)" }
            }
            if (-not [string]::IsNullOrWhiteSpace($text)) {
                $parsed = Get-JsonFromText -Text $text
                if ($parsed) {
                    $aiEvalRaw = if ($parsed.evaluation) { "$($parsed.evaluation)" } else { $localResult.Evaluation }
                    $aiEval = Normalize-AiEvaluation -Value $aiEvalRaw -Default $localResult.Evaluation
                    $aiSummary = if ($parsed.summary) { "$($parsed.summary)" } else { $localResult.Summary }
                    $aiConfidence = $localResult.Confidence
                    try {
                        if ($null -ne $parsed.confidence) {
                            $aiConfidence = [math]::Round([double]$parsed.confidence, 2)
                            if ($aiConfidence -lt 0) { $aiConfidence = 0 }
                            if ($aiConfidence -gt 1) { $aiConfidence = 1 }
                        }
                    } catch {}
                    $aiFindings = if ($parsed.evidence) { @($parsed.evidence | ForEach-Object { "$_" }) } else { @($localResult.Findings) }
                    $aiRecs = if ($parsed.recommended_actions) { @($parsed.recommended_actions | ForEach-Object { "$_" }) } else { @($localResult.Recommendations) }
                    $headline = ("PC評価: {0} ({1}/100)" -f $aiEval, $score)
                    return [PSCustomObject]@{
                        Source = "Anthropic"
                        Headline = $headline
                        Summary = $aiSummary
                        Evaluation = $aiEval
                        EvaluationCode = $localResult.EvaluationCode
                        Confidence = $aiConfidence
                        DataTimestamp = $dataTimestamp
                        PromptVersion = $promptVersion
                        InputMetrics = $inputMetrics
                        Findings = $aiFindings
                        Recommendations = $aiRecs
                        Narrative = $text
                    }
                }
                $fallbackReason = "Anthropic response was not valid JSON"
            }
            if (-not $fallbackReason) { $fallbackReason = "Anthropic returned empty content" }
        } catch {
            $fallbackReason = "Anthropic request failed: $($_.Exception.Message)"
        }
    } else {
        $fallbackReason = "Anthropic API key not provided"
    }
    if ($fallbackReason) {
        $localResult | Add-Member -NotePropertyName FallbackReason -NotePropertyValue $fallbackReason -Force
    }
    return $localResult
}

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER Path
Parameter description

.PARAMETER HealthScore
Parameter description

.PARAMETER Snapshot
Parameter description

.PARAMETER AIDiagnosis
Parameter description

.PARAMETER EventAnomaly
Parameter description

.PARAMETER BootTrend
Parameter description

.PARAMETER AgentSummary
Parameter description

.PARAMETER TaskResults
Parameter description

.PARAMETER DiskBefore
Parameter description

.PARAMETER DiskAfter
Parameter description

.PARAMETER HookHistory
Parameter description

.PARAMETER McpResults
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function Export-PowerBIDashboardJson {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [pscustomobject]$HealthScore,
        [pscustomobject]$Snapshot,
        [pscustomobject]$AIDiagnosis,
        [pscustomobject]$EventAnomaly,
        [pscustomobject]$BootTrend,
        [pscustomobject]$AgentSummary,
        [object[]]$TaskResults,
        [double]$DiskBefore,
        [double]$DiskAfter,
        [object[]]$HookHistory,
        [object[]]$McpResults
    )

    $obj = [PSCustomObject]@{
        schemaVersion = "1.0"
        generatedAt = (Get-Date).ToUniversalTime().ToString("o")
        hostName = $env:COMPUTERNAME
        userName = $env:USERNAME
        score = if ($HealthScore) { $HealthScore.Score } else { $null }
        scoreStatus = if ($HealthScore) { $HealthScore.Status } else { $null }
        aiEvaluation = if ($AIDiagnosis) { $AIDiagnosis.Evaluation } else { $null }
        aiSummary = if ($AIDiagnosis -and $AIDiagnosis.PSObject.Properties["Summary"]) { $AIDiagnosis.Summary } else { $null }
        aiHeadline = if ($AIDiagnosis) { $AIDiagnosis.Headline } else { $null }
        aiPromptVersion = if ($AIDiagnosis -and $AIDiagnosis.PSObject.Properties["PromptVersion"]) { $AIDiagnosis.PromptVersion } else { $null }
        aiConfidence = if ($AIDiagnosis -and $AIDiagnosis.PSObject.Properties["Confidence"]) { $AIDiagnosis.Confidence } else { $null }
        aiDataTimestamp = if ($AIDiagnosis -and $AIDiagnosis.PSObject.Properties["DataTimestamp"]) { $AIDiagnosis.DataTimestamp } else { $null }
        aiFallbackReason = if ($AIDiagnosis -and $AIDiagnosis.PSObject.Properties["FallbackReason"]) { $AIDiagnosis.FallbackReason } else { $null }
        aiInputMetrics = if ($AIDiagnosis -and $AIDiagnosis.PSObject.Properties["InputMetrics"]) { $AIDiagnosis.InputMetrics } else { $null }
        aiFindings = if ($AIDiagnosis) { @($AIDiagnosis.Findings) } else { @() }
        aiRecommendations = if ($AIDiagnosis) { @($AIDiagnosis.Recommendations) } else { @() }
        updatePending = if ($Snapshot -and $Snapshot.PSObject.Properties["updateDiagnostics"] -and $Snapshot.updateDiagnostics -and $Snapshot.updateDiagnostics.PSObject.Properties["PendingCount"]) { $Snapshot.updateDiagnostics.PendingCount } else { $null }
        eventAnomalyStatus = if ($EventAnomaly) { $EventAnomaly.Status } else { $null }
        eventSpikeRatio = if ($EventAnomaly) { $EventAnomaly.SpikeRatio } else { $null }
        bootAvgDurationMs = if ($BootTrend) { $BootTrend.AvgBootDurationMs } else { $null }
        standard_execution = [PSCustomObject]@{
            totalTasks = @($TaskResults).Count
            successTasks = @($TaskResults | Where-Object { $_.Status -eq "OK" }).Count
            failedTasks = @($TaskResults | Where-Object { $_.Status -eq "NG" }).Count
            skippedTasks = @($TaskResults | Where-Object { $_.Status -eq "SKIP" }).Count
            diskBeforeGB = if ($DiskBefore -or $DiskBefore -eq 0) { [math]::Round([double]$DiskBefore, 2) } else { $null }
            diskAfterGB = if ($DiskAfter -or $DiskAfter -eq 0) { [math]::Round([double]$DiskAfter, 2) } else { $null }
            diskDeltaGB = if (($DiskBefore -or $DiskBefore -eq 0) -and ($DiskAfter -or $DiskAfter -eq 0)) { [math]::Round(([double]$DiskAfter - [double]$DiskBefore), 2) } else { $null }
            tasks = @($TaskResults | ForEach-Object { [PSCustomObject]@{ id = $_.Id; name = $_.Name; status = $_.Status; durationSec = $_.Duration; error = $_.Error } })
        }
        agent_summary = if ($AgentSummary) {
            [PSCustomObject]@{
                schemaVersion = "1.1"
                runId = if ($AgentSummary.runId) { $AgentSummary.runId } else { $null }
                transactionId = if ($AgentSummary.PSObject.Properties["transactionId"]) { $AgentSummary.transactionId } else { $null }
                profile = if ($AgentSummary.planner -and $AgentSummary.planner.PSObject.Properties["profile"]) { $AgentSummary.planner.profile } else { $null }
                runMode = if ($AgentSummary.planner -and $AgentSummary.planner.PSObject.Properties["runMode"]) { $AgentSummary.planner.runMode } else { $null }
                overall = if ($AgentSummary.analyzer) { $AgentSummary.analyzer.overall } else { $null }
                highRiskCount = if ($AgentSummary.analyzer) { $AgentSummary.analyzer.highRiskCount } else { $null }
                mediumRiskCount = if ($AgentSummary.analyzer) { $AgentSummary.analyzer.mediumRiskCount } else { $null }
                workers = if ($AgentSummary.collector) { @($AgentSummary.collector.workers) } else { @() }
                recommendedActions = if ($AgentSummary.remediator) { @($AgentSummary.remediator.recommendedActions) } else { @() }
                dag = if ($AgentSummary.dagExecution) {
                    [PSCustomObject]@{
                        levelCount = $AgentSummary.dagExecution.levelCount
                        totalNodes = $AgentSummary.dagExecution.totalNodes
                        successNodes = $AgentSummary.dagExecution.successNodes
                        failedNodes = $AgentSummary.dagExecution.failedNodes
                        skippedNodes = $AgentSummary.dagExecution.skippedNodes
                    }
                } else { $null }
                metrics = if ($AgentSummary.qualityMetrics) {
                    [PSCustomObject]@{
                        agents = if ($AgentSummary.qualityMetrics.agents) { @($AgentSummary.qualityMetrics.agents) } else { @() }
                    }
                } else { $null }
                mcp = if ($AgentSummary.PSObject.Properties["mcp"] -and $AgentSummary.mcp) {
                    [PSCustomObject]@{
                        total = @($AgentSummary.mcp).Count
                        failed = @($AgentSummary.mcp | Where-Object { $_.status -eq "Failed" }).Count
                    }
                } else { $null }
            }
        } else { $null }
        agent_dag_timeline = if ($AgentSummary -and $AgentSummary.PSObject.Properties["dagTimeline"]) { @($AgentSummary.dagTimeline) } else { @() }
        hook_timeline = if ($AgentSummary -and $AgentSummary.PSObject.Properties["hookTimeline"]) { @($AgentSummary.hookTimeline) } else { @() }
        hooks = [PSCustomObject]@{
            total = @($HookHistory).Count
            failed = @($HookHistory | Where-Object { $_.status -eq "Failed" }).Count
        }
        mcp = [PSCustomObject]@{
            total = @($McpResults).Count
            failed = @($McpResults | Where-Object { $_.status -eq "Failed" }).Count
            transactionIds = @($McpResults | ForEach-Object { if ($_.PSObject.Properties["transactionId"]) { "$($_.transactionId)" } } | Where-Object { $_ } | Select-Object -Unique)
        }
    }
    $obj | ConvertTo-Json -Depth 10 | Set-Content -Path $Path -Encoding $script:_enc
    return $Path
}

Export-ModuleMember -Function `
    Start-RepairGuardrails,Test-RepairAllowListPath,Complete-RepairGuardrails, `
    Get-UpdateErrorClassification,Test-M365Connectivity,Get-EventLogAnomaly, `
    Update-BootShutdownTrend,Set-MonthlyMaintenanceTask,Remove-MonthlyMaintenanceTask, `
    Invoke-WinRMRemoteDiagnosticsBatch,Invoke-AssetCentralAggregation,Invoke-AIDiagnosis,Normalize-AiEvaluation, `
    Export-PowerBIDashboardJson



