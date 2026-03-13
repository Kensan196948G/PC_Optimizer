Set-StrictMode -Version Latest
$script:_enc = if ($PSVersionTable.PSVersion.Major -ge 7) { 'utf8NoBOM' } else { 'UTF8' }

function Get-StableSha256Hash {
    [CmdletBinding()]
    param([string]$Text)

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes([string]$Text)
        $hash = $sha.ComputeHash($bytes)
        return ([BitConverter]::ToString($hash) -replace "-", "").ToLowerInvariant()
    } finally {
        $sha.Dispose()
    }
}

function New-StandardHookPayload {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$EventName,
        [string]$RunId,
        [pscustomobject]$Context,
        [ValidateSet("Pending", "Success", "Failed", "Skipped")][string]$Status = "Pending",
        [string]$TransactionId = ""
    )

    $ctx = if ($Context) { $Context } else { [PSCustomObject]@{} }
    $tx = if ($TransactionId) { $TransactionId } elseif ($RunId) { $RunId } else { [guid]::NewGuid().ToString("N") }
    $correlationId = if ($ctx.PSObject.Properties["correlationId"]) { "$($ctx.correlationId)" } elseif ($RunId) { $RunId } else { [guid]::NewGuid().ToString("N") }

    return [PSCustomObject]@{
        schemaVersion = "1.1"
        eventId = [guid]::NewGuid().ToString("N")
        eventName = $EventName
        transactionId = $tx
        generatedAt = (Get-Date).ToString("s")
        runId = $RunId
        correlationId = $correlationId
        agentId = if ($ctx.PSObject.Properties["agentId"]) { "$($ctx.agentId)" } else { $null }
        nodeId = if ($ctx.PSObject.Properties["nodeId"]) { "$($ctx.nodeId)" } else { $null }
        stage = if ($ctx.PSObject.Properties["stage"]) { "$($ctx.stage)" } else { $null }
        status = $Status
        payload = $ctx
    }
}

function Convert-HookEntryToSiemLine {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][pscustomobject]$Entry,
        [ValidateSet("jsonl", "cef", "leef")][string]$Format = "jsonl"
    )

    $sev = if ($Entry.status -eq "Success") { 3 } else { 7 }
    if ($Format -eq "cef") {
        return ("CEF:0|PC_Optimizer|AgentTeams|1.1|{0}|{1}|{2}|rt={3} cs1Label=tx cs1={4} cs2Label=node cs2={5} msg={6}" -f `
            $Entry.event, $Entry.action, $sev, $Entry.finishedAt, $Entry.hookPayload.transactionId, $Entry.hookPayload.nodeId, ([string]$Entry.detail).Replace("|", "/"))
    }
    if ($Format -eq "leef") {
        return ("LEEF:2.0|PC_Optimizer|AgentTeams|1.1|{0}|sev={1}`ttx={2}`trunId={3}`tnode={4}`tmsg={5}" -f `
            $Entry.event, $sev, $Entry.hookPayload.transactionId, $Entry.hookPayload.runId, $Entry.hookPayload.nodeId, ([string]$Entry.detail).Replace("`t", " "))
    }
    return ($Entry | ConvertTo-Json -Depth 12 -Compress)
}

function Export-HookSiemLines {
    [CmdletBinding()]
    param(
        [object[]]$Entries,
        [pscustomobject]$HooksConfig,
        [string]$LogsDir,
        [string]$RunId
    )

    if (-not $HooksConfig -or -not $HooksConfig.PSObject.Properties["siem"]) { return }
    $siem = $HooksConfig.siem
    if ($siem.PSObject.Properties["enabled"] -and -not [bool]$siem.enabled) { return }

    $formats = @("jsonl")
    if ($siem.PSObject.Properties["formats"] -and @($siem.formats).Count -gt 0) {
        $formats = @($siem.formats | ForEach-Object { "$_".ToLowerInvariant() })
    }

    $baseDir = if ($LogsDir -and (Test-Path $LogsDir)) { $LogsDir } else { "." }
    $outDir = if ($siem.PSObject.Properties["outputDir"] -and $siem.outputDir) {
        if ([IO.Path]::IsPathRooted("$($siem.outputDir)")) { "$($siem.outputDir)" } else { Join-Path $baseDir "$($siem.outputDir)" }
    } else {
        Join-Path $baseDir "siem"
    }
    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

    foreach ($f in $formats) {
        if ($f -notin @("jsonl", "cef", "leef")) { continue }
        $path = Join-Path $outDir ("HookEvents_{0}_{1}.{2}" -f $RunId, (Get-Date -Format "yyyyMMdd"), $f)
        $lines = @($Entries | ForEach-Object { Convert-HookEntryToSiemLine -Entry $_ -Format $f })
        if (@($lines).Count -gt 0) {
            Add-Content -Path $path -Value $lines -Encoding $script:_enc
        }
    }
}

function Get-AgentTeamsProfilePreset {
    [CmdletBinding()]
    param([string]$RunMode)

    $mode = if ($RunMode) { "$RunMode".ToLowerInvariant() } else { "repair" }
    $hookEvents = @("pre_task", "post_task", "on_error", "on_fallback", "on_report")
    $mcpNames = @("local-archive", "slack-notify", "teams-notify", "servicenow-ticket", "jira-ticket")

    if ($mode -eq "diagnose") {
        return [PSCustomObject]@{
            name = "diagnose"
            maxParallel = 8
            nodeParallel = 3
            remediatorContinueOnFail = $true
            enabledHookEvents = $hookEvents
            enabledMcpProviders = $mcpNames
        }
    }

    return [PSCustomObject]@{
        name = "repair"
        maxParallel = 6
        nodeParallel = 2
        remediatorContinueOnFail = $false
        enabledHookEvents = $hookEvents
        enabledMcpProviders = $mcpNames
    }
}

function Resolve-AgentPluginPath {
    [CmdletBinding()]
    param(
        [string]$Role,
        [string]$AgentId,
        [string]$PluginsRoot
    )

    if (-not $PluginsRoot) { return $null }
    $candidates = @(
        (Join-Path $PluginsRoot ("{0}.{1}.psm1" -f $Role, $AgentId)),
        (Join-Path $PluginsRoot ("{0}.psm1" -f $AgentId)),
        (Join-Path $PluginsRoot ("{0}.psm1" -f $Role))
    )
    foreach ($c in $candidates) {
        if (Test-Path $c) { return $c }
    }
    return $null
}

function Get-NodePolicy {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][pscustomobject]$Node,
        [int]$GlobalMaxParallel = 6
    )

    $maxAttempts = 1
    $backoffSeconds = 1
    $jitterMs = 0
    $continueOnFail = $false
    $timeoutSeconds = 120
    $priority = 100
    $maxParallel = [Math]::Max($GlobalMaxParallel, 1)
    $parallelGroup = if ($Node.PSObject.Properties["parallelGroup"]) { "$($Node.parallelGroup)" } else { "$($Node.role)" }

    if ($Node.PSObject.Properties["retryPolicy"] -and $Node.retryPolicy) {
        if ($Node.retryPolicy.PSObject.Properties["maxAttempts"]) { $maxAttempts = [Math]::Max([int]$Node.retryPolicy.maxAttempts, 1) }
        if ($Node.retryPolicy.PSObject.Properties["backoffSeconds"]) { $backoffSeconds = [Math]::Max([int]$Node.retryPolicy.backoffSeconds, 0) }
        if ($Node.retryPolicy.PSObject.Properties["jitterMs"]) { $jitterMs = [Math]::Max([int]$Node.retryPolicy.jitterMs, 0) }
    }
    if ($Node.PSObject.Properties["continueOnFail"]) { $continueOnFail = [bool]$Node.continueOnFail }
    if ($Node.PSObject.Properties["timeoutSeconds"]) { $timeoutSeconds = [Math]::Max([int]$Node.timeoutSeconds, 1) }
    if ($Node.PSObject.Properties["priority"]) { $priority = [int]$Node.priority }
    if ($Node.PSObject.Properties["maxParallel"]) { $maxParallel = [Math]::Max([int]$Node.maxParallel, 1) }

    return [PSCustomObject]@{
        maxAttempts = $maxAttempts
        backoffSeconds = $backoffSeconds
        jitterMs = $jitterMs
        continueOnFail = $continueOnFail
        timeoutSeconds = $timeoutSeconds
        priority = $priority
        maxParallel = $maxParallel
        parallelGroup = $parallelGroup
    }
}

function New-DefaultAgentDagNodes {
    [CmdletBinding()]
    param([string]$RunMode)

    $preset = Get-AgentTeamsProfilePreset -RunMode $RunMode
    return @(
        [PSCustomObject]@{ id = "planner"; role = "planner"; agentId = "Planner"; dependsOn = @(); priority = 500; maxParallel = 1; parallelGroup = "planner"; timeoutSeconds = 30; retryPolicy = [PSCustomObject]@{ maxAttempts = 1; backoffSeconds = 0; jitterMs = 0 }; continueOnFail = $false },
        [PSCustomObject]@{ id = "collector.security"; role = "collector"; agentId = "SecurityAgent"; dependsOn = @("planner"); priority = 400; maxParallel = $preset.nodeParallel; parallelGroup = "collector"; timeoutSeconds = 120; retryPolicy = [PSCustomObject]@{ maxAttempts = 2; backoffSeconds = 1; jitterMs = 120 }; continueOnFail = $true },
        [PSCustomObject]@{ id = "collector.network"; role = "collector"; agentId = "NetworkAgent"; dependsOn = @("planner"); priority = 390; maxParallel = $preset.nodeParallel; parallelGroup = "collector"; timeoutSeconds = 120; retryPolicy = [PSCustomObject]@{ maxAttempts = 2; backoffSeconds = 1; jitterMs = 120 }; continueOnFail = $true },
        [PSCustomObject]@{ id = "collector.update"; role = "collector"; agentId = "UpdateAgent"; dependsOn = @("planner"); priority = 380; maxParallel = $preset.nodeParallel; parallelGroup = "collector"; timeoutSeconds = 120; retryPolicy = [PSCustomObject]@{ maxAttempts = 2; backoffSeconds = 1; jitterMs = 120 }; continueOnFail = $true },
        [PSCustomObject]@{ id = "analyzer.security"; role = "analyzer"; agentId = "SecurityAgent"; dependsOn = @("collector.security"); priority = 300; maxParallel = $preset.nodeParallel; parallelGroup = "analyzer"; timeoutSeconds = 120; retryPolicy = [PSCustomObject]@{ maxAttempts = 2; backoffSeconds = 1; jitterMs = 150 }; continueOnFail = $true },
        [PSCustomObject]@{ id = "analyzer.network"; role = "analyzer"; agentId = "NetworkAgent"; dependsOn = @("collector.network"); priority = 290; maxParallel = $preset.nodeParallel; parallelGroup = "analyzer"; timeoutSeconds = 120; retryPolicy = [PSCustomObject]@{ maxAttempts = 2; backoffSeconds = 1; jitterMs = 150 }; continueOnFail = $true },
        [PSCustomObject]@{ id = "analyzer.update"; role = "analyzer"; agentId = "UpdateAgent"; dependsOn = @("collector.update"); priority = 280; maxParallel = $preset.nodeParallel; parallelGroup = "analyzer"; timeoutSeconds = 120; retryPolicy = [PSCustomObject]@{ maxAttempts = 2; backoffSeconds = 1; jitterMs = 150 }; continueOnFail = $true },
        [PSCustomObject]@{ id = "analyzer.aggregate"; role = "analyzer"; agentId = "Analyzer"; dependsOn = @("analyzer.security", "analyzer.network", "analyzer.update"); priority = 250; maxParallel = 1; parallelGroup = "analyzer"; timeoutSeconds = 120; retryPolicy = [PSCustomObject]@{ maxAttempts = 2; backoffSeconds = 1; jitterMs = 150 }; continueOnFail = $false },
        [PSCustomObject]@{ id = "remediator"; role = "remediator"; agentId = "Remediator"; dependsOn = @("analyzer.aggregate"); priority = 200; maxParallel = 1; parallelGroup = "serial"; timeoutSeconds = 60; retryPolicy = [PSCustomObject]@{ maxAttempts = 1; backoffSeconds = 0; jitterMs = 0 }; continueOnFail = $preset.remediatorContinueOnFail },
        [PSCustomObject]@{ id = "reporter"; role = "reporter"; agentId = "Reporter"; dependsOn = @("remediator"); priority = 100; maxParallel = 1; parallelGroup = "serial"; timeoutSeconds = 60; retryPolicy = [PSCustomObject]@{ maxAttempts = 1; backoffSeconds = 0; jitterMs = 0 }; continueOnFail = $false }
    )
}
function Invoke-BuiltinAgentNode {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][pscustomobject]$Node,
        [string]$RunId,
        [pscustomobject]$ModuleSnapshot,
        [hashtable]$CollectorByAgent,
        [hashtable]$AnalyzerByAgent,
        [pscustomobject]$HealthScore,
        [pscustomobject]$AIDiagnosis
    )

    $role = "$($Node.role)".ToLowerInvariant()
    $agent = "$($Node.agentId)"

    if ($role -eq "planner") {
        return [PSCustomObject]@{ status = "Success"; risk = "Low"; payload = [PSCustomObject]@{ planned = $true; runId = $RunId }; message = "planner-ready" }
    }

    if ($role -eq "collector") {
        $payload = $null
        if ($agent -eq "SecurityAgent") { $payload = $ModuleSnapshot.securityDiagnostics }
        elseif ($agent -eq "NetworkAgent") { $payload = $ModuleSnapshot.networkDiagnostics }
        elseif ($agent -eq "UpdateAgent") { $payload = $ModuleSnapshot.updateDiagnostics }

        $risk = "Low"
        if ($agent -eq "SecurityAgent" -and $payload) {
            $defenderNg = ($payload.PSObject.Properties["Defender"] -and $payload.Defender -ne "Enabled")
            $firewallNg = ($payload.PSObject.Properties["Firewall"] -and $payload.Firewall -ne "Enabled")
            if ($defenderNg -or $firewallNg) { $risk = "High" }
        } elseif ($agent -eq "NetworkAgent" -and $payload) {
            if ((-not $payload.PSObject.Properties["PrimaryIPv4"]) -or [string]::IsNullOrWhiteSpace("$($payload.PrimaryIPv4)")) { $risk = "Medium" }
        } elseif ($agent -eq "UpdateAgent" -and $payload) {
            if ($payload.PSObject.Properties["WindowsUpdate"] -and $payload.WindowsUpdate -ne "Compliant") { $risk = "Medium" }
        }
        return [PSCustomObject]@{ status = "Success"; risk = $risk; payload = $payload; message = "collector-complete" }
    }

    if ($role -eq "analyzer") {
        if ($Node.id -eq "analyzer.aggregate") {
            $all = @($AnalyzerByAgent.Values | Where-Object { $_ })
            $high = @($all | Where-Object { $_.risk -eq "High" }).Count
            $medium = @($all | Where-Object { $_.risk -eq "Medium" }).Count
            $overall = if ($high -gt 0) { "Critical" } elseif ($medium -gt 0) { "Warning" } else { "Good" }
            $aggRisk = if ($overall -eq "Critical") { "High" } elseif ($overall -eq "Warning") { "Medium" } else { "Low" }
            return [PSCustomObject]@{
                status = "Success"
                risk = $aggRisk
                payload = [PSCustomObject]@{ overall = $overall; highRiskCount = $high; mediumRiskCount = $medium }
                message = "aggregate-complete"
            }
        }

        $sourceCollector = $CollectorByAgent[$agent]
        if (-not $sourceCollector) {
            return [PSCustomObject]@{ status = "Failed"; risk = "Unknown"; payload = $null; message = "missing-collector-data" }
        }
        return [PSCustomObject]@{
            status = "Success"
            risk = "$($sourceCollector.risk)"
            payload = [PSCustomObject]@{ sourceRisk = "$($sourceCollector.risk)" }
            message = "analyzer-complete"
        }
    }

    if ($role -eq "remediator") {
        $agg = $AnalyzerByAgent["Analyzer"]
        $actions = New-Object System.Collections.Generic.List[string]
        if ($agg -and $agg.payload -and $agg.payload.overall -eq "Critical") {
            [void]$actions.Add("Prioritize security and update remediation immediately.")
        } elseif ($agg -and $agg.payload -and $agg.payload.overall -eq "Warning") {
            [void]$actions.Add("Address medium risks on weekly maintenance cycle.")
        }
        if ($actions.Count -eq 0) {
            [void]$actions.Add("Keep preventive maintenance and monitoring.")
        }
        return [PSCustomObject]@{
            status = "Success"
            risk = "Low"
            payload = [PSCustomObject]@{ recommendedActions = @($actions) }
            message = "remediator-complete"
        }
    }

    if ($role -eq "reporter") {
        return [PSCustomObject]@{
            status = "Success"
            risk = "Low"
            payload = [PSCustomObject]@{
                healthScore = if ($HealthScore) { $HealthScore.Score } else { $null }
                aiEvaluation = if ($AIDiagnosis) { $AIDiagnosis.Evaluation } else { $null }
                generatedAt = (Get-Date).ToString("s")
            }
            message = "reporter-complete"
        }
    }

    return [PSCustomObject]@{ status = "Failed"; risk = "Unknown"; payload = $null; message = "unsupported-role" }
}

function Invoke-AgentPluginOrBuiltin {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][pscustomobject]$Node,
        [string]$RunId,
        [pscustomobject]$ModuleSnapshot,
        [hashtable]$CollectorByAgent,
        [hashtable]$AnalyzerByAgent,
        [pscustomobject]$HealthScore,
        [pscustomobject]$AIDiagnosis,
        [string]$PluginsRoot
    )

    $plugin = Resolve-AgentPluginPath -Role "$($Node.role)" -AgentId "$($Node.agentId)" -PluginsRoot $PluginsRoot
    if ($plugin) {
        try {
            Import-Module $plugin -Force -ErrorAction Stop
            if (Get-Command Invoke-AgentPlugin -ErrorAction SilentlyContinue) {
                $ctx = @{
                    RunId = $RunId
                    ModuleSnapshot = $ModuleSnapshot
                    CollectorByAgent = $CollectorByAgent
                    AnalyzerByAgent = $AnalyzerByAgent
                    HealthScore = $HealthScore
                    AIDiagnosis = $AIDiagnosis
                }
                $result = Invoke-AgentPlugin -Role "$($Node.role)" -AgentId "$($Node.agentId)" -Node $Node -Context $ctx
                if ($result) {
                    if ($result.PSObject.Properties["status"] -and $result.PSObject.Properties["risk"] -and $result.PSObject.Properties["message"]) {
                        return [PSCustomObject]@{
                            status = "$($result.status)"
                            risk = "$($result.risk)"
                            payload = if ($result.PSObject.Properties["payload"]) { $result.payload } else { $null }
                            message = "$($result.message)"
                        }
                    }
                    return [PSCustomObject]@{ status = "Failed"; risk = "Unknown"; payload = $null; message = "plugin-result-schema-invalid" }
                }
            }
        } catch {
            # Fallback to builtin
        }
    }

    return Invoke-BuiltinAgentNode -Node $Node -RunId $RunId -ModuleSnapshot $ModuleSnapshot -CollectorByAgent $CollectorByAgent -AnalyzerByAgent $AnalyzerByAgent -HealthScore $HealthScore -AIDiagnosis $AIDiagnosis
}

function Invoke-AgentNodeBatchRunspace {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]]$Nodes,
        [Parameter(Mandatory)][int]$MaxParallel,
        [Parameter(Mandatory)][string]$ModulePath,
        [Parameter(Mandatory)][string]$RunId,
        [pscustomobject]$ModuleSnapshot,
        [hashtable]$CollectorByAgent,
        [hashtable]$AnalyzerByAgent,
        [pscustomobject]$HealthScore,
        [pscustomobject]$AIDiagnosis,
        [string]$PluginsRoot
    )

    $items = @($Nodes)
    if (@($items).Count -eq 0) { return @() }
    $pool = [RunspaceFactory]::CreateRunspacePool(1, [Math]::Max($MaxParallel, 1))
    $pool.Open()

    $jobs = @()
    foreach ($node in $items) {
        $ps = [powershell]::Create()
        $ps.RunspacePool = $pool
        [void]$ps.AddScript({
            param($modulePath, $nodeArg, $runIdArg, $snapshotArg, $collectorArg, $analyzerArg, $healthArg, $aiArg, $pluginsRootArg)
            Import-Module $modulePath -Force | Out-Null
            $r = Invoke-AgentPluginOrBuiltin -Node $nodeArg -RunId $runIdArg -ModuleSnapshot $snapshotArg -CollectorByAgent $collectorArg -AnalyzerByAgent $analyzerArg -HealthScore $healthArg -AIDiagnosis $aiArg -PluginsRoot $pluginsRootArg
            return ($r | ConvertTo-Json -Depth 14 -Compress)
        })
        [void]$ps.AddArgument($ModulePath)
        [void]$ps.AddArgument($node)
        [void]$ps.AddArgument($RunId)
        [void]$ps.AddArgument($ModuleSnapshot)
        [void]$ps.AddArgument($CollectorByAgent)
        [void]$ps.AddArgument($AnalyzerByAgent)
        [void]$ps.AddArgument($HealthScore)
        [void]$ps.AddArgument($AIDiagnosis)
        [void]$ps.AddArgument($PluginsRoot)
        $jobs += [PSCustomObject]@{
            node = $node
            ps = $ps
            handle = $ps.BeginInvoke()
            startedAt = Get-Date
        }
    }

    $results = @()
    foreach ($j in $jobs) {
        $finishedAt = Get-Date
        try {
            $timeoutSec = 120
            if ($j.node -and $j.node.PSObject.Properties["timeoutSeconds"]) {
                $timeoutSec = [Math]::Max([int]$j.node.timeoutSeconds, 1)
            }
            $waitOk = $j.handle.AsyncWaitHandle.WaitOne([TimeSpan]::FromSeconds($timeoutSec))
            if (-not $waitOk) {
                try { $j.ps.Stop() } catch {}
                $finishedAt = Get-Date
                $results += [PSCustomObject]@{
                    node = $j.node
                    startedAt = $j.startedAt
                    finishedAt = $finishedAt
                    result = [PSCustomObject]@{
                        status = "Failed"
                        risk = "Unknown"
                        payload = $null
                        message = "runspace-timeout"
                        timedOut = $true
                    }
                }
                continue
            }

            $raw = $j.ps.EndInvoke($j.handle)
            $json = @($raw) -join ""
            $obj = if ([string]::IsNullOrWhiteSpace($json)) {
                [PSCustomObject]@{ status = "Failed"; risk = "Unknown"; payload = $null; message = "empty-runspace-result" }
            } else {
                $json | ConvertFrom-Json
            }
            $results += [PSCustomObject]@{
                node = $j.node
                startedAt = $j.startedAt
                finishedAt = $finishedAt
                result = $obj
            }
        } catch {
            $results += [PSCustomObject]@{
                node = $j.node
                startedAt = $j.startedAt
                finishedAt = $finishedAt
                result = [PSCustomObject]@{ status = "Failed"; risk = "Unknown"; payload = $null; message = "runspace-exception: $($_.Exception.Message)" }
            }
        } finally {
            $j.ps.Dispose()
        }
    }

    $pool.Close()
    $pool.Dispose()
    return @($results)
}

function Get-HookAckLedgerContext {
    [CmdletBinding()]
    param([string]$LogsDir)

    $baseDir = if ($LogsDir -and (Test-Path $LogsDir)) { $LogsDir } else { $env:TEMP }
    $ackDir = Join-Path (Join-Path $baseDir "hooks") "ack"
    if (-not (Test-Path $ackDir)) { New-Item -ItemType Directory -Path $ackDir -Force | Out-Null }
    $path = Join-Path $ackDir "Hook_AckLedger.json"
    if (-not (Test-Path $path)) {
        @() | ConvertTo-Json -Depth 8 | Set-Content -Path $path -Encoding $script:_enc
    }
    return [PSCustomObject]@{ ackDir = $ackDir; ledgerPath = $path }
}

function Update-HookAckLedger {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][pscustomobject]$AckContext,
        [Parameter(Mandatory)][pscustomobject]$Entry
    )
    $rows = @()
    if (Test-Path $AckContext.ledgerPath) {
        try { $rows = @(Get-Content -Path $AckContext.ledgerPath -Raw -Encoding utf8 -ErrorAction Stop | ConvertFrom-Json) } catch { $rows = @() }
    }
    $rows += $Entry
    $rows = @($rows | Select-Object -Last 20000)
    $rows | ConvertTo-Json -Depth 12 | Set-ContentWithRetry -Path $AckContext.ledgerPath -Encoding $script:_enc
}

function Invoke-HookAckResync {
    [CmdletBinding()]
    param(
        [string]$LogsDir,
        [int]$MaxItems = 20
    )

    $ackContext = Get-HookAckLedgerContext -LogsDir $LogsDir
    $queueContext = Get-HookQueueContext -LogsDir $LogsDir
    $rows = @()
    if (Test-Path $ackContext.ledgerPath) {
        try { $rows = @(Get-Content -Path $ackContext.ledgerPath -Raw -Encoding utf8 -ErrorAction Stop | ConvertFrom-Json) } catch { $rows = @() }
    }
    $now = Get-Date
    $targets = @($rows | Where-Object {
            $_ -and $_.PSObject.Properties["ackState"] -and $_.ackState -eq "PendingResync" -and
            $_.PSObject.Properties["retryAfter"] -and ([datetime]$_.retryAfter) -le $now
        } | Select-Object -First ([Math]::Max($MaxItems, 1)))
    $resynced = @()
    foreach ($t in $targets) {
        $exists = @(Get-ChildItem -Path $queueContext.queueDir -Filter "*_${($t.event)}_${($t.action)}.json" -ErrorAction SilentlyContinue).Count -gt 0
        if (-not $exists -and $t.PSObject.Properties["payload"] -and $t.payload) {
            $q = New-HookQueueItem -ContextInfo $queueContext -Payload $t.payload -EventName "$($t.event)" -ActionName "$($t.action)" -Type "$($t.type)"
            $resynced += [PSCustomObject]@{ queueId = $q.item.queueId; sequence = $q.item.sequence; event = $t.event; action = $t.action }
        }
        $t.ackState = "ResyncQueued"
        $t.updatedAt = (Get-Date).ToString("s")
    }
    $rows | ConvertTo-Json -Depth 12 | Set-ContentWithRetry -Path $ackContext.ledgerPath -Encoding $script:_enc
    return @($resynced)
}

function Get-HookQueueContext {
    [CmdletBinding()]
    param([string]$LogsDir)

    $baseDir = if ($LogsDir -and (Test-Path $LogsDir)) { $LogsDir } else { $env:TEMP }
    $hookDir = Join-Path $baseDir "hooks"
    $queueDir = Join-Path $hookDir "queue"
    $doneDir = Join-Path $queueDir "done"
    $statePath = Join-Path $queueDir "state.json"
    foreach ($d in @($hookDir, $queueDir, $doneDir)) {
        if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
    }
    if (-not (Test-Path $statePath)) {
        [PSCustomObject]@{ nextSequence = 1; lastDeliveredSequence = 0; updatedAt = (Get-Date).ToString("s") } | ConvertTo-Json -Depth 6 | Set-Content -Path $statePath -Encoding $script:_enc
    }
    return [PSCustomObject]@{
        hookDir = $hookDir
        queueDir = $queueDir
        doneDir = $doneDir
        statePath = $statePath
    }
}

function New-HookQueueItem {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][pscustomobject]$ContextInfo,
        [Parameter(Mandatory)][pscustomobject]$Payload,
        [Parameter(Mandatory)][string]$EventName,
        [Parameter(Mandatory)][string]$ActionName,
        [Parameter(Mandatory)][string]$Type
    )

    $state = Get-Content -Path $ContextInfo.statePath -Raw -Encoding utf8 | ConvertFrom-Json
    $seq = [int]$state.nextSequence
    $state.nextSequence = $seq + 1
    $state.updatedAt = (Get-Date).ToString("s")
    $state | ConvertTo-Json -Depth 6 | Set-Content -Path $ContextInfo.statePath -Encoding $script:_enc

    $id = [guid]::NewGuid().ToString("N")
    $item = [PSCustomObject]@{
        schemaVersion = "1.1"
        queueId = $id
        sequence = $seq
        event = $EventName
        action = $ActionName
        type = $Type
        payload = $Payload
        status = "Pending"
        attempts = 0
        createdAt = (Get-Date).ToString("s")
        updatedAt = (Get-Date).ToString("s")
        detail = ""
        attemptHistory = @()
    }
    $path = Join-Path $ContextInfo.queueDir ("{0:D12}_{1}_{2}.json" -f $seq, $EventName, $ActionName)
    $item | ConvertTo-Json -Depth 14 | Set-Content -Path $path -Encoding $script:_enc
    return [PSCustomObject]@{ item = $item; path = $path }
}

function Process-HookQueue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][pscustomobject]$ContextInfo,
        [Parameter(Mandatory)][pscustomobject]$HooksConfig,
        [int]$RetryCount = 1,
        [int]$RetryDelaySeconds = 1
    )

    $processed = @()
    $ackContext = Get-HookAckLedgerContext -LogsDir (Split-Path $ContextInfo.hookDir -Parent)
    $pending = @(Get-ChildItem -Path $ContextInfo.queueDir -Filter "*.json" -File -ErrorAction SilentlyContinue | Sort-Object Name)
    if (@($pending).Count -eq 0) { return @() }

    foreach ($file in $pending) {
        $q = Get-Content -Path $file.FullName -Raw -Encoding utf8 | ConvertFrom-Json
        # Normalize: ensure all expected properties exist for Set-StrictMode -Version Latest compatibility
        foreach ($prop in @('payload','type','status','attempts','detail','attemptHistory','queueId','sequence','event','action','createdAt','updatedAt')) {
            if (-not $q.PSObject.Properties[$prop]) { $q | Add-Member -NotePropertyName $prop -NotePropertyValue $null }
        }
        if ($null -eq $q.attempts) { $q.attempts = 0 }
        if ($null -eq $q.attemptHistory) { $q.attemptHistory = @() }
        if ($null -eq $q.detail) { $q.detail = "" }
        if ($null -eq $q.status) { $q.status = "Pending" }
        if (-not $q.payload) { continue }

        $detail = ""
        $ok = $false
        $attempt = [int]$q.attempts + 1
        $ackState = "None"
        $retryAfter = $null
        try {
            $a = $q.payload.payload
            if ($q.type -eq "command") {
                # Invoke-Expression による任意コード実行を防ぐためサポートしない
                throw "Hook type 'command' is not supported. Use 'webhook' or 'file' type."
            } elseif ($q.type -eq "webhook" -and $a.url) {
                $method = if ($a.method) { "$($a.method)" } else { "POST" }
                $headers = @{ "X-PCO-Hook-QueueId" = "$($q.queueId)"; "X-PCO-Hook-Sequence" = "$($q.sequence)" }
                Invoke-RestMethod -Uri "$($a.url)" -Method $method -Headers $headers -Body ($q.payload | ConvertTo-Json -Depth 12) -ContentType "application/json" -TimeoutSec 20 | Out-Null
                $detail = "posted:$($a.url)"
                $ackState = "Confirmed"
            } elseif ($q.type -eq "file") {
                $outPath = Join-Path $ContextInfo.hookDir ("HookPayload_{0}_{1}_{2}.json" -f $q.event, $q.action, (Get-Date -Format "yyyyMMddHHmmssfff"))
                $q.payload | ConvertTo-Json -Depth 12 | Set-Content -Path $outPath -Encoding $script:_enc
                $detail = "written:$outPath"
            } else {
                $detail = if ($a.message) { "$($a.message)" } else { "$($q.event) hook executed." }
            }
            $ok = $true
        } catch {
            $detail = "$($_.Exception.Message)"
            if ($q.type -eq "webhook") {
                $ackState = "PendingResync"
                $retryAfter = (Get-Date).AddMinutes(2).ToString("s")
            }
        }

        $q.attempts = $attempt
        $q.updatedAt = (Get-Date).ToString("s")
        $q.detail = $detail
        $histEntry = [PSCustomObject]@{
            attempt = $attempt
            status = if ($ok) { "Success" } else { "Failed" }
            detail = $detail
            at = (Get-Date).ToString("s")
        }
        if ($q.PSObject.Properties["attemptHistory"] -and $null -ne $q.attemptHistory) {
            $q.attemptHistory = @($q.attemptHistory) + $histEntry
        } else {
            $q.attemptHistory = @($histEntry)
        }

        if ($ok) {
            $q.status = "Success"
            $donePath = Join-Path $ContextInfo.doneDir ($file.Name)
            $q | ConvertTo-Json -Depth 14 | Set-ContentWithRetry -Path $donePath -Encoding $script:_enc
            Remove-Item -Path $file.FullName -Force -ErrorAction SilentlyContinue

            $state = Get-Content -Path $ContextInfo.statePath -Raw -Encoding utf8 | ConvertFrom-Json
            if ([int]$q.sequence -gt [int]$state.lastDeliveredSequence) {
                $state.lastDeliveredSequence = [int]$q.sequence
                $state.updatedAt = (Get-Date).ToString("s")
                $state | ConvertTo-Json -Depth 6 | Set-ContentWithRetry -Path $ContextInfo.statePath -Encoding $script:_enc
            }
        } else {
            $q.status = if ($attempt -ge $RetryCount) { "Failed" } else { "Pending" }
            $q | ConvertTo-Json -Depth 14 | Set-ContentWithRetry -Path $file.FullName -Encoding $script:_enc
            $qStatus = if ($q.PSObject.Properties["status"] -and $q.status) { "$($q.status)" } else { "Failed" }
            if ($qStatus -eq "Pending" -and $RetryDelaySeconds -gt 0) {
                Start-Sleep -Seconds ([int]([Math]::Pow(2, [Math]::Max($attempt - 1, 0)) * $RetryDelaySeconds))
                continue
            }
            # Order guarantee: stop on undeliverable item and leave later entries pending.
            $qEvent   = if ($q.PSObject.Properties["event"])  { "$($q.event)" }  else { "" }
            $qAction  = if ($q.PSObject.Properties["action"]) { "$($q.action)" } else { "" }
            $qType    = if ($q.PSObject.Properties["type"])   { "$($q.type)" }   else { "" }
            $qCreated = if ($q.PSObject.Properties["createdAt"]) { "$($q.createdAt)" } else { "" }
            $qSeq     = if ($q.PSObject.Properties["sequence"])  { [int]$q.sequence } else { 0 }
            $processed += [PSCustomObject]@{
                schemaVersion = "1.1"
                event = $qEvent
                action = $qAction
                type = $qType
                startedAt = $qCreated
                finishedAt = (Get-Date).ToString("s")
                status = "Failed"
                detail = $detail
                attempts = $attempt
                order = $qSeq
                hookPayload = $q.payload
                attemptHistory = @($q.attemptHistory)
            }
            break
        }

        if ($q.type -eq "webhook") {
            Update-HookAckLedger -AckContext $ackContext -Entry ([PSCustomObject]@{
                schemaVersion = "1.1"
                queueId = $q.queueId
                sequence = $q.sequence
                event = $q.event
                action = $q.action
                type = $q.type
                ackState = if ($ok) { "Confirmed" } else { $ackState }
                retryAfter = $retryAfter
                payload = $q.payload
                detail = $detail
                attempt = $attempt
                updatedAt = (Get-Date).ToString("s")
            })
        }

        $processed += [PSCustomObject]@{
            schemaVersion = "1.1"
            event = $q.event
            action = $q.action
            type = $q.type
            startedAt = $q.createdAt
            finishedAt = (Get-Date).ToString("s")
            status = "Success"
            detail = $detail
            attempts = $attempt
            order = [int]$q.sequence
            hookPayload = $q.payload
            attemptHistory = @($q.attemptHistory)
        }
    }
    return @($processed)
}

function Invoke-McpRollbackExecutor {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]]$McpResults,
        [Parameter(Mandatory)][string]$ReportsDir,
        [Parameter(Mandatory)][string]$RunId,
        [string]$TransactionId = ""
    )

    $rollbackDir = Join-Path (Join-Path $ReportsDir "mcp") "rollback"
    if (-not (Test-Path $rollbackDir)) { New-Item -ItemType Directory -Path $rollbackDir -Force | Out-Null }
    $rows = @()
    foreach ($r in @($McpResults | Where-Object { $_.status -eq "Failed" -or $_.status -eq "Success" })) {
        $status = "ManualRequired"
        $detail = if ($r.rollbackHint) { "$($r.rollbackHint)" } else { "No rollback hint." }
        if ($r.type -eq "file" -and $r.message -like "written:*") {
            $filePath = $r.message.Substring(8)
            if (Test-Path $filePath) {
                try {
                    Remove-Item -Path $filePath -Force -ErrorAction Stop
                    $status = "RolledBack"
                    $detail = "deleted:$filePath"
                } catch {
                    $status = "RollbackFailed"
                    $detail = "delete-failed:$($_.Exception.Message)"
                }
            } else {
                $status = "RolledBack"
                $detail = "already-missing:$filePath"
            }
        } elseif ($r.type -eq "slack") {
            if ($r.providerSnapshot -and $r.providerSnapshot.webhookUrl) {
                try {
                    Invoke-RestMethod -Uri "$($r.providerSnapshot.webhookUrl)" -Method Post -Body (@{ text = "PC_Optimizer rollback notice: transaction $($r.transactionId)" } | ConvertTo-Json) -ContentType "application/json" -TimeoutSec 30 | Out-Null
                    $status = "RolledBack"
                    $detail = "compensating-message:slack"
                } catch {
                    $status = "RollbackFailed"
                    $detail = "slack-failed:$($_.Exception.Message)"
                }
            }
        } elseif ($r.type -eq "teams") {
            if ($r.providerSnapshot -and $r.providerSnapshot.webhookUrl) {
                try {
                    Invoke-RestMethod -Uri "$($r.providerSnapshot.webhookUrl)" -Method Post -Body (@{ text = "PC_Optimizer rollback notice: transaction $($r.transactionId)" } | ConvertTo-Json) -ContentType "application/json" -TimeoutSec 30 | Out-Null
                    $status = "RolledBack"
                    $detail = "compensating-message:teams"
                } catch {
                    $status = "RollbackFailed"
                    $detail = "teams-failed:$($_.Exception.Message)"
                }
            }
        } elseif ($r.type -eq "servicenow") {
            if ($r.providerSnapshot -and $r.providerSnapshot.instanceUrl -and $r.providerSnapshot.table -and $r.referenceId) {
                try {
                    $url = ("{0}/api/now/table/{1}/{2}" -f "$($r.providerSnapshot.instanceUrl)".TrimEnd("/"), "$($r.providerSnapshot.table)", "$($r.referenceId)")
                    $headers = @{}
                    if ($r.providerSnapshot.token) { $headers["Authorization"] = "Bearer $($r.providerSnapshot.token)" }
                    $body = @{ state = "8"; close_notes = "Rolled back by PC_Optimizer transaction $($r.transactionId)" } | ConvertTo-Json
                    Invoke-RestMethod -Uri $url -Method Patch -Headers $headers -Body $body -ContentType "application/json" -TimeoutSec 30 | Out-Null
                    $status = "RolledBack"
                    $detail = "servicenow-updated:$($r.referenceId)"
                } catch {
                    $status = "RollbackFailed"
                    $detail = "servicenow-failed:$($_.Exception.Message)"
                }
            }
        } elseif ($r.type -eq "jira") {
            if ($r.providerSnapshot -and $r.providerSnapshot.url -and $r.referenceId) {
                try {
                    $headers = @{}
                    if ($r.providerSnapshot.userEmail -and $r.providerSnapshot.apiToken) {
                        $pair = "{0}:{1}" -f $r.providerSnapshot.userEmail, $r.providerSnapshot.apiToken
                        $headers["Authorization"] = "Basic $([Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($pair)))"
                    }
                    $commentUrl = ("{0}/rest/api/2/issue/{1}/comment" -f "$($r.providerSnapshot.url)".TrimEnd("/"), "$($r.referenceId)")
                    $commentBody = @{ body = "Rollback requested by PC_Optimizer transaction $($r.transactionId)." } | ConvertTo-Json
                    Invoke-RestMethod -Uri $commentUrl -Method Post -Headers $headers -Body $commentBody -ContentType "application/json" -TimeoutSec 30 | Out-Null
                    $status = "RolledBack"
                    $detail = "jira-commented:$($r.referenceId)"
                } catch {
                    $status = "RollbackFailed"
                    $detail = "jira-failed:$($_.Exception.Message)"
                }
            }
        }
        $rows += [PSCustomObject]@{
            transactionId = if ($TransactionId) { $TransactionId } else { $r.transactionId }
            operationId = $r.operationId
            provider = $r.name
            type = $r.type
            sourceStatus = $r.status
            rollbackStatus = $status
            detail = $detail
            at = (Get-Date).ToString("s")
        }
    }

    $path = Join-Path $rollbackDir ("MCP_Rollback_{0}_{1}.json" -f $RunId, (Get-Date -Format "yyyyMMddHHmmss"))
    [PSCustomObject]@{
        schemaVersion = "1.1"
        runId = $RunId
        transactionId = $TransactionId
        generatedAt = (Get-Date).ToString("s")
        items = @($rows)
    } | ConvertTo-Json -Depth 12 | Set-Content -Path $path -Encoding $script:_enc
    return [PSCustomObject]@{ path = $path; items = @($rows) }
}

function Invoke-AgentHookEvent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$EventName,
        [pscustomobject]$Context,
        [pscustomobject]$HooksConfig,
        [string]$RunId = "",
        [string]$LogsDir = "",
        [string]$TransactionId = ""
    )

    $history = @()
    if (-not $HooksConfig -or -not $HooksConfig.PSObject.Properties[$EventName]) { return @($history) }

    $retryCount = 1
    $retryDelay = 1
    $deduplicate = $true
    if ($HooksConfig.PSObject.Properties["runtime"] -and $HooksConfig.runtime) {
        if ($HooksConfig.runtime.PSObject.Properties["retryCount"]) { $retryCount = [Math]::Max([int]$HooksConfig.runtime.retryCount, 1) }
        if ($HooksConfig.runtime.PSObject.Properties["retryDelaySeconds"]) { $retryDelay = [Math]::Max([int]$HooksConfig.runtime.retryDelaySeconds, 0) }
        if ($HooksConfig.runtime.PSObject.Properties["deduplicate"]) { $deduplicate = [bool]$HooksConfig.runtime.deduplicate }
    }

    $queueCtx = Get-HookQueueContext -LogsDir $LogsDir
    [void](Invoke-HookAckResync -LogsDir $LogsDir -MaxItems 10)
    $ledgerPath = Join-Path $queueCtx.hookDir "Hook_Dedup_Ledger.json"
    $ledger = @()
    if (Test-Path $ledgerPath) {
        try { $ledger = @(Get-Content -Path $ledgerPath -Raw -Encoding utf8 | ConvertFrom-Json) } catch { $ledger = @() }
    }

    $order = 0
    $pendingItems = @()

    # フェーズ1: 全アクションをキューに一括投入（重複チェック・dedup込み）
    foreach ($action in @($HooksConfig.$EventName)) {
        $order++
        $type = if ($action.PSObject.Properties["type"]) { "$($action.type)".ToLowerInvariant() } else { "log" }
        $name = if ($action.PSObject.Properties["name"]) { "$($action.name)" } else { "{0}-{1}-{2}" -f $EventName, $type, $order }
        $payload = New-StandardHookPayload -EventName $EventName -RunId $RunId -Context ([PSCustomObject]@{ payload = $action; context = $Context }) -TransactionId $TransactionId

        $dedupeKey = Get-StableSha256Hash -Text ("{0}|{1}|{2}|{3}|{4}" -f $EventName, $name, $type, $payload.transactionId, ($payload | ConvertTo-Json -Depth 10 -Compress))
        $alreadyDone = @($ledger | Where-Object { $_.PSObject.Properties["dedupeKey"] -and $_.PSObject.Properties["status"] -and $_.dedupeKey -eq $dedupeKey -and $_.status -eq "Success" }).Count -gt 0
        if ($deduplicate -and $alreadyDone) {
            $history += [PSCustomObject]@{
                schemaVersion = "1.1"
                event = $EventName
                action = $name
                type = $type
                startedAt = (Get-Date).ToString("s")
                finishedAt = (Get-Date).ToString("s")
                status = "SkippedDedupe"
                detail = "deduplicated"
                attempts = 0
                order = $order
                hookPayload = $payload
                attemptHistory = @()
            }
            continue
        }

        $queued = New-HookQueueItem -ContextInfo $queueCtx -Payload $payload -EventName $EventName -ActionName $name -Type $type
        $pendingItems += [PSCustomObject]@{
            order       = $order
            name        = $name
            type        = $type
            dedupeKey   = $dedupeKey
            payload     = $payload
            sequence    = $queued.item.sequence
        }
    }

    # フェーズ2: バッチ投入したアイテムを1回のキュー処理でまとめて実行
    if ($pendingItems.Count -gt 0) {
        $processed = Process-HookQueue -ContextInfo $queueCtx -HooksConfig $HooksConfig -RetryCount $retryCount -RetryDelaySeconds $retryDelay
        foreach ($pi in $pendingItems) {
            $match = @($processed | Where-Object { $_.order -eq $pi.sequence })
            if (@($match).Count -gt 0) {
                $history += $match[0]
                $ledger += [PSCustomObject]@{
                    dedupeKey = $pi.dedupeKey
                    event = $EventName
                    action = $pi.name
                    status = $match[0].status
                    transactionId = $pi.payload.transactionId
                    runId = $RunId
                    updatedAt = (Get-Date).ToString("s")
                }
            } else {
                # Process-HookQueue が途中ブレーク（順序保証）でこのアイテムまで到達しなかった
                $history += [PSCustomObject]@{
                    schemaVersion = "1.1"
                    event = $EventName
                    action = $pi.name
                    type = $pi.type
                    startedAt = (Get-Date).ToString("s")
                    finishedAt = (Get-Date).ToString("s")
                    status = "Pending"
                    detail = "queued but not processed — order guarantee blocked by earlier failure"
                    attempts = 0
                    order = $pi.order
                    hookPayload = $pi.payload
                    attemptHistory = @()
                }
            }
        }
    }

    $ledger = @($ledger | Select-Object -Last 10000)
    $ledger | ConvertTo-Json -Depth 8 | Set-ContentWithRetry -Path $ledgerPath -Encoding $script:_enc
    if ($RunId) {
        $history | ConvertTo-Json -Depth 12 | Set-Content -Path (Join-Path $queueCtx.hookDir ("Hook_{0}_{1}_{2}.json" -f $EventName, $RunId, (Get-Date -Format "yyyyMMddHHmmss"))) -Encoding $script:_enc
    }
    $historyArray = @($history)
    Export-HookSiemLines -Entries $historyArray -HooksConfig $HooksConfig -LogsDir $LogsDir -RunId $RunId
    return $historyArray
}

function Update-AgentQualityMetrics {
    [CmdletBinding()]
    param(
        [string]$RunId,
        [string]$ReportsDir,
        [object[]]$NodeResults
    )

    $metricsDir = Join-Path (Join-Path $ReportsDir "agent-teams") "metrics"
    if (-not (Test-Path $metricsDir)) { New-Item -ItemType Directory -Path $metricsDir -Force | Out-Null }

    $historyPath = Join-Path $metricsDir "AgentMetrics_History.json"
    $summaryPath = Join-Path $metricsDir "AgentMetrics_Summary_latest.json"
    $history = @()
    if (Test-Path $historyPath) {
        try { $history = @(Get-Content -Path $historyPath -Raw -Encoding utf8 | ConvertFrom-Json) } catch { $history = @() }
    }

    $newRows = @($NodeResults | Where-Object { $_.agentId -and $_.role -in @("collector", "analyzer") } | ForEach-Object {
        [PSCustomObject]@{
            timestamp = (Get-Date).ToString("s")
            runId = $RunId
            agentId = $_.agentId
            status = $_.status
            durationMs = $_.durationMs
            retryCount = [Math]::Max([int]$_.attempts - 1, 0)
        }
    })
    $history = @($history + $newRows | Select-Object -Last 5000)
    $history | ConvertTo-Json -Depth 8 | Set-Content -Path $historyPath -Encoding $script:_enc

    $agents = @()
    $validHistory = @($history | Where-Object { $_ -and $_.PSObject.Properties["agentId"] })
    foreach ($grp in @($validHistory | Group-Object agentId)) {
        $rows = @($grp.Group)
        $count = @($rows).Count
        $ok = @($rows | Where-Object { $_.PSObject.Properties["status"] -and $_.status -eq "Success" }).Count
        $durVals = @($rows | ForEach-Object { if ($_.PSObject.Properties["durationMs"]) { [double]$_.durationMs } })
        $retryVals = @($rows | ForEach-Object { if ($_.PSObject.Properties["retryCount"]) { [double]$_.retryCount } })
        $avgDur = if (@($durVals).Count -gt 0) { [Math]::Round((($durVals | Measure-Object -Average).Average), 2) } else { 0 }
        $avgRetry = if (@($retryVals).Count -gt 0) { [Math]::Round((($retryVals | Measure-Object -Average).Average), 2) } else { 0 }
        $agents += [PSCustomObject]@{
            agentId = $grp.Name
            runCount = $count
            successRate = if ($count -gt 0) { [Math]::Round(($ok * 100.0) / $count, 2) } else { 0 }
            avgDurationMs = $avgDur
            avgRetryCount = $avgRetry
            updatedAt = (Get-Date).ToString("s")
        }
    }

    $summary = [PSCustomObject]@{
        schemaVersion = "1.0"
        generatedAt = (Get-Date).ToString("s")
        runId = $RunId
        agents = @($agents | Sort-Object agentId)
        historyPath = $historyPath
    }
    $summary | ConvertTo-Json -Depth 8 | Set-Content -Path $summaryPath -Encoding $script:_enc

    return [PSCustomObject]@{
        historyPath = $historyPath
        summaryPath = $summaryPath
        summary = $summary
    }
}

function Invoke-McpProviders {
    [CmdletBinding()]
    param(
        [pscustomobject]$McpProviders,
        [pscustomobject]$Payload,
        [string]$RunId = "",
        [string]$ReportsDir = "",
        [pscustomobject]$HooksConfig,
        [string]$LogsDir = "",
        [string]$TransactionId = ""
    )

    $results = @()
    if (-not $McpProviders) { return @($results) }

    $tx = if ($TransactionId) { $TransactionId } elseif ($RunId) { $RunId } else { [guid]::NewGuid().ToString("N") }
    if (-not $ReportsDir) { $ReportsDir = "." }
    $mcpDir = Join-Path $ReportsDir "mcp"
    if (-not (Test-Path $mcpDir)) { New-Item -ItemType Directory -Path $mcpDir -Force | Out-Null }
    $dlqDir = Join-Path $mcpDir "dead-letter"
    if (-not (Test-Path $dlqDir)) { New-Item -ItemType Directory -Path $dlqDir -Force | Out-Null }

    $ledgerPath = Join-Path $mcpDir "MCP_Transactions.json"
    $ledger = @()
    if (Test-Path $ledgerPath) {
        try { $ledger = @(Get-Content -Path $ledgerPath -Raw -Encoding utf8 | ConvertFrom-Json) } catch { $ledger = @() }
    }

    $payloadHash = Get-StableSha256Hash -Text ($Payload | ConvertTo-Json -Depth 12 -Compress)

    $ledgerSafe = @($ledger | Where-Object { $_ -and $_.PSObject.Properties["operationId"] -and $_.PSObject.Properties["status"] })
    foreach ($provider in @($McpProviders)) {
        $name = if ($provider.name) { "$($provider.name)" } else { "mcp" }
        $type = if ($provider.type) { "$($provider.type)".ToLowerInvariant() } else { "file" }
        $enabled = if ($provider.PSObject.Properties["enabled"]) { [bool]$provider.enabled } else { $true }
        $retryCount = if ($provider.PSObject.Properties["retryCount"] -and $provider.retryCount) { [Math]::Max([int]$provider.retryCount, 1) } else { 1 }
        $retryDelaySeconds = if ($provider.PSObject.Properties["retryDelaySeconds"] -and $provider.retryDelaySeconds) { [Math]::Max([int]$provider.retryDelaySeconds, 0) } else { 1 }
        $operationId = Get-StableSha256Hash -Text ("{0}|{1}|{2}|{3}|{4}" -f $tx, $name, $type, $RunId, $payloadHash)

        $alreadyOk = @($ledgerSafe | Where-Object { $_.operationId -eq $operationId -and $_.status -eq "Success" }).Count -gt 0
        if (-not $enabled) {
            $results += [PSCustomObject]@{
                schemaVersion = "1.1"
                transactionId = $tx
                operationId = $operationId
                name = $name
                type = $type
                status = "Skipped"
                attempts = 0
                retryHistory = @()
                message = "disabled"
                deadLetterPath = $null
                rollbackHint = $null
                updatedAt = (Get-Date).ToString("s")
            }
            continue
        }
        if ($alreadyOk) {
            $results += [PSCustomObject]@{
                schemaVersion = "1.1"
                transactionId = $tx
                operationId = $operationId
                name = $name
                type = $type
                status = "SkippedIdempotent"
                attempts = 0
                retryHistory = @()
                message = "already-succeeded"
                deadLetterPath = $null
                rollbackHint = $null
                updatedAt = (Get-Date).ToString("s")
            }
            continue
        }

        $ok = $false
        $attempts = 0
        $message = ""
        $deadLetterPath = $null
        $retryHistory = @()
        $rollbackHint = $null
        $response = $null
        $referenceId = $null
        $providerSnapshot = [PSCustomObject]@{
            name = $name
            type = $type
            url = if ($provider.PSObject.Properties["url"]) { "$($provider.url)" } else { $null }
            webhookUrl = if ($provider.PSObject.Properties["webhookUrl"]) { "$($provider.webhookUrl)" } else { $null }
            instanceUrl = if ($provider.PSObject.Properties["instanceUrl"]) { "$($provider.instanceUrl)" } else { $null }
            table = if ($provider.PSObject.Properties["table"]) { "$($provider.table)" } else { $null }
            token = if ($provider.PSObject.Properties["token"]) { "$($provider.token)" } else { $null }
            userEmail = if ($provider.PSObject.Properties["userEmail"]) { "$($provider.userEmail)" } else { $null }
            apiToken = if ($provider.PSObject.Properties["apiToken"]) { "$($provider.apiToken)" } else { $null }
        }


        # Extract notification context from payload for rich messages
        $notifScore  = 0
        $notifEval   = "N/A"
        $notifTopRec = ""
        $notifHost   = if ($env:COMPUTERNAME) { "$env:COMPUTERNAME" } else { "unknown" }
        if ($Payload -and $Payload.PSObject.Properties["reporter"] -and $Payload.reporter) {
            if ($Payload.reporter.PSObject.Properties["healthScore"] -and $null -ne $Payload.reporter.healthScore) {
                try { $notifScore = [int]$Payload.reporter.healthScore } catch { $notifScore = 0 }
            }
            if ($Payload.reporter.PSObject.Properties["aiEvaluation"] -and $Payload.reporter.aiEvaluation) {
                $notifEval = "$($Payload.reporter.aiEvaluation)"
            }
        }
        if ($Payload -and $Payload.PSObject.Properties["remediator"] -and $Payload.remediator) {
            if ($Payload.remediator.PSObject.Properties["recommendedActions"] -and @($Payload.remediator.recommendedActions).Count -gt 0) {
                $notifTopRec = "$(@($Payload.remediator.recommendedActions)[0])"
            }
        }

        # Score threshold: skip ServiceNow/Jira ticket when PC score is healthy
        if ($type -in @("servicenow", "jira")) {
            $scoreThreshold = if ($provider.PSObject.Properties["scoreThreshold"]) { [int]$provider.scoreThreshold } else { 70 }
            if ($notifScore -gt 0 -and $notifScore -ge $scoreThreshold) {
                $results += [PSCustomObject]@{
                    schemaVersion = "1.1"; transactionId = $tx; operationId = $operationId; name = $name; type = $type
                    status = "SkippedThreshold"; attempts = 0; retryHistory = @()
                    message = "score $notifScore >= threshold $scoreThreshold"
                    referenceId = $null; providerSnapshot = $providerSnapshot; deadLetterPath = $null; rollbackHint = $null
                    updatedAt = (Get-Date).ToString("s")
                }
                $ledger += $results[-1]
                continue
            }
        }

        while ($attempts -lt $retryCount -and -not $ok) {
            $attempts++
            try {
                if ($type -eq "file") {
                    $outFile = Join-Path $mcpDir ("MCP_{0}_{1}.json" -f $name, $RunId)
                    $Payload | ConvertTo-Json -Depth 12 | Set-Content -Path $outFile -Encoding $script:_enc
                    $message = "written:$outFile"
                    $rollbackHint = "Delete generated file."
                } elseif ($type -eq "webhook") {
                    if (-not $provider.url) { throw "webhook url is required." }
                    $method = if ($provider.method) { "$($provider.method)" } else { "POST" }
                    $response = Invoke-RestMethod -Uri "$($provider.url)" -Method $method -Body ($Payload | ConvertTo-Json -Depth 12) -ContentType "application/json" -TimeoutSec 30
                    $message = "posted:$($provider.url)"
                    $rollbackHint = "Send compensating webhook."
                } elseif ($type -eq "slack") {
                    if (-not $provider.webhookUrl) { throw "slack webhookUrl is required." }
                    $slColor = if ($notifScore -ge 80) { "#36a64f" } elseif ($notifScore -ge 50) { "#ffcc00" } else { "#ff0000" }
                    $slBody  = @{
                        text        = "PC Health Report - Score: $notifScore/100"
                        attachments = @(@{
                            color  = $slColor
                            fields = @(
                                @{ title = "ホスト名";       value = $notifHost;   short = $true  }
                                @{ title = "評価";           value = $notifEval;   short = $true  }
                                @{ title = "推奨アクション"; value = $notifTopRec; short = $false }
                                @{ title = "RunID";          value = $RunId;       short = $true  }
                            )
                        })
                    } | ConvertTo-Json -Depth 8 -Compress
                    $slBytes = [System.Text.Encoding]::UTF8.GetBytes($slBody)
                    if ($PSVersionTable.PSVersion.Major -ge 7) {
                        $response = Invoke-RestMethod -Uri "$($provider.webhookUrl)" -Method Post -Body $slBytes -ContentType "application/json" -TimeoutSec 30
                    } else {
                        Invoke-WebRequest -Uri "$($provider.webhookUrl)" -Method Post -Body $slBytes -ContentType "application/json" -TimeoutSec 30 -UseBasicParsing | Out-Null
                    }
                    $message = "posted:slack:score=$notifScore"
                    $rollbackHint = "Post correction message."
                } elseif ($type -eq "teams") {
                    if (-not $provider.webhookUrl) { throw "teams webhookUrl is required." }
                    $tmColor = if ($notifScore -ge 80) { "00B050" } elseif ($notifScore -ge 50) { "FFC000" } else { "FF0000" }
                    $tmBody  = @{
                        "@type"    = "MessageCard"
                        "@context" = "http://schema.org/extensions"
                        themeColor = $tmColor
                        summary    = "PC Health Report"
                        sections   = @(@{
                            activityTitle    = "PC Health Report - Score: $notifScore/100"
                            activitySubtitle = "ホスト名: $notifHost"
                            facts            = @(
                                @{ name = "評価";           value = $notifEval   }
                                @{ name = "推奨アクション"; value = $notifTopRec }
                                @{ name = "RunID";          value = $RunId       }
                            )
                        })
                    } | ConvertTo-Json -Depth 8 -Compress
                    $tmBytes = [System.Text.Encoding]::UTF8.GetBytes($tmBody)
                    if ($PSVersionTable.PSVersion.Major -ge 7) {
                        $response = Invoke-RestMethod -Uri "$($provider.webhookUrl)" -Method Post -Body $tmBytes -ContentType "application/json" -TimeoutSec 30
                    } else {
                        Invoke-WebRequest -Uri "$($provider.webhookUrl)" -Method Post -Body $tmBytes -ContentType "application/json" -TimeoutSec 30 -UseBasicParsing | Out-Null
                    }
                    $message = "posted:teams:score=$notifScore"
                    $rollbackHint = "Post correction message."
                } elseif ($type -eq "servicenow") {
                    if (-not $provider.instanceUrl -or -not $provider.table) { throw "servicenow instanceUrl and table are required." }
                    $url = ("{0}/api/now/table/{1}" -f "$($provider.instanceUrl)".TrimEnd("/"), "$($provider.table)")
                    $body = @{
                        short_description = ("PC_Optimizer Run {0}" -f $RunId)
                        description = ($Payload | ConvertTo-Json -Depth 8)
                    } | ConvertTo-Json
                    $headers = @{}
                    if ($provider.token) { $headers["Authorization"] = "Bearer $($provider.token)" }
                    $response = Invoke-RestMethod -Uri $url -Method Post -Headers $headers -Body $body -ContentType "application/json" -TimeoutSec 30
                    $message = "posted:servicenow"
                    $rollbackHint = "Close or cancel created incident."
                    if ($response -and $response.result -and $response.result.sys_id) { $referenceId = "$($response.result.sys_id)" }
                } elseif ($type -eq "jira") {
                    if (-not $provider.url -or -not $provider.projectKey) { throw "jira url and projectKey are required." }
                    $url = ("{0}/rest/api/2/issue" -f "$($provider.url)".TrimEnd("/"))
                    $issueType = if ($provider.issueType) { "$($provider.issueType)" } else { "Task" }
                    $body = @{
                        fields = @{
                            project = @{ key = $provider.projectKey }
                            summary = ("PC_Optimizer Run {0}" -f $RunId)
                            description = ($Payload | ConvertTo-Json -Depth 8)
                            issuetype = @{ name = $issueType }
                        }
                    } | ConvertTo-Json -Depth 10
                    $headers = @{}
                    if ($provider.userEmail -and $provider.apiToken) {
                        $pair = "{0}:{1}" -f $provider.userEmail, $provider.apiToken
                        $headers["Authorization"] = "Basic $([Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($pair)))"
                    }
                    $response = Invoke-RestMethod -Uri $url -Method Post -Headers $headers -Body $body -ContentType "application/json" -TimeoutSec 30
                    $message = "posted:jira"
                    $rollbackHint = "Transition ticket to canceled."
                    if ($response -and $response.key) { $referenceId = "$($response.key)" }
                } else {
                    throw "provider-type-not-implemented:$type"
                }

                $ok = $true
                $retryHistory += [PSCustomObject]@{ attempt = $attempts; status = "Success"; detail = $message; at = (Get-Date).ToString("s") }
            } catch {
                $message = "$($_.Exception.Message)"
                $retryHistory += [PSCustomObject]@{ attempt = $attempts; status = "Failed"; detail = $message; at = (Get-Date).ToString("s") }
                if ($attempts -lt $retryCount -and $retryDelaySeconds -gt 0) {
                    Start-Sleep -Seconds ([int]([Math]::Pow(2, [Math]::Max($attempts - 1, 0)) * $retryDelaySeconds))
                }
            }
        }

        if (-not $ok) {
            $deadLetterPath = Join-Path $dlqDir ("DLQ_{0}_{1}_{2}.json" -f $name, $RunId, (Get-Date -Format "yyyyMMddHHmmssfff"))
            [PSCustomObject]@{
                transactionId = $tx
                operationId = $operationId
                runId = $RunId
                provider = $provider
                payload = $Payload
                retryHistory = @($retryHistory)
                error = $message
                rollbackHint = if ($rollbackHint) { $rollbackHint } else { "Manual rollback required." }
            } | ConvertTo-Json -Depth 12 | Set-Content -Path $deadLetterPath -Encoding $script:_enc
        }

        $entry = [PSCustomObject]@{
            schemaVersion = "1.1"
            transactionId = $tx
            operationId = $operationId
            name = $name
            type = $type
            status = if ($ok) { "Success" } else { "Failed" }
            attempts = $attempts
            retryHistory = @($retryHistory)
            message = $message
            referenceId = $referenceId
            providerSnapshot = $providerSnapshot
            deadLetterPath = $deadLetterPath
            rollbackHint = $rollbackHint
            updatedAt = (Get-Date).ToString("s")
        }
        $results += $entry
        $ledger += $entry
    }

    $ledger = @($ledger | Select-Object -Last 5000)
    $ledger | ConvertTo-Json -Depth 15 | Set-Content -Path $ledgerPath -Encoding $script:_enc
    return @($results)
}

function Invoke-McpProvidersParallel {
    <#
    .SYNOPSIS
        複数のMCPプロバイダーをRunspacePoolで並列実行します。
    .DESCRIPTION
        file/webhook/slack/teams/servicenow/jira の各プロバイダーを独立したRunspaceで
        同時実行し、総実行時間を短縮します。1プロバイダーの場合は逐次実行にフォールバック。
        結果はプロバイダーの定義順にソートして返します（ledger + DLQ への書き込みは
        全runspace完了後に一括で実施します）。
    .PARAMETER McpProviders
        MCPプロバイダー設定配列
    .PARAMETER Payload
        各プロバイダーに渡すペイロード
    .PARAMETER RunId
        実行識別子
    .PARAMETER ReportsDir
        レポート出力ディレクトリ
    .PARAMETER MaxParallel
        最大並列数（既定: 8）
    .PARAMETER TimeoutSeconds
        各プロバイダーのタイムアウト秒数（既定: 60）
    .OUTPUTS
        MCPプロバイダー実行結果の配列（Invoke-McpProvidersと同形式）
    #>
    [CmdletBinding()]
    param(
        [pscustomobject]$McpProviders,
        [pscustomobject]$Payload,
        [string]$RunId = "",
        [string]$ReportsDir = "",
        [int]$MaxParallel = 8,
        [int]$TimeoutSeconds = 60
    )

    # 有効プロバイダーが1件以下なら通常の逐次実行にフォールバック
    $enabledCount = @(@($McpProviders) | Where-Object {
        $_ -and (-not $_.PSObject.Properties["enabled"] -or [bool]$_.enabled)
    }).Count
    if ($enabledCount -le 1) {
        return @(Invoke-McpProviders -McpProviders $McpProviders -Payload $Payload -RunId $RunId -ReportsDir $ReportsDir)
    }

    $modulePath = $PSCommandPath
    if (-not $modulePath) {
        # フォールバック: モジュールパスを取得できない場合は逐次実行
        Write-Warning "[Invoke-McpProvidersParallel] モジュールパスを取得できません。逐次実行にフォールバックします。"
        return @(Invoke-McpProviders -McpProviders $McpProviders -Payload $Payload -RunId $RunId -ReportsDir $ReportsDir)
    }

    $pool = [RunspaceFactory]::CreateRunspacePool(1, [Math]::Max($MaxParallel, 1))
    $pool.Open()

    $jobs = @()
    $providerList = @($McpProviders)
    $providerIndex = 0

    foreach ($provider in $providerList) {
        $idx = $providerIndex++
        $ps = [powershell]::Create()
        $ps.RunspacePool = $pool
        [void]$ps.AddScript({
            param($modPath, $provArg, $payArg, $runIdArg, $reportsDirArg, $idxArg)
            Import-Module $modPath -Force | Out-Null
            $single = @($provArg)
            $r = Invoke-McpProviders -McpProviders $single -Payload $payArg -RunId $runIdArg -ReportsDir $reportsDirArg
            return ([PSCustomObject]@{
                index   = $idxArg
                results = $r
            } | ConvertTo-Json -Depth 16 -Compress)
        })
        [void]$ps.AddArgument($modulePath)
        [void]$ps.AddArgument($provider)
        [void]$ps.AddArgument($Payload)
        [void]$ps.AddArgument($RunId)
        [void]$ps.AddArgument($ReportsDir)
        [void]$ps.AddArgument($idx)
        $jobs += [PSCustomObject]@{
            index     = $idx
            provider  = $provider
            ps        = $ps
            handle    = $ps.BeginInvoke()
            startedAt = Get-Date
        }
    }

    # 全Runspace の完了を待機して結果収集
    $rawResults = @{}
    foreach ($j in $jobs) {
        try {
            $waitOk = $j.handle.AsyncWaitHandle.WaitOne([TimeSpan]::FromSeconds($TimeoutSeconds))
            if (-not $waitOk) {
                try { $j.ps.Stop() } catch {}
                $pName = if ($j.provider -and $j.provider.PSObject.Properties["name"]) { "$($j.provider.name)" } else { "mcp-$($j.index)" }
                $rawResults[$j.index] = [PSCustomObject]@{
                    schemaVersion = "1.1"
                    name = $pName
                    type = if ($j.provider -and $j.provider.PSObject.Properties["type"]) { "$($j.provider.type)" } else { "unknown" }
                    status = "Failed"
                    message = "parallel-timeout:${TimeoutSeconds}s"
                    attempts = 1
                    retryHistory = @([PSCustomObject]@{ attempt = 1; status = "Failed"; detail = "timeout"; at = (Get-Date).ToString("s") })
                    deadLetterPath = $null
                    rollbackHint = "Parallel timeout — retry sequentially."
                    updatedAt = (Get-Date).ToString("s")
                }
                continue
            }
            $raw = @($j.ps.EndInvoke($j.handle)) -join ""
            if (-not [string]::IsNullOrWhiteSpace($raw)) {
                $parsed = $raw | ConvertFrom-Json
                if ($parsed -and $parsed.PSObject.Properties["results"]) {
                    $rawResults[$j.index] = @($parsed.results)[0]
                } else {
                    $pName = if ($j.provider -and $j.provider.PSObject.Properties["name"]) { "$($j.provider.name)" } else { "mcp-$($j.index)" }
                    $rawResults[$j.index] = [PSCustomObject]@{
                        schemaVersion = "1.1"; name = $pName
                        type = if ($j.provider -and $j.provider.PSObject.Properties["type"]) { "$($j.provider.type)" } else { "unknown" }
                        status = "Failed"; message = "parallel-empty-results"; attempts = 1
                        retryHistory = @([PSCustomObject]@{ attempt = 1; status = "Failed"; detail = "no results property in runspace output"; at = (Get-Date).ToString("s") })
                        deadLetterPath = $null; rollbackHint = $null; updatedAt = (Get-Date).ToString("s")
                    }
                }
            } else {
                $pName = if ($j.provider -and $j.provider.PSObject.Properties["name"]) { "$($j.provider.name)" } else { "mcp-$($j.index)" }
                $rawResults[$j.index] = [PSCustomObject]@{
                    schemaVersion = "1.1"; name = $pName
                    type = if ($j.provider -and $j.provider.PSObject.Properties["type"]) { "$($j.provider.type)" } else { "unknown" }
                    status = "Failed"; message = "parallel-no-output"; attempts = 1
                    retryHistory = @([PSCustomObject]@{ attempt = 1; status = "Failed"; detail = "runspace produced no output"; at = (Get-Date).ToString("s") })
                    deadLetterPath = $null; rollbackHint = $null; updatedAt = (Get-Date).ToString("s")
                }
            }
        } catch {
            $pName = if ($j.provider -and $j.provider.PSObject.Properties["name"]) { "$($j.provider.name)" } else { "mcp-$($j.index)" }
            $rawResults[$j.index] = [PSCustomObject]@{
                schemaVersion = "1.1"
                name = $pName
                type = if ($j.provider -and $j.provider.PSObject.Properties["type"]) { "$($j.provider.type)" } else { "unknown" }
                status = "Failed"
                message = "parallel-exception:$($_.Exception.Message)"
                attempts = 1
                retryHistory = @([PSCustomObject]@{ attempt = 1; status = "Failed"; detail = "$($_.Exception.Message)"; at = (Get-Date).ToString("s") })
                deadLetterPath = $null
                rollbackHint = $null
                updatedAt = (Get-Date).ToString("s")
            }
        } finally {
            $j.ps.Dispose()
        }
    }

    $pool.Close()
    $pool.Dispose()

    # 元の定義順で結果を返す
    $orderedResults = @()
    for ($i = 0; $i -lt $providerList.Count; $i++) {
        if ($rawResults.ContainsKey($i)) {
            $orderedResults += $rawResults[$i]
        }
    }
    return @($orderedResults)
}

function Invoke-AgentTeamsOrchestration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$RunId,
        [Parameter(Mandatory)][string]$ReportsDir,
        [pscustomobject]$ModuleSnapshot,
        [pscustomobject]$HealthScore,
        [pscustomobject]$AIDiagnosis,
        [pscustomobject]$HooksConfig,
        [pscustomobject]$McpProviders,
        [pscustomobject]$AgentTeamsConfig,
        [string]$LogsDir = "",
        [string]$RunMode = "repair"
    )

    if (-not (Test-Path $ReportsDir)) { New-Item -ItemType Directory -Path $ReportsDir -Force | Out-Null }
    $agentDir = Join-Path $ReportsDir "agent-teams"
    if (-not (Test-Path $agentDir)) { New-Item -ItemType Directory -Path $agentDir -Force | Out-Null }

    $pluginsRoot = Join-Path $PSScriptRoot "agents"
    $preset = Get-AgentTeamsProfilePreset -RunMode $RunMode
    $globalParallel = if ($AgentTeamsConfig -and $AgentTeamsConfig.PSObject.Properties["maxParallel"]) { [Math]::Max([int]$AgentTeamsConfig.maxParallel, 1) } else { [Math]::Max([int]$preset.maxParallel, 1) }
    $nodes = if ($AgentTeamsConfig -and $AgentTeamsConfig.PSObject.Properties["dagNodes"] -and @($AgentTeamsConfig.dagNodes).Count -gt 0) { @($AgentTeamsConfig.dagNodes) } else { @(New-DefaultAgentDagNodes -RunMode $RunMode) }
    $transactionId = if ($RunId) { $RunId } else { [guid]::NewGuid().ToString("N") }

    $plan = [PSCustomObject]@{
        schemaVersion = "1.1"
        runId = $RunId
        transactionId = $transactionId
        generatedAt = (Get-Date).ToString("s")
        planner = [PSCustomObject]@{
            version = "2026-03-05.4"
            profile = $preset.name
            runMode = $RunMode
            maxParallel = $globalParallel
            enabledHookEvents = @($preset.enabledHookEvents)
            enabledMcpProviders = @($preset.enabledMcpProviders)
        }
        dag = @($nodes)
    }
    $planPath = Join-Path $agentDir ("AgentTeams_Plan_{0}.json" -f $RunId)
    $plan | ConvertTo-Json -Depth 16 | Set-Content -Path $planPath -Encoding $script:_enc

    $effectiveMcp = @()
    foreach ($m in @($McpProviders)) {
        $clone = [PSCustomObject]@{}
        foreach ($p in $m.PSObject.Properties) {
            $clone | Add-Member -NotePropertyName $p.Name -NotePropertyValue $p.Value -Force
        }
        $allowedByPreset = @($preset.enabledMcpProviders | Where-Object { $_ -eq "$($clone.name)" }).Count -gt 0
        $enabledByConfig = if ($clone.PSObject.Properties["enabled"]) { [bool]$clone.enabled } else { $true }
        $clone | Add-Member -NotePropertyName "enabled" -NotePropertyValue ([bool]($allowedByPreset -and $enabledByConfig)) -Force
        $effectiveMcp += $clone
    }

    $pending = @{}
    foreach ($n in $nodes) { $pending["$($n.id)"] = $n }

    $collectorByAgent = @{}
    $analyzerByAgent = @{}
    $nodeResults = New-Object System.Collections.Generic.List[object]
    $hookHistory = @()
    $levelTimeline = @()
    $conversationLog = New-Object System.Collections.Generic.List[object]
    $level = 0

    # 会話ログ用ヘッダー表示
    $sep = "=" * 66
    Write-Host ""
    Write-Host $sep -ForegroundColor Cyan
    Write-Host ("  >> Agent Teams オーケストレーション開始 <<") -ForegroundColor Cyan
    Write-Host ("  RunId: $RunId  |  Profile: $($preset.name)  |  MaxParallel: $globalParallel") -ForegroundColor DarkCyan
    Write-Host ("  DAGノード数: $(@($nodes).Count)  |  RunMode: $RunMode") -ForegroundColor DarkCyan
    Write-Host $sep -ForegroundColor Cyan

    while ($pending.Count -gt 0 -and $level -lt 100) {
        $ready = New-Object System.Collections.Generic.List[object]
        foreach ($node in $pending.Values) {
            $deps = @($node.dependsOn)
            $dependencyOk = $true
            foreach ($depId in $deps) {
                $depResult = @($nodeResults | Where-Object { $_.nodeId -eq "$depId" } | Select-Object -Last 1)
                if (-not $depResult) { $dependencyOk = $false; break }
                if ($depResult[0].status -ne "Success" -and -not $depResult[0].continueOnFail) { $dependencyOk = $false; break }
            }
            if ($dependencyOk) { [void]$ready.Add($node) }
        }

        if ($ready.Count -eq 0) { break }

        $sortedReady = @($ready | Sort-Object @{ Expression = { (Get-NodePolicy -Node $_ -GlobalMaxParallel $globalParallel).priority }; Descending = $true })
        $levelTimeline += [PSCustomObject]@{
            level = $level
            nodes = @($sortedReady | ForEach-Object { "$($_.id)" })
        }

        while (@($sortedReady).Count -gt 0) {
            $batch = @($sortedReady | Select-Object -First $globalParallel)
            $sortedReady = @($sortedReady | Select-Object -Skip $globalParallel)
            foreach ($node in $batch) {
                $policy = Get-NodePolicy -Node $node -GlobalMaxParallel $globalParallel
                $contextStart = [PSCustomObject]@{
                    runId = $RunId
                    transactionId = $transactionId
                    nodeId = $node.id
                    agentId = $node.agentId
                    stage = "start"
                    role = $node.role
                    priority = $policy.priority
                }
                $hookHistory += @(Invoke-AgentHookEvent -EventName "pre_task" -Context $contextStart -HooksConfig $HooksConfig -RunId $RunId -LogsDir $LogsDir -TransactionId $transactionId)
            }

            $modulePathForRunspace = $MyInvocation.MyCommand.Module.Path
            $parallelResults = Invoke-AgentNodeBatchRunspace `
                -Nodes $batch `
                -MaxParallel $globalParallel `
                -ModulePath $modulePathForRunspace `
                -RunId $RunId `
                -ModuleSnapshot $ModuleSnapshot `
                -CollectorByAgent $collectorByAgent `
                -AnalyzerByAgent $analyzerByAgent `
                -HealthScore $HealthScore `
                -AIDiagnosis $AIDiagnosis `
                -PluginsRoot $pluginsRoot

            foreach ($pr in @($parallelResults)) {
                $node = $pr.node
                $result = $pr.result
                if (-not $result -or -not $result.PSObject.Properties["status"]) {
                    $result = [PSCustomObject]@{
                        status = "Failed"
                        risk = "Unknown"
                        payload = $null
                        message = "invalid-runspace-result-shape"
                    }
                }
                $policy = Get-NodePolicy -Node $node -GlobalMaxParallel $globalParallel
                $started = if ($pr.startedAt) { [datetime]$pr.startedAt } else { Get-Date }
                $finished = if ($pr.finishedAt) { [datetime]$pr.finishedAt } else { Get-Date }

                $row = [PSCustomObject]@{
                    runId = $RunId
                    level = $level
                    nodeId = "$($node.id)"
                    role = "$($node.role)"
                    agentId = "$($node.agentId)"
                    status = "$($result.status)"
                    startedAt = $started.ToString("s")
                    finishedAt = $finished.ToString("s")
                    durationMs = [int][Math]::Round(($finished - $started).TotalMilliseconds, 0)
                    attempts = 1
                    timeoutSeconds = $policy.timeoutSeconds
                    continueOnFail = $policy.continueOnFail
                    priority = $policy.priority
                    maxParallel = $policy.maxParallel
                    parallelGroup = $policy.parallelGroup
                    jitterMs = $policy.jitterMs
                    risk = "$($result.risk)"
                    message = "$($result.message)"
                    payload = $result.payload
                }
                [void]$nodeResults.Add($row)
                if ($row.role -eq "collector" -and $row.status -eq "Success") { $collectorByAgent[$row.agentId] = $row }
                if ($row.role -eq "analyzer" -and $row.status -eq "Success") { $analyzerByAgent[$row.agentId] = $row }

                $eventName = if ($row.status -eq "Success") { "post_task" } else { "on_error" }
                $contextFinish = [PSCustomObject]@{
                    runId = $RunId
                    transactionId = $transactionId
                    nodeId = $row.nodeId
                    agentId = $row.agentId
                    stage = if ($row.status -eq "Success") { "success" } else { "error" }
                    role = $row.role
                    detail = $row.message
                    attempt = $row.attempts
                }
                $hookHistory += @(Invoke-AgentHookEvent -EventName $eventName -Context $contextFinish -HooksConfig $HooksConfig -RunId $RunId -LogsDir $LogsDir -TransactionId $transactionId)
                [void]$pending.Remove("$($node.id)")

                # 会話ログエントリ生成
                $convEntry = Get-AgentConversationMessage -NodeRow $row -AllNodes $nodes
                $convEntry | Add-Member -NotePropertyName "level" -NotePropertyValue $level -Force
                [void]$conversationLog.Add($convEntry)
            }
        }

        $level++

        # レベル完了後にリアルタイム会話表示
        $levelConvEntries = @($conversationLog | Where-Object {
            $_.PSObject.Properties["level"] -and [int]$_.level -eq ($level - 1)
        })
        if (@($levelConvEntries).Count -gt 0) {
            Show-AgentTeamsConversation -ConversationLog $levelConvEntries -RunId $RunId `
                -LevelTimeline @($levelTimeline) -ShowHeader $false -ShowSummaryLine $false
        }
    }

    if ($pending.Count -gt 0) {
        foreach ($node in $pending.Values) {
            $policy = Get-NodePolicy -Node $node -GlobalMaxParallel $globalParallel
            [void]$nodeResults.Add([PSCustomObject]@{
                runId = $RunId
                level = $level
                nodeId = "$($node.id)"
                role = "$($node.role)"
                agentId = "$($node.agentId)"
                status = "Failed"
                startedAt = (Get-Date).ToString("s")
                finishedAt = (Get-Date).ToString("s")
                durationMs = 0
                attempts = 0
                timeoutSeconds = $policy.timeoutSeconds
                continueOnFail = $policy.continueOnFail
                priority = $policy.priority
                maxParallel = $policy.maxParallel
                parallelGroup = $policy.parallelGroup
                jitterMs = $policy.jitterMs
                risk = "Unknown"
                message = "dag-dependency-cycle-or-unsatisfied"
                payload = $null
            })
        }
    }

    $nodeResultsArray = @($nodeResults | ForEach-Object { $_ })
    $collectorResults = @($nodeResultsArray | Where-Object { $_.role -eq "collector" })
    $analyzerResults = @($nodeResultsArray | Where-Object { $_.role -eq "analyzer" })
    $aggregate = @($analyzerResults | Where-Object { $_.nodeId -eq "analyzer.aggregate" } | Select-Object -Last 1)
    $overall = if ($aggregate -and $aggregate[0].payload) { "$($aggregate[0].payload.overall)" } else { "Good" }
    $quality = Update-AgentQualityMetrics -RunId $RunId -ReportsDir $ReportsDir -NodeResults $nodeResultsArray
    $remediator = @($nodeResultsArray | Where-Object { $_.role -eq "remediator" } | Select-Object -Last 1)
    $reporter = @($nodeResultsArray | Where-Object { $_.role -eq "reporter" } | Select-Object -Last 1)

    # 全会話ログ最終サマリー表示
    $allConvArray = @($conversationLog | ForEach-Object { $_ })
    if (@($allConvArray).Count -gt 0) {
        Show-AgentTeamsConversation -ConversationLog $allConvArray -RunId $RunId `
            -LevelTimeline @($levelTimeline) -ShowHeader $true -ShowSummaryLine $true
    }

    # 会話ログを JSON エクスポート
    $conversationLogPath = Export-AgentConversationLog -ConversationLog $allConvArray -ReportsDir $ReportsDir -RunId $RunId

    $recommendedActions = @("Keep preventive maintenance and monitoring.")
    if ($remediator -and $remediator[0].payload -and $remediator[0].payload.recommendedActions) {
        $recommendedActions = @($remediator[0].payload.recommendedActions)
    }

    $merged = @($collectorResults | ForEach-Object {
        [PSCustomObject]@{
            key = ("{0}:{1}" -f $_.runId, $_.agentId)
            runId = $_.runId
            agentId = $_.agentId
            nodeId = $_.nodeId
            risk = $_.risk
            collectedAt = $_.finishedAt
        }
    })

    $aggHigh = @($collectorResults | Where-Object { $_.risk -eq "High" }).Count
    $aggMedium = @($collectorResults | Where-Object { $_.risk -eq "Medium" }).Count
    if ($aggregate -and $aggregate[0].payload) {
        $aggHigh = $aggregate[0].payload.highRiskCount
        $aggMedium = $aggregate[0].payload.mediumRiskCount
    }

    $reporterHealthScore = $null
    $reporterAiEvaluation = $null
    $reporterGeneratedAt = (Get-Date).ToString("s")
    if ($reporter -and $reporter[0].payload) {
        $reporterHealthScore = $reporter[0].payload.healthScore
        $reporterAiEvaluation = $reporter[0].payload.aiEvaluation
        $reporterGeneratedAt = $reporter[0].payload.generatedAt
    } elseif ($HealthScore) {
        $reporterHealthScore = $HealthScore.Score
        if ($AIDiagnosis) { $reporterAiEvaluation = $AIDiagnosis.Evaluation }
    }

    $qualityAgents = @()
    if ($quality.summary) { $qualityAgents = @($quality.summary.agents) }

    $dagTimeline = @($nodeResultsArray | ForEach-Object {
        [PSCustomObject]@{
            runId = $_.runId
            transactionId = $transactionId
            level = $_.level
            nodeId = $_.nodeId
            role = $_.role
            agentId = $_.agentId
            status = $_.status
            startedAt = $_.startedAt
            finishedAt = $_.finishedAt
            durationMs = $_.durationMs
            attempts = $_.attempts
            timeoutSeconds = $_.timeoutSeconds
            continueOnFail = $_.continueOnFail
            priority = $_.priority
            maxParallel = $_.maxParallel
            parallelGroup = $_.parallelGroup
            jitterMs = $_.jitterMs
        }
    })

    $hookTimeline = @($hookHistory | ForEach-Object {
        [PSCustomObject]@{
            event = $_.event
            action = $_.action
            type = $_.type
            status = $_.status
            startedAt = $_.startedAt
            finishedAt = $_.finishedAt
            runId = $_.hookPayload.runId
            transactionId = $_.hookPayload.transactionId
            agentId = $_.hookPayload.agentId
            nodeId = $_.hookPayload.nodeId
            stage = $_.hookPayload.stage
            attempts = $_.attempts
            order = $_.order
        }
    })

    $summary = [PSCustomObject]@{
        schemaVersion = "1.1"
        runId = $RunId
        transactionId = $transactionId
        startedAt = $plan.generatedAt
        finishedAt = (Get-Date).ToString("s")
        planner = $plan.planner
        dagExecution = [PSCustomObject]@{
            levelCount = $level
            levels = @($levelTimeline)
            totalNodes = @($nodeResultsArray).Count
            successNodes = @($nodeResultsArray | Where-Object { $_.status -eq "Success" }).Count
            failedNodes = @($nodeResultsArray | Where-Object { $_.status -eq "Failed" }).Count
            skippedNodes = @($nodeResultsArray | Where-Object { $_.status -eq "Skipped" }).Count
        }
        collector = [PSCustomObject]@{
            parallelMode = "batched"
            workers = @($collectorResults | Select-Object -ExpandProperty agentId -Unique)
            results = @($collectorResults)
        }
        analyzer = [PSCustomObject]@{
            overall = $overall
            highRiskCount = $aggHigh
            mediumRiskCount = $aggMedium
            results = @($analyzerResults)
        }
        remediator = [PSCustomObject]@{
            recommendedActions = @($recommendedActions)
            autoRemediation = $false
        }
        reporter = [PSCustomObject]@{
            healthScore = $reporterHealthScore
            aiEvaluation = $reporterAiEvaluation
            generatedAt = $reporterGeneratedAt
        }
        mergedByRunAndAgent = @($merged)
        hook = [PSCustomObject]@{
            schemaVersion = "1.1"
            count = @($hookHistory).Count
            entries = @($hookHistory)
        }
        qualityMetrics = [PSCustomObject]@{
            summaryPath = $quality.summaryPath
            historyPath = $quality.historyPath
            agents = @($qualityAgents)
        }
        dagTimeline = @($dagTimeline)
        hookTimeline = @($hookTimeline)
        conversation = [PSCustomObject]@{
            count         = @($allConvArray).Count
            logPath       = $conversationLogPath
            entries       = @($allConvArray)
        }
    }

    $summaryPath = Join-Path $agentDir ("AgentTeams_Summary_{0}.json" -f $RunId)
    $summary | ConvertTo-Json -Depth 24 | Set-Content -Path $summaryPath -Encoding $script:_enc

    $mcpResults = Invoke-McpProviders -McpProviders $effectiveMcp -Payload $summary -RunId $RunId -ReportsDir $ReportsDir -HooksConfig $HooksConfig -LogsDir $LogsDir -TransactionId $transactionId
    $autoRollback = $false
    if ($AgentTeamsConfig -and $AgentTeamsConfig.PSObject.Properties["mcpRuntime"] -and $AgentTeamsConfig.mcpRuntime.PSObject.Properties["autoRollbackOnFailure"]) {
        $autoRollback = [bool]$AgentTeamsConfig.mcpRuntime.autoRollbackOnFailure
    }
    if ($autoRollback -and @($mcpResults | Where-Object { $_.status -eq "Failed" }).Count -gt 0) {
        $rollbackResult = Invoke-McpRollbackExecutor -McpResults $mcpResults -ReportsDir $ReportsDir -RunId $RunId -TransactionId $transactionId
        $summary | Add-Member -NotePropertyName "mcpRollback" -NotePropertyValue $rollbackResult -Force
    }
    $summary | Add-Member -NotePropertyName "mcp" -NotePropertyValue @($mcpResults) -Force
    $summary | ConvertTo-Json -Depth 24 | Set-Content -Path $summaryPath -Encoding $script:_enc

    # Fire on_report hook after all processing is complete
    $onReportCtx = [PSCustomObject]@{
        runId           = $RunId
        transactionId   = $transactionId
        stage           = "report"
        healthScore     = $reporterHealthScore
        aiEvaluation    = $reporterAiEvaluation
        highRiskCount   = $aggHigh
        mediumRiskCount = $aggMedium
    }
    $hookHistory += @(Invoke-AgentHookEvent -EventName "on_report" -Context $onReportCtx -HooksConfig $HooksConfig -RunId $RunId -LogsDir $LogsDir -TransactionId $transactionId)
    $onReportEntries = @($hookHistory | Where-Object { $_.PSObject.Properties["event"] -and $_.event -eq "on_report" })
    if ($onReportEntries.Count -gt 0) {
        $summary | Add-Member -NotePropertyName "onReportHooks" -NotePropertyValue $onReportEntries -Force
        $summary | ConvertTo-Json -Depth 24 | Set-Content -Path $summaryPath -Encoding $script:_enc
    }

    return [PSCustomObject]@{
        planPath         = $planPath
        summaryPath      = $summaryPath
        summary          = $summary
        nodeResults      = @($nodeResultsArray)
        hookHistory      = @($hookHistory)
        mcpResults       = @($mcpResults)
        preset           = $preset.name
        transactionId    = $transactionId
        conversationLog  = $allConvArray
        conversationLogPath = $conversationLogPath
    }
}

# ============================================================
# Agent Teams 会話可視化システム
# ============================================================

function Get-AgentNodeIcon {
    [CmdletBinding()]
    param(
        [string]$Role,
        [string]$AgentId = ""
    )
    switch ($Role.ToLowerInvariant()) {
        "planner"    { return "[PLAN]" }
        "collector"  {
            switch ($AgentId) {
                "SecurityAgent" { return "[SEC] " }
                "NetworkAgent"  { return "[NET] " }
                "UpdateAgent"   { return "[UPD] " }
                default         { return "[COL] " }
            }
        }
        "analyzer"   {
            if ($AgentId -eq "Analyzer" -or $Role -match "aggregate") { return "[AGG] " }
            return "[ANL] "
        }
        "remediator" { return "[REM] " }
        "reporter"   { return "[RPT] " }
        default      { return "[AGT] " }
    }
}

function Get-AgentConversationMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][pscustomobject]$NodeRow,
        [object[]]$AllNodes = @()
    )

    $role   = "$($NodeRow.role)".ToLowerInvariant()
    $nodeId = "$($NodeRow.nodeId)"
    $status = "$($NodeRow.status)"
    $risk   = "$($NodeRow.risk)"
    $dur    = $NodeRow.durationMs

    # ダウンストリームノードの特定
    $downstream = @($AllNodes | Where-Object {
        $_ -and $_.PSObject.Properties["dependsOn"] -and @($_.dependsOn) -contains $nodeId
    } | ForEach-Object { "$($_.id)" })
    $toStr = if ($downstream.Count -gt 0) { " → " + ($downstream -join ", ") } else { "" }

    # ロールごとの日本語メッセージ生成
    $msg = switch ($role) {
        "planner" {
            $cnt = $AllNodes.Count
            "DAG計画完了。${cnt}ノードを並列DAGで実行します。${toStr}"
        }
        "collector" {
            $riskLabel = switch ($risk) {
                "High"    { "高リスク[!]" }
                "Medium"  { "中リスク[~]" }
                default   { "低リスク[OK]" }
            }
            $agent = "$($NodeRow.agentId)"
            switch ($agent) {
                "SecurityAgent" { "セキュリティ診断データ収集完了。リスク評価: ${riskLabel}${toStr}" }
                "NetworkAgent"  { "ネットワーク診断データ収集完了。リスク評価: ${riskLabel}${toStr}" }
                "UpdateAgent"   { "Windows Update診断データ収集完了。リスク評価: ${riskLabel}${toStr}" }
                default         { "データ収集完了 [${agent}]。リスク評価: ${riskLabel}${toStr}" }
            }
        }
        "analyzer" {
            if ($nodeId -eq "analyzer.aggregate") {
                $overall = if ($NodeRow.payload -and $NodeRow.payload.PSObject.Properties["overall"]) {
                    switch ("$($NodeRow.payload.overall)") {
                        "Critical" { "Critical[!!]" }
                        "Warning"  { "Warning[!]" }
                        default    { "Good[OK]" }
                    }
                } else { "N/A" }
                "全エージェント集計完了。総合評価: ${overall}${toStr}"
            } else {
                $riskLabel = switch ($risk) {
                    "High"    { "高リスク検出[!!]" }
                    "Medium"  { "中リスク検出[!]" }
                    default   { "リスクなし[OK]" }
                }
                "分析完了。判定: ${riskLabel}${toStr}"
            }
        }
        "remediator" {
            $acts = @()
            if ($NodeRow.payload -and $NodeRow.payload.PSObject.Properties["recommendedActions"]) {
                $acts = @($NodeRow.payload.recommendedActions | Select-Object -First 2)
            }
            $actStr = if ($acts.Count -gt 0) { $acts[0] } else { "監視継続推奨" }
            "修復アクション決定: ${actStr}${toStr}"
        }
        "reporter" {
            $score = if ($NodeRow.payload -and $NodeRow.payload.PSObject.Properties["healthScore"]) {
                "スコア=$($NodeRow.payload.healthScore)/100"
            } else { "" }
            "レポート生成完了。${score}${toStr}"
        }
        default { "処理完了。メッセージ: $($NodeRow.message)${toStr}" }
    }

    $statusIcon = if ($status -eq "Success") { "[OK]" } elseif ($status -eq "Failed") { "[NG]" } else { "[--]" }
    return [PSCustomObject]@{
        timestamp  = (Get-Date).ToString("HH:mm:ss.fff")
        nodeId     = $nodeId
        role       = $role
        agentId    = "$($NodeRow.agentId)"
        status     = $status
        statusIcon = $statusIcon
        risk       = $risk
        message    = $msg
        durationMs = $dur
        downstream = $downstream
    }
}

function Show-AgentTeamsConversation {
    <#
    .SYNOPSIS
        Agent Teams の会話ログをコンソールにリアルタイム表示します。
    .PARAMETER ConversationLog
        会話エントリ配列 (Get-AgentConversationMessage の戻り値)
    .PARAMETER RunId
        実行識別子
    .PARAMETER LevelTimeline
        DAGレベルタイムライン配列
    .PARAMETER ShowHeader
        ヘッダーバナーを表示するか (既定: $true)
    .PARAMETER ShowSummaryLine
        最終サマリー行を表示するか (既定: $true)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]]$ConversationLog,
        [string]$RunId = "",
        [object[]]$LevelTimeline = @(),
        [bool]$ShowHeader = $true,
        [bool]$ShowSummaryLine = $true
    )

    $sep   = "=" * 66
    $sepThin = "-" * 66

    if ($ShowHeader) {
        Write-Host ""
        Write-Host $sep -ForegroundColor Cyan
        $runLabel = if ($RunId) { "  RunId: $($RunId.Substring(0, [Math]::Min(12, $RunId.Length)))..." } else { "" }
        Write-Host ("  >> Agent Teams 会話ログ <<" + $runLabel) -ForegroundColor Cyan
        Write-Host $sep -ForegroundColor Cyan
    }

    # レベルごとにグループ化して表示
    $byLevel = @{}
    foreach ($entry in $ConversationLog) {
        $lv = if ($entry.PSObject.Properties["level"]) { [int]$entry.level } else { 0 }
        if (-not $byLevel.ContainsKey($lv)) { $byLevel[$lv] = @() }
        $byLevel[$lv] += $entry
    }

    $roleColors = @{
        "planner"    = "Magenta"
        "collector"  = "Yellow"
        "analyzer"   = "Cyan"
        "remediator" = "Green"
        "reporter"   = "Blue"
    }
    $statusColors = @{
        "[OK]" = "Green"
        "[NG]" = "Red"
        "[--]" = "DarkGray"
    }

    foreach ($lv in ($byLevel.Keys | Sort-Object)) {
        $entries = @($byLevel[$lv])
        $lvLabel  = $LevelTimeline | Where-Object { $_.PSObject.Properties["level"] -and [int]$_.level -eq $lv }
        $nodeIds  = if ($lvLabel) { @($lvLabel[0].nodes) -join ", " } else { "" }
        $parallel = if (@($entries).Count -gt 1) { " [並列x$(@($entries).Count)]" } else { "" }

        Write-Host ""
        Write-Host "  [Level ${lv}]${parallel} $nodeIds" -ForegroundColor White

        foreach ($e in $entries) {
            $icon  = Get-AgentNodeIcon -Role $e.role -AgentId $e.agentId
            $color = if ($roleColors.ContainsKey($e.role)) { $roleColors[$e.role] } else { "Gray" }
            $sIcon = $e.statusIcon
            $sCol  = if ($statusColors.ContainsKey($sIcon)) { $statusColors[$sIcon] } else { "Gray" }
            $durStr = if ($e.durationMs -gt 0) { " ({0}ms)" -f $e.durationMs } else { "" }

            Write-Host "  ┌─ " -NoNewline -ForegroundColor DarkGray
            Write-Host "${icon} $($e.nodeId)" -NoNewline -ForegroundColor $color
            Write-Host " ─────────────────────────────────────────" -ForegroundColor DarkGray
            Write-Host "  │ " -NoNewline -ForegroundColor DarkGray
            Write-Host $sIcon -NoNewline -ForegroundColor $sCol
            Write-Host " $($e.message)" -ForegroundColor White
            if ($durStr) {
                Write-Host "  │   " -NoNewline -ForegroundColor DarkGray
                Write-Host "所要時間:${durStr}" -ForegroundColor DarkGray
            }
            Write-Host "  └──────────────────────────────────────────────────────" -ForegroundColor DarkGray

            # Hookイベントがあれば表示（on_error検出用）
            if ($e.status -eq "Failed") {
                Write-Host "  !!! on_error フック発火" -ForegroundColor Red
            }
        }

        # レベル間の矢印
        if ($lv -lt ($byLevel.Keys | Measure-Object -Maximum).Maximum) {
            Write-Host "          |" -ForegroundColor DarkGray
            Write-Host "          V" -ForegroundColor DarkGray
        }
    }

    if ($ShowSummaryLine) {
        $total   = @($ConversationLog).Count
        $success = @($ConversationLog | Where-Object { $_.status -eq "Success" }).Count
        $failed  = @($ConversationLog | Where-Object { $_.status -eq "Failed" }).Count
        $totalMs = @($ConversationLog | ForEach-Object { [int]$_.durationMs } | Measure-Object -Sum).Sum
        $totalSec = [Math]::Round($totalMs / 1000.0, 2)

        Write-Host ""
        Write-Host $sepThin -ForegroundColor Cyan
        $summaryCol = if ($failed -gt 0) { "Red" } else { "Green" }
        Write-Host ("  >> 実行完了 | 成功: ${success}/${total} | 失敗: ${failed} | 累計実行時間: ${totalSec}秒") -ForegroundColor $summaryCol
        Write-Host $sepThin -ForegroundColor Cyan
        Write-Host ""
    }
}

function Export-AgentConversationLog {
    <#
    .SYNOPSIS
        Agent Teams の会話ログを JSON ファイルにエクスポートします。
    .PARAMETER ConversationLog
        会話エントリ配列
    .PARAMETER ReportsDir
        出力先ルートディレクトリ
    .PARAMETER RunId
        実行識別子
    .OUTPUTS
        出力ファイルパス (文字列)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]]$ConversationLog,
        [Parameter(Mandatory)][string]$ReportsDir,
        [Parameter(Mandatory)][string]$RunId
    )

    $convDir = Join-Path $ReportsDir "agent-teams\conversations"
    if (-not (Test-Path $convDir)) { New-Item -ItemType Directory -Path $convDir -Force | Out-Null }

    $outPath = Join-Path $convDir ("AgentConversation_${RunId}.json")
    $export = [PSCustomObject]@{
        schemaVersion = "1.0"
        runId         = $RunId
        exportedAt    = (Get-Date).ToString("s")
        totalEntries  = @($ConversationLog).Count
        successCount  = @($ConversationLog | Where-Object { $_.status -eq "Success" }).Count
        failedCount   = @($ConversationLog | Where-Object { $_.status -eq "Failed" }).Count
        entries       = @($ConversationLog)
    }
    $export | ConvertTo-Json -Depth 16 | Set-Content -Path $outPath -Encoding $script:_enc
    return $outPath
}

Export-ModuleMember -Function `
    New-StandardHookPayload, `
    Convert-HookEntryToSiemLine, `
    Export-HookSiemLines, `
    Invoke-HookAckResync, `
    Invoke-AgentHookEvent, `
    Invoke-McpProviders, `
    Invoke-McpProvidersParallel, `
    Invoke-McpRollbackExecutor, `
    Invoke-AgentTeamsOrchestration, `
    Update-AgentQualityMetrics, `
    Get-AgentTeamsProfilePreset, `
    Show-AgentTeamsConversation, `
    Export-AgentConversationLog, `
    Get-AgentNodeIcon, `
    Get-AgentConversationMessage
