param(
    [ValidateSet('powershell','pwsh')]
    [string]$Engine = 'powershell'
)
Set-StrictMode -Version Latest

$repoRoot = Split-Path $PSScriptRoot -Parent
$scriptPath = Join-Path $repoRoot 'PC_Optimizer.ps1'
$reportsDir = Join-Path $repoRoot 'reports'
$schemasDir = Join-Path $repoRoot 'docs\schemas'

function Test-RequiredKey {
    param(
        [Parameter(Mandatory)]$Object,
        [Parameter(Mandatory)][string[]]$Required,
        [Parameter(Mandatory)][string]$Name
    )
    foreach ($k in $Required) {
        if (-not $Object.PSObject.Properties[$k]) {
            throw "$Name missing required key: $k"
        }
    }
}

Write-Host '[E2E] Running agent-teams profile in WhatIf diagnose mode...'
& $Engine -NoProfile -ExecutionPolicy Bypass -File $scriptPath -NonInteractive -WhatIf -NoRebootPrompt -Mode diagnose -ExecutionProfile agent-teams -Tasks '20' -FailureMode continue | Out-Null
if ($LASTEXITCODE -ne 0) { throw "PC_Optimizer.ps1 failed with exit code $LASTEXITCODE" }

$audit = Get-ChildItem -Path $reportsDir -Filter 'Audit_Run_*.json' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTimeUtc | Select-Object -Last 1
if (-not $audit) { throw 'Audit_Run_*.json not found.' }
$auditObj = Get-Content -Path $audit.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
Test-RequiredKeys -Object $auditObj -Required @('schemaVersion','runId','execution','summary','ai','outputs') -Name 'Audit'
if ($auditObj.execution.executionProfile -ne 'agent-teams') { throw "executionProfile is not agent-teams: $($auditObj.execution.executionProfile)" }
if (-not $auditObj.agentTeams) { throw 'agentTeams section missing in audit.' }
if (@($auditObj.agentTeams.mergedByRunAndAgent).Count -lt 3) { throw 'mergedByRunAndAgent count is less than 3.' }

$pbiPath = Join-Path $reportsDir 'PowerBI_Dashboard_latest.json'
if (-not (Test-Path $pbiPath)) { throw 'PowerBI_Dashboard_latest.json not found.' }
$pbiObj = Get-Content -Path $pbiPath -Raw -Encoding UTF8 | ConvertFrom-Json
Test-RequiredKeys -Object $pbiObj -Required @('schemaVersion','generatedAt','hostName','score','scoreStatus') -Name 'PowerBI'
if (-not $pbiObj.agent_summary) { throw 'agent_summary missing in PowerBI JSON.' }
if (@($pbiObj.agent_summary.workers).Count -lt 3) { throw 'agent_summary.workers count is less than 3.' }

$summary = Get-ChildItem -Path (Join-Path $reportsDir 'agent-teams') -Filter 'AgentTeams_Summary_*.json' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTimeUtc | Select-Object -Last 1
if (-not $summary) { throw 'AgentTeams_Summary_*.json not found.' }
$summaryObj = Get-Content -Path $summary.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
Test-RequiredKeys -Object $summaryObj -Required @('schemaVersion','runId','planner','collector','analyzer','remediator','reporter','mergedByRunAndAgent') -Name 'AgentTeamsSummary'
if (-not $summaryObj.qualityMetrics) { throw 'qualityMetrics missing in AgentTeams summary.' }
if (-not $summaryObj.mcp) { throw 'mcp result missing in AgentTeams summary.' }
if (-not $summaryObj.PSObject.Properties['dagTimeline']) { throw 'dagTimeline property missing in AgentTeams summary.' }
if (-not $summaryObj.PSObject.Properties['hookTimeline']) { throw 'hookTimeline property missing in AgentTeams summary.' }
if ($summaryObj.collector.parallelMode -ne 'batched') { throw "collector.parallelMode expected batched, got: $($summaryObj.collector.parallelMode)" }

$historyPath = Join-Path $reportsDir 'agent-teams\metrics\AgentMetrics_History.json'
if (-not (Test-Path $historyPath)) { throw 'AgentMetrics_History.json not found.' }

$ledgerPath = Join-Path $reportsDir 'mcp\MCP_Transactions.json'
if (-not (Test-Path $ledgerPath)) { throw 'MCP_Transactions.json not found.' }

$hookQueueStatePath = Join-Path $repoRoot 'logs\hooks\queue\state.json'
if (-not (Test-Path $hookQueueStatePath)) { throw 'hooks queue state.json not found.' }

if (-not (Test-Path (Join-Path $schemasDir 'audit-run-v1.schema.json'))) { throw 'audit-run-v1.schema.json not found.' }
if (-not (Test-Path (Join-Path $schemasDir 'powerbi-dashboard-v1.schema.json'))) { throw 'powerbi-dashboard-v1.schema.json not found.' }
if (-not (Test-Path (Join-Path $schemasDir 'agent-teams-summary-v1.schema.json'))) { throw 'agent-teams-summary-v1.schema.json not found.' }

if (-not $pbiObj.PSObject.Properties['agent_dag_timeline']) { throw 'agent_dag_timeline property missing in PowerBI JSON.' }
if (-not $pbiObj.PSObject.Properties['hook_timeline']) { throw 'hook_timeline property missing in PowerBI JSON.' }

Write-Host '[E2E] agent-teams validation succeeded.'
