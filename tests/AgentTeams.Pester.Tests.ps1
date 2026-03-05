Set-StrictMode -Version Latest

$repoRoot = Split-Path $PSScriptRoot -Parent
$orchestrationModulePath = Join-Path $repoRoot "modules\Orchestration.psm1"

Describe "Agent Teams Orchestration (Pester)" {
    BeforeAll {
        Import-Module $orchestrationModulePath -Force
        $testOut = Join-Path $repoRoot "reports\_pester_agent_teams"
        if (Test-Path $testOut) {
            Remove-Item -Path $testOut -Recurse -Force -ErrorAction SilentlyContinue
        }
        New-Item -ItemType Directory -Path $testOut -Force | Out-Null
    }

    It "exports required orchestration functions" {
        (Get-Command Invoke-AgentHookEvent -ErrorAction SilentlyContinue) | Should -Not -BeNullOrEmpty
        (Get-Command Invoke-McpProviders -ErrorAction SilentlyContinue) | Should -Not -BeNullOrEmpty
        (Get-Command Invoke-AgentTeamsOrchestration -ErrorAction SilentlyContinue) | Should -Not -BeNullOrEmpty
    }

    It "has sub-agent plugin modules" {
        (Test-Path (Join-Path $repoRoot "modules\\agents\\SecurityAgent.psm1")) | Should -BeTrue
        (Test-Path (Join-Path $repoRoot "modules\\agents\\NetworkAgent.psm1")) | Should -BeTrue
        (Test-Path (Join-Path $repoRoot "modules\\agents\\UpdateAgent.psm1")) | Should -BeTrue
    }

    It "runs agent teams and merges collector results by runId + agentId" {
        $runId = "pester-agent-" + ([guid]::NewGuid().ToString("N"))
        $snapshot = [PSCustomObject]@{
            securityDiagnostics = [PSCustomObject]@{ Defender = "Enabled"; Firewall = "Enabled" }
            networkDiagnostics  = [PSCustomObject]@{ PrimaryIPv4 = "192.168.1.10" }
            updateDiagnostics   = [PSCustomObject]@{ WindowsUpdate = "Compliant" }
        }
        $score = [PSCustomObject]@{ Score = 88 }
        $ai = [PSCustomObject]@{ Evaluation = "良好" }
        $hooks = [PSCustomObject]@{
            pre_task = @([PSCustomObject]@{ type = "log"; name = "pre" })
            post_task = @([PSCustomObject]@{ type = "log"; name = "post" })
            on_error = @()
            on_fallback = @()
            on_report = @()
        }
        $mcp = @([PSCustomObject]@{ name = "local-archive"; type = "file"; enabled = $true; retryCount = 1 })

        $res = Invoke-AgentTeamsOrchestration -RunId $runId -ReportsDir $testOut -ModuleSnapshot $snapshot -HealthScore $score -AIDiagnosis $ai -HooksConfig $hooks -McpProviders $mcp

        $res.summary.schemaVersion | Should -Be "1.0"
        @($res.summary.collector.results).Count | Should -Be 3
        @($res.summary.mergedByRunAndAgent).Count | Should -Be 3
        @($res.summary.dagTimeline).Count | Should -BeGreaterThan 0
        @($res.summary.hookTimeline).Count | Should -BeGreaterOrEqual 0
        ($res.summary.mergedByRunAndAgent | Where-Object { $_.key -eq "$runId:SecurityAgent" }).Count | Should -Be 1
        ($res.summary.mergedByRunAndAgent | Where-Object { $_.key -eq "$runId:NetworkAgent" }).Count | Should -Be 1
        ($res.summary.mergedByRunAndAgent | Where-Object { $_.key -eq "$runId:UpdateAgent" }).Count | Should -Be 1
        (Test-Path $res.planPath) | Should -BeTrue
        (Test-Path $res.summaryPath) | Should -BeTrue
        $res.summary.PSObject.Properties.Name | Should -Contain "dagExecution"
        $res.summary.PSObject.Properties.Name | Should -Contain "qualityMetrics"
        (Test-Path $res.summary.qualityMetrics.summaryPath) | Should -BeTrue
        (Test-Path $res.summary.qualityMetrics.historyPath) | Should -BeTrue
    }

    It "supports file MCP provider output with idempotency key" {
        $runId = "pester-mcp-" + ([guid]::NewGuid().ToString("N"))
        $providers = @([PSCustomObject]@{ name = "archive"; type = "file"; enabled = $true; retryCount = 1 })
        $payload = [PSCustomObject]@{ id = $runId; value = 1 }

        $result = Invoke-McpProviders -McpProviders $providers -Payload $payload -RunId $runId -ReportsDir $testOut

        @($result).Count | Should -Be 1
        $result[0].status | Should -Be "Success"
        $result[0].PSObject.Properties.Name | Should -Contain "idempotencyKey"
        $expectedFile = Join-Path $testOut ("mcp\MCP_archive_{0}.json" -f $runId)
        (Test-Path $expectedFile) | Should -BeTrue
        (Test-Path (Join-Path $testOut "mcp\MCP_Transactions.json")) | Should -BeTrue
    }
}




