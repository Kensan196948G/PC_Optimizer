Set-StrictMode -Version Latest

Describe "Agent Teams Orchestration (Pester)" {
    BeforeAll {
        $script:repoRoot = Split-Path $PSScriptRoot -Parent
        $script:orchestrationModulePath = Join-Path $script:repoRoot "modules\Orchestration.psm1"
        Import-Module $script:orchestrationModulePath -Force
        $script:testOut = Join-Path $script:repoRoot "reports\_pester_agent_teams"
        if (Test-Path $script:testOut) {
            Remove-Item -Path $script:testOut -Recurse -Force -ErrorAction SilentlyContinue
        }
        New-Item -ItemType Directory -Path $script:testOut -Force | Out-Null
    }

    It "exports required orchestration functions" {
        (Get-Command Invoke-AgentHookEvent -ErrorAction SilentlyContinue) | Should -Not -BeNullOrEmpty
        (Get-Command Invoke-McpProviders -ErrorAction SilentlyContinue) | Should -Not -BeNullOrEmpty
        (Get-Command Invoke-AgentTeamsOrchestration -ErrorAction SilentlyContinue) | Should -Not -BeNullOrEmpty
    }

    It "has sub-agent plugin modules" {
        (Test-Path (Join-Path $script:repoRoot "modules\\agents\\SecurityAgent.psm1")) | Should -BeTrue
        (Test-Path (Join-Path $script:repoRoot "modules\\agents\\NetworkAgent.psm1")) | Should -BeTrue
        (Test-Path (Join-Path $script:repoRoot "modules\\agents\\UpdateAgent.psm1")) | Should -BeTrue
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

        $res = Invoke-AgentTeamsOrchestration -RunId $runId -ReportsDir $script:testOut -ModuleSnapshot $snapshot -HealthScore $score -AIDiagnosis $ai -HooksConfig $hooks -McpProviders $mcp

        $res.summary.schemaVersion | Should -Be "1.1"
        @($res.summary.collector.results).Count | Should -Be 3
        @($res.summary.mergedByRunAndAgent).Count | Should -Be 3
        @($res.summary.dagTimeline).Count | Should -BeGreaterThan 0
        @($res.summary.hookTimeline).Count | Should -BeGreaterOrEqual 0

        # Use member enumeration and direct property comparison to avoid closure-scope issues
        $mergedAgentIds = @($res.summary.mergedByRunAndAgent.agentId)
        $mergedAgentIds | Should -Contain "SecurityAgent"
        $mergedAgentIds | Should -Contain "NetworkAgent"
        $mergedAgentIds | Should -Contain "UpdateAgent"

        # Verify each runId matches the expected value using direct hashtable lookup
        $mergedByAgent = @{}
        foreach ($m in @($res.summary.mergedByRunAndAgent)) { $mergedByAgent[$m.agentId] = $m }
        $mergedByAgent["SecurityAgent"] | Should -Not -BeNullOrEmpty
        $mergedByAgent["SecurityAgent"].runId | Should -Be $runId
        $mergedByAgent["NetworkAgent"] | Should -Not -BeNullOrEmpty
        $mergedByAgent["NetworkAgent"].runId | Should -Be $runId
        $mergedByAgent["UpdateAgent"] | Should -Not -BeNullOrEmpty
        $mergedByAgent["UpdateAgent"].runId | Should -Be $runId

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

        $result = Invoke-McpProviders -McpProviders $providers -Payload $payload -RunId $runId -ReportsDir $script:testOut

        @($result).Count | Should -Be 1
        $result[0].status | Should -Be "Success"
        $result[0].PSObject.Properties.Name | Should -Contain "idempotencyKey"
        $expectedFile = Join-Path $script:testOut ("mcp\MCP_archive_{0}.json" -f $runId)
        (Test-Path $expectedFile) | Should -BeTrue
        (Test-Path (Join-Path $script:testOut "mcp\MCP_Transactions.json")) | Should -BeTrue
    }
}
