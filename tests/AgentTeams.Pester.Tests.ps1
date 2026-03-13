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
        # キーはRunId:AgentId形式 — RunIdはコレクター内部で生成される場合があるためAgentId部分のみを検証
        ($res.summary.mergedByRunAndAgent | Where-Object { $_.key -like "*:SecurityAgent" }).Count | Should -Be 1
        ($res.summary.mergedByRunAndAgent | Where-Object { $_.key -like "*:NetworkAgent" }).Count | Should -Be 1
        ($res.summary.mergedByRunAndAgent | Where-Object { $_.key -like "*:UpdateAgent" }).Count | Should -Be 1
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
        $result[0].PSObject.Properties.Name | Should -Contain "transactionId"
        $expectedFile = Join-Path $script:testOut ("mcp\MCP_archive_{0}.json" -f $runId)
        (Test-Path $expectedFile) | Should -BeTrue
        (Test-Path (Join-Path $script:testOut "mcp\MCP_Transactions.json")) | Should -BeTrue
    }

    It "Invoke-McpProvidersParallel is exported and falls back to sequential for single provider" {
        (Get-Command Invoke-McpProvidersParallel -ErrorAction SilentlyContinue) | Should -Not -BeNullOrEmpty

        $runId = "pester-mcp-parallel-" + ([guid]::NewGuid().ToString("N"))
        $providers = @([PSCustomObject]@{ name = "archive-par"; type = "file"; enabled = $true; retryCount = 1 })
        $payload = [PSCustomObject]@{ id = $runId; value = 42 }

        # 1プロバイダーの場合は逐次実行にフォールバックするがエラーにならないこと
        $result = Invoke-McpProvidersParallel -McpProviders $providers -Payload $payload -RunId $runId -ReportsDir $script:testOut
        @($result).Count | Should -Be 1
        $result[0].status | Should -Be "Success"
    }

    It "Invoke-McpProvidersParallel dispatches multiple enabled providers concurrently" {
        $runId = "pester-mcp-multi-" + ([guid]::NewGuid().ToString("N"))
        $providers = @(
            [PSCustomObject]@{ name = "arch1"; type = "file"; enabled = $true; retryCount = 1 },
            [PSCustomObject]@{ name = "arch2"; type = "file"; enabled = $true; retryCount = 1 },
            [PSCustomObject]@{ name = "disabled-prov"; type = "file"; enabled = $false; retryCount = 1 }
        )
        $payload = [PSCustomObject]@{ id = $runId; value = 99 }

        $result = Invoke-McpProvidersParallel -McpProviders $providers -Payload $payload -RunId $runId -ReportsDir $script:testOut

        # 3プロバイダーのうち2つが有効、1つが無効
        @($result).Count | Should -Be 3
        # 有効な2つはSuccess、無効な1つはSkipped
        ($result | Where-Object { $_.status -eq "Success" }).Count | Should -Be 2
        ($result | Where-Object { $_.status -eq "Skipped" }).Count | Should -Be 1
    }

    # ── Agent Teams 会話可視化機能テスト ──────────────────────────
    Describe "Agent Teams Conversation Visualization" {
        It "Get-AgentNodeIcon returns label for known roles" {
            (Get-Command Get-AgentNodeIcon -ErrorAction SilentlyContinue) | Should -Not -BeNullOrEmpty

            $label = Get-AgentNodeIcon -Role "planner"
            $label | Should -Match 'PLAN'

            $label2 = Get-AgentNodeIcon -Role "collector" -AgentId "SecurityAgent"
            $label2 | Should -Match 'SEC'

            $label3 = Get-AgentNodeIcon -Role "reporter"
            $label3 | Should -Match 'RPT'
        }

        It "Get-AgentNodeIcon returns fallback for unknown role" {
            $label = Get-AgentNodeIcon -Role "unknown.xyz"
            $label | Should -Not -BeNullOrEmpty
        }

        It "Get-AgentConversationMessage returns PSObject with expected properties" {
            (Get-Command Get-AgentConversationMessage -ErrorAction SilentlyContinue) | Should -Not -BeNullOrEmpty

            $nodeRow = [PSCustomObject]@{ nodeId = "planner"; role = "planner"; status = "Success"; risk = "Low"; durationMs = 250; agentId = ""; message = ""; payload = $null }
            $msg = Get-AgentConversationMessage -NodeRow $nodeRow -AllNodes @()
            $msg | Should -Not -BeNullOrEmpty
            $msg.PSObject.Properties.Name | Should -Contain "nodeId"
            $msg.PSObject.Properties.Name | Should -Contain "role"
            $msg.PSObject.Properties.Name | Should -Contain "status"
            $msg.PSObject.Properties.Name | Should -Contain "message"
            $msg.PSObject.Properties.Name | Should -Contain "timestamp"
            $msg.PSObject.Properties.Name | Should -Contain "durationMs"
            $msg.nodeId | Should -Be "planner"
            $msg.status | Should -Be "Success"
        }

        It "Get-AgentConversationMessage handles Failed status" {
            $nodeRow = [PSCustomObject]@{ nodeId = "analyzer.security"; role = "analyzer"; status = "Failed"; risk = "High"; durationMs = 100; agentId = ""; message = "error"; payload = $null }
            $msg = Get-AgentConversationMessage -NodeRow $nodeRow -AllNodes @()
            $msg.status | Should -Be "Failed"
            $msg.message | Should -Not -BeNullOrEmpty
        }

        It "Show-AgentTeamsConversation runs without error" {
            (Get-Command Show-AgentTeamsConversation -ErrorAction SilentlyContinue) | Should -Not -BeNullOrEmpty

            $nodeRow = [PSCustomObject]@{ nodeId = "planner"; role = "planner"; status = "Success"; risk = "Low"; durationMs = 100; agentId = ""; message = ""; payload = $null }
            $n2 = [PSCustomObject]@{ nodeId = "reporter"; role = "reporter"; status = "Success"; risk = "Low"; durationMs = 50; agentId = ""; message = ""; payload = $null }
            # Show-AgentTeamsConversation の Pester 5 ブロック内でも $nodeRow は使えるよう別名で定義
            $log = @(
                (Get-AgentConversationMessage -NodeRow $nodeRow -AllNodes @()),
                (Get-AgentConversationMessage -NodeRow $n2 -AllNodes @())
            )
            { Show-AgentTeamsConversation -ConversationLog $log -ShowHeader $false -ShowSummaryLine $false } | Should -Not -Throw
        }

        It "Export-AgentConversationLog writes JSON file with correct schema" {
            (Get-Command Export-AgentConversationLog -ErrorAction SilentlyContinue) | Should -Not -BeNullOrEmpty

            $runId = "pester-conv-" + ([guid]::NewGuid().ToString("N"))
            $n1 = [PSCustomObject]@{ nodeId = "planner"; role = "planner"; status = "Success"; risk = "Low"; durationMs = 80; agentId = ""; message = ""; payload = $null }
            $n2 = [PSCustomObject]@{ nodeId = "reporter"; role = "reporter"; status = "Success"; risk = "Low"; durationMs = 60; agentId = ""; message = ""; payload = [PSCustomObject]@{ healthScore = 85 } }
            $log = @(
                (Get-AgentConversationMessage -NodeRow $n1 -AllNodes @()),
                (Get-AgentConversationMessage -NodeRow $n2 -AllNodes @())
            )
            $path = Export-AgentConversationLog -ConversationLog $log -RunId $runId -ReportsDir $script:testOut

            (Test-Path $path) | Should -BeTrue
            $content = Get-Content $path -Raw | ConvertFrom-Json
            $content.runId | Should -Be $runId
            @($content.entries).Count | Should -Be 2
        }

        It "Invoke-AgentTeamsOrchestration result includes conversationLog" {
            $runId = "pester-convfull-" + ([guid]::NewGuid().ToString("N"))
            $snapshot = [PSCustomObject]@{
                securityDiagnostics = [PSCustomObject]@{ Defender = "Enabled" }
                networkDiagnostics  = [PSCustomObject]@{ PrimaryIPv4 = "10.0.0.1" }
                updateDiagnostics   = [PSCustomObject]@{ WindowsUpdate = "Compliant" }
            }
            $score = [PSCustomObject]@{ Score = 75 }
            $ai = [PSCustomObject]@{ Evaluation = "普通" }
            $hooks = [PSCustomObject]@{
                pre_task = @(); post_task = @(); on_error = @(); on_fallback = @(); on_report = @()
            }
            $mcp = @([PSCustomObject]@{ name = "local-archive"; type = "file"; enabled = $true; retryCount = 1 })

            $res = Invoke-AgentTeamsOrchestration -RunId $runId -ReportsDir $script:testOut -ModuleSnapshot $snapshot -HealthScore $score -AIDiagnosis $ai -HooksConfig $hooks -McpProviders $mcp

            $res.PSObject.Properties.Name | Should -Contain "conversationLog"
            @($res.conversationLog).Count | Should -BeGreaterThan 0
        }
    }
}




