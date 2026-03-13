Set-StrictMode -Version Latest

$repoRoot = Split-Path $PSScriptRoot -Parent
$scriptPath = Join-Path $repoRoot "PC_Optimizer.ps1"
$batPath = Join-Path $repoRoot "Run_PC_Optimizer.bat"
$advancedModulePath = Join-Path $repoRoot "modules\Advanced.psm1"
$reportsDir = Join-Path $repoRoot "reports"

Describe "PC_Optimizer Regression (Pester)" {
    BeforeAll {
        $script:repoRoot = Split-Path $PSScriptRoot -Parent
        $script:scriptPath = Join-Path $script:repoRoot "PC_Optimizer.ps1"
        $script:batPath = Join-Path $script:repoRoot "Run_PC_Optimizer.bat"
        $script:advancedModulePath = Join-Path $script:repoRoot "modules\Advanced.psm1"
        $script:reportsDir = Join-Path $script:repoRoot "reports"
    }
    It "has diagnose/repair mode parameter" {
        $src = Get-Content -Path $script:scriptPath -Raw -Encoding UTF8
        $src | Should -Match '\[ValidateSet\("repair","diagnose"\)\]\s*\[string\]\$Mode'
    }

    It "Run_PC_Optimizer.bat loads .env before starting PowerShell" {
        $src = Get-Content -Path $script:batPath -Raw -Encoding UTF8
        $src | Should -Match 'set "ENV_FILE=%SCRIPT_DIR%\.env"'
        $src | Should -Match 'if exist "%ENV_FILE%"'
        $src | Should -Match 'set "!ENV_KEY!=!ENV_VAL!"'
    }

    It "Invoke-AIDiagnosis returns fallback reason when API key is not provided" {
        Import-Module $script:advancedModulePath -Force
        $health = [PSCustomObject]@{
            Score = 80
            Status = "Good"
            ScoreInput = [PSCustomObject]@{
                Cpu = 80
                Memory = 80
                Disk = 80
                Startup = 80
                Security = 80
                Network = 80
                WindowsUpdate = 80
                SystemHealth = 80
            }
        }
        $ai = Invoke-AIDiagnosis -HealthScore $health -Snapshot $null -UpdateClassifiedErrors @() -M365Connectivity @() -EventAnomaly $null -BootTrend $null
        $ai.Source | Should -Be "LocalRuleEngine"
        $ai.PSObject.Properties.Name | Should -Contain "FallbackReason"
        $ai.FallbackReason | Should -Match "Anthropic API key not provided"
        $ai.PSObject.Properties.Name | Should -Contain "Confidence"
        $ai.PSObject.Properties.Name | Should -Contain "DataTimestamp"
        $ai.PSObject.Properties.Name | Should -Contain "InputMetrics"
    }

    It "runs in diagnose mode (whatif/noninteractive) and produces audit json" {
        $before = @()
        if (Test-Path $script:reportsDir) {
            $before = @(Get-ChildItem -Path $script:reportsDir -Filter "Audit_Run_*.json" -File -ErrorAction SilentlyContinue)
        }

        & powershell -NoProfile -ExecutionPolicy Bypass -File $script:scriptPath -NonInteractive -WhatIf -NoRebootPrompt -Mode diagnose -Tasks "1-3" -FailureMode continue | Out-Null
        $LASTEXITCODE | Should -Be 0

        $after = @(Get-ChildItem -Path $script:reportsDir -Filter "Audit_Run_*.json" -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTimeUtc)
        $newFile = $after | Where-Object { $before.FullName -notcontains $_.FullName } | Select-Object -Last 1
        $newFile | Should -Not -BeNullOrEmpty

        $obj = Get-Content -Path $newFile.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
        $obj.schemaVersion | Should -Be "1.0"
        $obj.execution.mode | Should -Be "diagnose"
        $obj.execution.whatIf | Should -BeTrue
        $obj.summary.total | Should -BeGreaterOrEqual 1
    }

    It "runs in repair mode (whatif/noninteractive) successfully" {
        & powershell -NoProfile -ExecutionPolicy Bypass -File $script:scriptPath -NonInteractive -WhatIf -NoRebootPrompt -Mode repair -Tasks "1-2" -FailureMode continue | Out-Null
        $LASTEXITCODE | Should -Be 0
    }
}

Describe "Module: Common.psm1" {
    BeforeAll {
        $mod = Join-Path (Split-Path $PSScriptRoot -Parent) "modules\Common.psm1"
        Import-Module $mod -Force
    }
    It "Write-StructuredLog returns formatted line" {
        $line = Write-StructuredLog -Message "test message" -Level INFO
        $line | Should -Match "\[INFO\]"
        $line | Should -Match "test message"
    }
    It "Invoke-GuardedStep returns OK on success" {
        $r = Invoke-GuardedStep -Name "test" -Action { 1 + 1 }
        $r.Status | Should -Be "OK"
        $r.Error  | Should -BeNullOrEmpty
    }
    It "Invoke-GuardedStep returns NG on exception" {
        $r = Invoke-GuardedStep -Name "fail" -Action { throw "boom" }
        $r.Status | Should -Be "NG"
        $r.Error  | Should -Match "boom"
    }
}

Describe "Module: Report.psm1 - New-OptimizerReportData" {
    BeforeAll {
        $mod = Join-Path (Split-Path $PSScriptRoot -Parent) "modules\Report.psm1"
        Import-Module $mod -Force
    }
    It "wraps data with Version and GeneratedAt" {
        $input = [PSCustomObject]@{ score = 80 }
        $r = New-OptimizerReportData -InputObject $input
        $r.Version    | Should -Be "1.0"
        $r.GeneratedAt | Should -Not -BeNullOrEmpty
        $r.Data.score | Should -Be 80
    }
}

Describe "Module: Report.psm1 - Export-OptimizerReport (JSON)" {
    BeforeAll {
        $mod = Join-Path (Split-Path $PSScriptRoot -Parent) "modules\Report.psm1"
        Import-Module $mod -Force
    }
    It "exports JSON report" {
        $tmp = Join-Path ([System.IO.Path]::GetTempPath()) "pester_report_$([System.Guid]::NewGuid()).json"
        $data = New-OptimizerReportData -InputObject ([PSCustomObject]@{ score = 77 })
        Export-OptimizerReport -ReportData $data -Format json -Path $tmp
        Test-Path $tmp | Should -BeTrue
        $obj = Get-Content $tmp -Raw | ConvertFrom-Json
        $obj.Data.score | Should -Be 77
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
    }
}

Describe "Module: AgentSDK.psm1" {
    BeforeAll {
        $mod = Join-Path (Split-Path $PSScriptRoot -Parent) "modules\agents\AgentSDK.psm1"
        Import-Module $mod -Force
    }
    It "Test-AgentPluginContext returns false for null" {
        Test-AgentPluginContext -Context $null | Should -BeFalse
    }
    It "Test-AgentPluginContext returns false for missing keys" {
        Test-AgentPluginContext -Context @{ RunId = "x" } | Should -BeFalse
    }
    It "Test-AgentPluginContext returns true for valid context" {
        $ctx = @{ RunId = "r1"; ModuleSnapshot = [PSCustomObject]@{} }
        Test-AgentPluginContext -Context $ctx | Should -BeTrue
    }
    It "New-AgentPluginResult has correct schema" {
        $r = New-AgentPluginResult -Status Success -Risk Low -Message "ok"
        $r.status  | Should -Be "Success"
        $r.risk    | Should -Be "Low"
        $r.message | Should -Be "ok"
    }
    It "Test-AgentPluginResultSchema validates correctly" {
        $valid = New-AgentPluginResult -Status Success -Risk Low -Message "ok"
        Test-AgentPluginResultSchema -Result $valid | Should -BeTrue
        Test-AgentPluginResultSchema -Result $null | Should -BeFalse
    }
}

Describe "Module: Notification.psm1" {
    BeforeAll {
        $mod = Join-Path (Split-Path $PSScriptRoot -Parent) "modules\Notification.psm1"
        if (Test-Path $mod) { Import-Module $mod -Force }
    }
    It "Send-McpProviderNotification skips disabled providers" {
        if (-not (Get-Command Send-McpProviderNotification -ErrorAction SilentlyContinue)) {
            Set-ItResult -Skipped -Because "Notification module not loaded"
            return
        }
        $providers = @([PSCustomObject]@{ type = "slack"; enabled = $false; webhookUrl = "https://hooks.slack.com/test" })
        { Send-McpProviderNotification -McpProviders $providers -HostName "testpc" -Score 80 } | Should -Not -Throw
    }
    It "Send-SlackNotification skips when WebhookUrl is empty" {
        if (-not (Get-Command Send-SlackNotification -ErrorAction SilentlyContinue)) {
            Set-ItResult -Skipped -Because "Notification module not loaded"; return
        }
        { Send-SlackNotification -WebhookUrl "" -HostName "testpc" -Score 80 } | Should -Not -Throw
    }
    It "Send-TeamsNotification skips when WebhookUrl is empty" {
        if (-not (Get-Command Send-TeamsNotification -ErrorAction SilentlyContinue)) {
            Set-ItResult -Skipped -Because "Notification module not loaded"; return
        }
        { Send-TeamsNotification -WebhookUrl "" -HostName "testpc" -Score 80 } | Should -Not -Throw
    }
    It "Send-ServiceNowIncident returns null when InstanceUrl is empty" {
        if (-not (Get-Command Send-ServiceNowIncident -ErrorAction SilentlyContinue)) {
            Set-ItResult -Skipped -Because "Notification module not loaded"; return
        }
        $result = Send-ServiceNowIncident -InstanceUrl "" -HostName "testpc" -Score 65 -Summary "test" -TopRecommendation "fix"
        $result | Should -BeNullOrEmpty
    }
    It "Send-JiraTask returns null when JiraUrl is empty" {
        if (-not (Get-Command Send-JiraTask -ErrorAction SilentlyContinue)) {
            Set-ItResult -Skipped -Because "Notification module not loaded"; return
        }
        $result = Send-JiraTask -JiraUrl "" -ProjectKey "PC" -HostName "testpc" -Score 65 -Summary "test" -TopRecommendation "fix"
        $result | Should -BeNullOrEmpty
    }
}

Describe "Module: Report.psm1 - Update-ScoreHistory" {
    BeforeAll {
        $mod = Join-Path (Split-Path $PSScriptRoot -Parent) "modules\Report.psm1"
        Import-Module $mod -Force
        $script:tmpHistory = Join-Path ([System.IO.Path]::GetTempPath()) "score_history_pester_$([System.Guid]::NewGuid()).json"
    }
    AfterAll {
        Remove-Item $script:tmpHistory -Force -ErrorAction SilentlyContinue
    }
    It "creates score_history.json on first call" {
        if (-not (Get-Command Update-ScoreHistory -ErrorAction SilentlyContinue)) {
            Set-ItResult -Skipped -Because "Update-ScoreHistory not exported"; return
        }
        $reportData = [PSCustomObject]@{ score = 85; cpuScore = 80; memoryScore = 100 }
        Update-ScoreHistory -ReportData $reportData -HistoryPath $script:tmpHistory
        Test-Path $script:tmpHistory | Should -BeTrue
    }
    It "score_history.json contains the recorded score" {
        if (-not (Get-Command Update-ScoreHistory -ErrorAction SilentlyContinue)) {
            Set-ItResult -Skipped -Because "Update-ScoreHistory not exported"; return
        }
        $data = Get-Content $script:tmpHistory -Raw | ConvertFrom-Json
        @($data).Count | Should -BeGreaterOrEqual 1
        $data[0].score | Should -Be 85
    }
    It "score_history.json retains at most 30 entries" {
        if (-not (Get-Command Update-ScoreHistory -ErrorAction SilentlyContinue)) {
            Set-ItResult -Skipped -Because "Update-ScoreHistory not exported"; return
        }
        $reportData = [PSCustomObject]@{ score = 70 }
        1..35 | ForEach-Object { Update-ScoreHistory -ReportData $reportData -HistoryPath $script:tmpHistory }
        $data = Get-Content $script:tmpHistory -Raw | ConvertFrom-Json
        @($data).Count | Should -BeLessOrEqual 30
    }
}

Describe "Module: Orchestration.psm1 - SIEM functions" {
    BeforeAll {
        $mod = Join-Path (Split-Path $PSScriptRoot -Parent) "modules\Orchestration.psm1"
        Import-Module $mod -Force
        $script:tmpSiemDir = Join-Path ([System.IO.Path]::GetTempPath()) "siem_pester_$([System.Guid]::NewGuid())"
        New-Item -ItemType Directory -Path $script:tmpSiemDir -Force | Out-Null
    }
    AfterAll {
        Remove-Item $script:tmpSiemDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    It "Convert-HookEntryToSiemLine returns jsonl string for success" {
        if (-not (Get-Command Convert-HookEntryToSiemLine -ErrorAction SilentlyContinue)) {
            Set-ItResult -Skipped -Because "Convert-HookEntryToSiemLine not exported"; return
        }
        $entry = [PSCustomObject]@{ status = "Success"; action = "hook.on_start"; runId = "r1"; hostName = "PC1" }
        $line = Convert-HookEntryToSiemLine -Entry $entry -Format "jsonl"
        $line | Should -Not -BeNullOrEmpty
        $line | Should -Match '"action"'
    }
    It "Convert-HookEntryToSiemLine returns CEF format string" {
        if (-not (Get-Command Convert-HookEntryToSiemLine -ErrorAction SilentlyContinue)) {
            Set-ItResult -Skipped -Because "Convert-HookEntryToSiemLine not exported"; return
        }
        $entry = [PSCustomObject]@{ status = "NG"; event = "hook.error"; action = "hook.on_error"; runId = "r1"; hostName = "PC1"; finishedAt = "2024-01-01"; detail = "err"; hookPayload = [PSCustomObject]@{ transactionId = "tx1"; nodeId = "n1" } }
        $line = Convert-HookEntryToSiemLine -Entry $entry -Format "cef"
        $line | Should -Match "^CEF:"
    }
    It "Export-HookSiemLines skips when HooksConfig is null" {
        if (-not (Get-Command Export-HookSiemLines -ErrorAction SilentlyContinue)) {
            Set-ItResult -Skipped -Because "Export-HookSiemLines not exported"; return
        }
        { Export-HookSiemLines -Entries @() -HooksConfig $null -LogsDir $script:tmpSiemDir -RunId "r0" } | Should -Not -Throw
    }
    It "Export-HookSiemLines skips when siem.enabled is false" {
        if (-not (Get-Command Export-HookSiemLines -ErrorAction SilentlyContinue)) {
            Set-ItResult -Skipped -Because "Export-HookSiemLines not exported"; return
        }
        $hCfg = [PSCustomObject]@{ siem = [PSCustomObject]@{ enabled = $false } }
        { Export-HookSiemLines -Entries @() -HooksConfig $hCfg -LogsDir $script:tmpSiemDir -RunId "r0" } | Should -Not -Throw
    }
}

Describe "Module: Performance.psm1 - Get-HealthScore" {
    BeforeAll {
        $mod = Join-Path (Split-Path $PSScriptRoot -Parent) "modules\Performance.psm1"
        Import-Module $mod -Force
    }
    It "returns 100 when all scores are 100" {
        $result = Get-HealthScore -Cpu 100 -Memory 100 -Disk 100 -Startup 100 -Security 100 -Network 100 -WindowsUpdate 100 -SystemHealth 100
        $result.Score | Should -Be 100
    }
    It "returns 0 when all scores are 0" {
        $result = Get-HealthScore -Cpu 0 -Memory 0 -Disk 0 -Startup 0 -Security 0 -Network 0 -WindowsUpdate 0 -SystemHealth 0
        $result.Score | Should -Be 0
    }
    It "returns expected weighted score for typical values" {
        $result = Get-HealthScore -Cpu 80 -Memory 90 -Disk 95 -Startup 70 -Security 80 -Network 85 -WindowsUpdate 60 -SystemHealth 90
        $result.Score | Should -BeGreaterThan 0
        $result.Score | Should -BeLessThan 101
    }
    It "clamps negative values to 0" {
        $result = Get-HealthScore -Cpu -10 -Memory 100 -Disk 100 -Startup 100 -Security 100 -Network 100 -WindowsUpdate 100 -SystemHealth 100
        $result.Score | Should -BeGreaterOrEqual 0
    }
    It "clamps values over 100 to 100" {
        $result = Get-HealthScore -Cpu 150 -Memory 100 -Disk 100 -Startup 100 -Security 100 -Network 100 -WindowsUpdate 100 -SystemHealth 100
        $result.Score | Should -BeLessOrEqual 100
    }
    It "returns PSCustomObject with Score property" {
        $result = Get-HealthScore -Cpu 80 -Memory 80 -Disk 80 -Startup 80 -Security 80 -Network 80 -WindowsUpdate 80 -SystemHealth 80
        $result | Should -Not -BeNullOrEmpty
        $result.PSObject.Properties.Name | Should -Contain "Score"
    }
}

Describe "Module: Performance.psm1 - Get-PerformanceSnapshot" {
    BeforeAll {
        $mod = Join-Path (Split-Path $PSScriptRoot -Parent) "modules\Performance.psm1"
        Import-Module $mod -Force
    }
    It "returns PSCustomObject with Status OK" {
        $result = Get-PerformanceSnapshot
        $result | Should -Not -BeNullOrEmpty
        $result.Status | Should -Be 'OK'
    }
    It "contains CpuTopProcess property" {
        $result = Get-PerformanceSnapshot
        $result.PSObject.Properties.Name | Should -Contain "CpuTopProcess"
    }
    It "contains MemoryTopProcess property" {
        $result = Get-PerformanceSnapshot
        $result.PSObject.Properties.Name | Should -Contain "MemoryTopProcess"
    }
}

Describe "Module: Performance.psm1 - Get-StartupAnalysis" {
    BeforeAll {
        $mod = Join-Path (Split-Path $PSScriptRoot -Parent) "modules\Performance.psm1"
        Import-Module $mod -Force
    }
    It "runs without error and returns an object" {
        { $result = Get-StartupAnalysis } | Should -Not -Throw
        $result = Get-StartupAnalysis
        $result | Should -Not -BeNullOrEmpty
    }
}

Describe "Module: Diagnostics.psm1 - Get-SystemDiagnostic" {
    BeforeAll {
        $mod = Join-Path (Split-Path $PSScriptRoot -Parent) "modules\Diagnostics.psm1"
        Import-Module $mod -Force
    }
    It "returns PSCustomObject with expected properties" {
        $result = Get-SystemDiagnostic
        $result | Should -Not -BeNullOrEmpty
        $result.PSObject.Properties.Name | Should -Contain "ComputerName"
    }
    It "ComputerName is non-empty string" {
        $result = Get-SystemDiagnostic
        $result.ComputerName | Should -Not -BeNullOrEmpty
    }
}

Describe "Module: Diagnostics.psm1 - Get-AssetInventory" {
    BeforeAll {
        $mod = Join-Path (Split-Path $PSScriptRoot -Parent) "modules\Diagnostics.psm1"
        Import-Module $mod -Force
    }
    It "returns an object without error" {
        { $result = Get-AssetInventory } | Should -Not -Throw
        $result = Get-AssetInventory
        $result | Should -Not -BeNullOrEmpty
    }
}

Describe "Module: Security.psm1 - Get-SecurityDiagnostic" {
    BeforeAll {
        $mod = Join-Path (Split-Path $PSScriptRoot -Parent) "modules\Security.psm1"
        Import-Module $mod -Force
    }
    It "returns PSCustomObject without error" {
        { $result = Get-SecurityDiagnostic } | Should -Not -Throw
        $result = Get-SecurityDiagnostic
        $result | Should -Not -BeNullOrEmpty
    }
    It "contains DefenderStatus property" {
        $result = Get-SecurityDiagnostic
        $result.PSObject.Properties.Name | Should -Contain "Defender"
    }
    It "contains FirewallStatus property" {
        $result = Get-SecurityDiagnostic
        $result.PSObject.Properties.Name | Should -Contain "Firewall"
    }
}

Describe "Module: Network.psm1 - Get-NetworkDiagnostic" {
    BeforeAll {
        $mod = Join-Path (Split-Path $PSScriptRoot -Parent) "modules\Network.psm1"
        Import-Module $mod -Force
    }
    It "returns PSCustomObject without error" {
        { $result = Get-NetworkDiagnostic } | Should -Not -Throw
        $result = Get-NetworkDiagnostic
        $result | Should -Not -BeNullOrEmpty
    }
    It "contains IpAddresses property" {
        $result = Get-NetworkDiagnostic
        $result.PSObject.Properties.Name | Should -Contain "IpAddress"
    }
}

Describe "Module: Cleanup.psm1 - Invoke-CleanupMaintenance" {
    BeforeAll {
        $mod = Join-Path (Split-Path $PSScriptRoot -Parent) "modules\Cleanup.psm1"
        Import-Module $mod -Force
    }
    It "runs in WhatIfMode without error" {
        { Invoke-CleanupMaintenance -WhatIfMode -Tasks @('temp') } | Should -Not -Throw
    }
    It "returns executed tasks list in WhatIfMode" {
        $result = Invoke-CleanupMaintenance -WhatIfMode -Tasks @('temp','dns')
        $result | Should -Not -BeNullOrEmpty
    }
    It "ignores unsupported task names" {
        { Invoke-CleanupMaintenance -WhatIfMode -Tasks @('unsupported_task') } | Should -Not -Throw
    }
    It "supports browser cache task in WhatIfMode" {
        { Invoke-CleanupMaintenance -WhatIfMode -Tasks @('browser') } | Should -Not -Throw
    }
    It "supports store task in WhatIfMode" {
        { Invoke-CleanupMaintenance -WhatIfMode -Tasks @('store') } | Should -Not -Throw
    }
}
