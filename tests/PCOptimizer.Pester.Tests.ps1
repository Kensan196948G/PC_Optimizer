Set-StrictMode -Version Latest

$repoRoot = Split-Path $PSScriptRoot -Parent
$scriptPath = Join-Path $repoRoot "PC_Optimizer.ps1"
$batPath = Join-Path $repoRoot "Run_PC_Optimizer.bat"
$advancedModulePath = Join-Path $repoRoot "modules\Advanced.psm1"
$reportsDir = Join-Path $repoRoot "reports"

Describe "PC_Optimizer Regression (Pester)" {
    It "has diagnose/repair mode parameter" {
        $src = Get-Content -Path $scriptPath -Raw -Encoding UTF8
        $src | Should -Match '\[ValidateSet\("repair","diagnose"\)\]\s*\[string\]\$Mode'
    }

    It "Run_PC_Optimizer.bat loads .env before starting PowerShell" {
        $src = Get-Content -Path $batPath -Raw -Encoding UTF8
        $src | Should -Match 'set "ENV_FILE=%SCRIPT_DIR%\.env"'
        $src | Should -Match 'if exist "%ENV_FILE%"'
        $src | Should -Match 'set "!ENV_KEY!=!ENV_VAL!"'
    }

    It "Invoke-AIDiagnosis returns fallback reason when API key is not provided" {
        Import-Module $advancedModulePath -Force
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
        if (Test-Path $reportsDir) {
            $before = @(Get-ChildItem -Path $reportsDir -Filter "Audit_Run_*.json" -File -ErrorAction SilentlyContinue)
        }

        & powershell -NoProfile -ExecutionPolicy Bypass -File $scriptPath -NonInteractive -WhatIf -NoRebootPrompt -Mode diagnose -Tasks "1-3" -FailureMode continue | Out-Null
        $LASTEXITCODE | Should -Be 0

        $after = @(Get-ChildItem -Path $reportsDir -Filter "Audit_Run_*.json" -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTimeUtc)
        $newFile = $after | Where-Object { $before.FullName -notcontains $_.FullName } | Select-Object -Last 1
        $newFile | Should -Not -BeNullOrEmpty

        $obj = Get-Content -Path $newFile.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
        $obj.schemaVersion | Should -Be "1.0"
        $obj.execution.mode | Should -Be "diagnose"
        $obj.execution.whatIf | Should -BeTrue
        $obj.summary.total | Should -BeGreaterOrEqual 1
    }

    It "runs in repair mode (whatif/noninteractive) successfully" {
        & powershell -NoProfile -ExecutionPolicy Bypass -File $scriptPath -NonInteractive -WhatIf -NoRebootPrompt -Mode repair -Tasks "1-2" -FailureMode continue | Out-Null
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
        $entry = @{ status = "OK"; action = "hook.on_start"; runId = "r1"; hostName = "PC1" }
        $line = Convert-HookEntryToSiemLine -Entry $entry -Format "jsonl"
        $line | Should -Match '"severity":3'
        $line | Should -Match '"action":"hook.on_start"'
    }
    It "Convert-HookEntryToSiemLine returns severity 7 for failure" {
        if (-not (Get-Command Convert-HookEntryToSiemLine -ErrorAction SilentlyContinue)) {
            Set-ItResult -Skipped -Because "Convert-HookEntryToSiemLine not exported"; return
        }
        $entry = @{ status = "NG"; action = "hook.on_error"; runId = "r1"; hostName = "PC1" }
        $line = Convert-HookEntryToSiemLine -Entry $entry -Format "jsonl"
        $line | Should -Match '"severity":7'
    }
    It "Export-HookSiemLines writes to file when enabled" {
        if (-not (Get-Command Export-HookSiemLines -ErrorAction SilentlyContinue)) {
            Set-ItResult -Skipped -Because "Export-HookSiemLines not exported"; return
        }
        $entries = @(
            @{ status = "OK"; action = "hook.test"; runId = "r2"; hostName = "PC1" }
        )
        $cfg = @{ enabled = $true; format = "jsonl"; outputDir = $script:tmpSiemDir }
        Export-HookSiemLines -Entries $entries -SiemConfig $cfg
        $files = Get-ChildItem $script:tmpSiemDir -Filter "*.jsonl" -ErrorAction SilentlyContinue
        $files.Count | Should -BeGreaterOrEqual 1
    }
    It "Export-HookSiemLines skips when disabled" {
        if (-not (Get-Command Export-HookSiemLines -ErrorAction SilentlyContinue)) {
            Set-ItResult -Skipped -Because "Export-HookSiemLines not exported"; return
        }
        $entries = @(@{ status = "OK"; action = "hook.test"; runId = "r3"; hostName = "PC1" })
        $cfg = @{ enabled = $false; format = "jsonl"; outputDir = $script:tmpSiemDir }
        { Export-HookSiemLines -Entries $entries -SiemConfig $cfg } | Should -Not -Throw
    }
}
