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
