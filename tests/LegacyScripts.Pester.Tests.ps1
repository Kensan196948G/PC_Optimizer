Set-StrictMode -Version Latest

BeforeDiscovery {
    $legacyScripts = @(
        "tests\Test_PCOptimizer.ps1",
        "tests\Test_CLIModes.ps1",
        "tests\Test_TasksValidation.ps1",
        "tests\Test_ReportSchema.ps1",
        "tests\Test_WhatIfExport.ps1"
    )
}

Describe "Legacy Script Tests via Pester" {
    BeforeAll {
        $script:repoRoot = Split-Path $PSScriptRoot -Parent
    }

    foreach ($rel in $legacyScripts) {
        It "passes $rel" {
            $path = Join-Path $script:repoRoot $rel
            Test-Path $path | Should -BeTrue
            & powershell -NoProfile -ExecutionPolicy Bypass -File $path | Out-Null
            $LASTEXITCODE | Should -Be 0
        }
    }
}

