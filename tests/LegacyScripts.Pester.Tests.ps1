Set-StrictMode -Version Latest

$repoRoot = Split-Path $PSScriptRoot -Parent

Describe "Legacy Script Tests via Pester" {
    $legacyScripts = @(
        "tests\Test_PCOptimizer.ps1",
        "tests\Test_CLIModes.ps1",
        "tests\Test_TasksValidation.ps1",
        "tests\Test_ReportSchema.ps1",
        "tests\Test_WhatIfExport.ps1"
    )

    foreach ($rel in $legacyScripts) {
        It "passes $rel" {
            $path = Join-Path $repoRoot $rel
            Test-Path $path | Should -BeTrue
            & powershell -NoProfile -ExecutionPolicy Bypass -File $path | Out-Null
            $LASTEXITCODE | Should -Be 0
        }
    }
}

