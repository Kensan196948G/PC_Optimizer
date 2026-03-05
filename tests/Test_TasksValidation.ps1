# ==============================================================
# Test_TasksValidation.ps1
# Runtime validation tests for invalid -Tasks arguments
# ==============================================================

param(
    [string]$RepoRoot = (Split-Path $PSScriptRoot -Parent)
)

$scriptPath = Join-Path $RepoRoot "PC_Optimizer.ps1"
$pass = 0
$fail = 0

function Assert-True {
    param([string]$Name, [bool]$Condition, [string]$Detail = "")
    if ($Condition) {
        $script:pass++
        Write-Host "  [PASS] $Name" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "  [FAIL] $Name" -ForegroundColor Red
        if ($Detail) { Write-Host "         => $Detail" -ForegroundColor DarkYellow }
    }
}

Assert-True "TV-01: script exists" (Test-Path $scriptPath) $scriptPath
if ($fail -gt 0) { exit 1 }

# Invalid cases: expect InvalidArgs(3)
$invalidCases = @(
    @{ Name = "TV-02: -Tasks '1,,,' returns InvalidArgs(3)";        Tasks = "1,,,"    },
    @{ Name = "TV-03: -Tasks '1-' returns InvalidArgs(3)";          Tasks = "1-"      },
    @{ Name = "TV-04: -Tasks '-3' returns InvalidArgs(3)";          Tasks = "-3"      },
    @{ Name = "TV-05: -Tasks '1-3-5' returns InvalidArgs(3)";       Tasks = "1-3-5"   },
    @{ Name = "TV-06: -Tasks '1-5,3' returns InvalidArgs(3)";       Tasks = "1-5,3"   },
    @{ Name = "TV-08: -Tasks '0' (below min) returns InvalidArgs(3)";  Tasks = "0"    },
    @{ Name = "TV-09: -Tasks '21' (above max) returns InvalidArgs(3)"; Tasks = "21"   },
    @{ Name = "TV-10: -Tasks '1-25' (range above max) returns InvalidArgs(3)"; Tasks = "1-25" },
    @{ Name = "TV-11: -Tasks '1-5,2-4' (overlapping ranges) returns InvalidArgs(3)"; Tasks = "1-5,2-4" },
    @{ Name = "TV-12: -Tasks '1,2,' (trailing comma) returns InvalidArgs(3)"; Tasks = "1,2," },
    @{ Name = "TV-13: -Tasks '@1' (invalid char) returns InvalidArgs(3)"; Tasks = "@1" },
    @{ Name = "TV-14: -Tasks '5-1' (reverse range) returns InvalidArgs(3)"; Tasks = "5-1" }
)

foreach ($c in $invalidCases) {
    & powershell -NoProfile -ExecutionPolicy Bypass -File $scriptPath `
        -NonInteractive -WhatIf -NoRebootPrompt -Tasks $c.Tasks | Out-Null
    $ec = $LASTEXITCODE
    Assert-True $c.Name ($ec -eq 3) "exit=$ec tasks=$($c.Tasks)"
}

# Valid cases: expect exit code 0
$validCases = @(
    @{ Name = "TV-07: -Tasks '1, 2' (space mixed) exits 0";     Tasks = "1, 2"   },
    @{ Name = "TV-15: -Tasks '20' (max single task) exits 0";   Tasks = "20"     },
    @{ Name = "TV-16: -Tasks '1-20' (full range) exits 0";      Tasks = "1-20"   },
    @{ Name = "TV-17: -Tasks '1,3,5' (non-contiguous) exits 0"; Tasks = "1,3,5"  }
)

foreach ($c in $validCases) {
    & powershell -NoProfile -ExecutionPolicy Bypass -File $scriptPath `
        -NonInteractive -WhatIf -NoRebootPrompt -Tasks $c.Tasks | Out-Null
    $ec = $LASTEXITCODE
    Assert-True $c.Name ($ec -eq 0) "exit=$ec tasks=$($c.Tasks)"
}

Write-Host ""
Write-Host "PASS: $pass / $($pass + $fail)" -ForegroundColor Green
Write-Host "FAIL: $fail / $($pass + $fail)" -ForegroundColor Yellow

if ($fail -gt 0) { exit 1 }
exit 0
