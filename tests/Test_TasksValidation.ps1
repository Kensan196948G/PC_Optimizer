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

# ── 異常系: InvalidArgs(3) が期待されるケース ──
$invalidCases = @(
    @{ Name = "TV-02: -Tasks '1,,,' returns InvalidArgs(3)";        Tasks = "1,,,"    },
    @{ Name = "TV-03: -Tasks '1-' returns InvalidArgs(3)";          Tasks = "1-"      },
    @{ Name = "TV-04: -Tasks '-3' returns InvalidArgs(3)";          Tasks = "-3"      },
    @{ Name = "TV-05: -Tasks '1-3-5' returns InvalidArgs(3)";       Tasks = "1-3-5"   },
    @{ Name = "TV-06: -Tasks '1-5,3' returns InvalidArgs(3)";       Tasks = "1-5,3"   },
    # v4.0.1 追加: 個別タスク範囲外
    @{ Name = "TV-08: -Tasks '0' (below min) returns InvalidArgs(3)";  Tasks = "0"    },
    @{ Name = "TV-09: -Tasks '21' (above max) returns InvalidArgs(3)"; Tasks = "21"   },
    # v4.0.1 追加: 範囲指定の上限超過
    @{ Name = "TV-10: -Tasks '1-25' (range above max) returns InvalidArgs(3)"; Tasks = "1-25" },
    # v4.0.1 追加: 重複する範囲指定
    @{ Name = "TV-11: -Tasks '1-5,2-4' (overlapping ranges) returns InvalidArgs(3)"; Tasks = "1-5,2-4" },
    # v4.0.1 追加: 末尾カンマ（スペースあり）
    @{ Name = "TV-12: -Tasks '1,2,' (trailing comma) returns InvalidArgs(3)"; Tasks = "1,2," },
    # v4.0.1 追加: 未サポート文字
    @{ Name = "TV-13: -Tasks '@1' (invalid char) returns InvalidArgs(3)"; Tasks = "@1" },
    # v4.0.1 追加: 逆順範囲
    @{ Name = "TV-14: -Tasks '5-1' (reverse range) returns InvalidArgs(3)"; Tasks = "5-1" }
)

foreach ($c in $invalidCases) {
    & powershell -NoProfile -ExecutionPolicy Bypass -File $scriptPath `
        -NonInteractive -WhatIf -NoRebootPrompt -Tasks $c.Tasks | Out-Null
    $ec = $LASTEXITCODE
    Assert-True $c.Name ($ec -eq 3) "exit=$ec tasks=$($c.Tasks)"
}

# ── 正常系: 終了コード 0 が期待されるケース ──
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
