Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path $PSScriptRoot -Parent
$guiScript = Join-Path $repoRoot 'GUI\PC_Optimizer_GUI.ps1'
$guiBat = Join-Path $repoRoot 'GUI\Run_PC_Optimizer_GUI.bat'
$docsGuiDir = Join-Path $repoRoot 'docs\GUI'

function Assert-True {
    param(
        [string]$Name,
        [bool]$Condition,
        [string]$Detail = ''
    )

    if ($Condition) {
        Write-Host "[PASS] $Name" -ForegroundColor Green
    } else {
        Write-Host "[FAIL] $Name" -ForegroundColor Red
        if ($Detail) {
            Write-Host "       $Detail" -ForegroundColor Yellow
        }
        exit 1
    }
}

Assert-True 'GUI script exists' (Test-Path -LiteralPath $guiScript) $guiScript
Assert-True 'GUI batch exists' (Test-Path -LiteralPath $guiBat) $guiBat
Assert-True 'GUI docs dir exists' (Test-Path -LiteralPath $docsGuiDir) $docsGuiDir

Assert-True 'GUI doc exists: README.md' (Test-Path -LiteralPath (Join-Path $docsGuiDir 'README.md')) (Join-Path $docsGuiDir 'README.md')

$docFiles = @(Get-ChildItem -LiteralPath $docsGuiDir -Filter '*.md' -File -ErrorAction SilentlyContinue)
Assert-True 'GUI markdown docs count is 6 or more' ($docFiles.Count -ge 6) "count=$($docFiles.Count)"

$tokens = $null
$errors = $null
[void][System.Management.Automation.Language.Parser]::ParseFile($guiScript, [ref]$tokens, [ref]$errors)
Assert-True 'GUI script parses without syntax error' ($errors.Count -eq 0) ($errors | ForEach-Object { $_.Message } | Out-String)

$content = Get-Content -LiteralPath $guiScript -Raw -Encoding UTF8
Assert-True 'GUI script launches backend CLI' ($content -match 'PC_Optimizer\.ps1') 'backend reference missing'
Assert-True 'GUI script enforces NonInteractive' ($content -match '\-NonInteractive') 'NonInteractive missing'
Assert-True 'GUI script enforces NoRebootPrompt' ($content -match '\-NoRebootPrompt') 'NoRebootPrompt missing'
Assert-True 'GUI script uses Windows Forms' ($content -match 'System\.Windows\.Forms') 'WinForms missing'

Write-Host 'GUI static checks completed.' -ForegroundColor Cyan

