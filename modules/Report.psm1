Set-StrictMode -Version Latest

function New-OptimizerReportData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$InputObject
    )

    [PSCustomObject]@{
        Version      = '1.0'
        GeneratedAt  = (Get-Date).ToString('s')
        Data         = $InputObject
    }
}

function Export-OptimizerReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$ReportData,
        [Parameter(Mandatory)]
        [ValidateSet('html','csv','json')]
        [string]$Format,
        [Parameter(Mandatory)]
        [string]$Path,
        [switch]$UseLocalChartJs,
        [string]$ChartJsScriptPath = ""
    )

    switch ($Format) {
        'json' { $ReportData | ConvertTo-Json -Depth 10 | Set-Content -Path $Path -Encoding utf8 }
        'csv'  {
            $dataObj = if ($ReportData.PSObject.Properties['Data']) { $ReportData.Data } else { $ReportData }
            @($dataObj) | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
        }
        'html' { $ReportData | ConvertTo-Html -Title 'PC Health Report' | Set-Content -Path $Path -Encoding utf8 }
    }

    return $Path
}

function Update-ScoreHistory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$ReportData,
        [string]$HistoryPath = ""
    )

    if (-not $HistoryPath) { return }

    $entry = [PSCustomObject]@{
        recordedAt = (Get-Date).ToString('s')
        score      = if ($ReportData.PSObject.Properties['score']) { $ReportData.score } else { $null }
        cpuScore   = if ($ReportData.PSObject.Properties['cpuScore']) { $ReportData.cpuScore } else { $null }
        memoryScore= if ($ReportData.PSObject.Properties['memoryScore']) { $ReportData.memoryScore } else { $null }
    }

    $history = @()
    if (Test-Path $HistoryPath) {
        try { $history = @(Get-Content -Path $HistoryPath -Raw -Encoding utf8 | ConvertFrom-Json) } catch { $history = @() }
    }

    $history += $entry
    if ($history.Count -gt 30) { $history = @($history | Select-Object -Last 30) }

    $history | ConvertTo-Json -Depth 5 | Set-Content -Path $HistoryPath -Encoding utf8
}

Export-ModuleMember -Function New-OptimizerReportData,Export-OptimizerReport,Update-ScoreHistory
