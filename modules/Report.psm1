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
        [string]$Path
    )

    switch ($Format) {
        'json' { $ReportData | ConvertTo-Json -Depth 10 | Set-Content -Path $Path -Encoding utf8 }
        'csv'  { @($ReportData.Data) | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 }
        'html' { $ReportData | ConvertTo-Html -Title 'PC Health Report' | Set-Content -Path $Path -Encoding utf8 }
    }

    return $Path
}

Export-ModuleMember -Function New-OptimizerReportData,Export-OptimizerReport
