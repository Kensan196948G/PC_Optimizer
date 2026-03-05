Set-StrictMode -Version Latest
$script:_enc = if ($PSVersionTable.PSVersion.Major -ge 7) { 'utf8NoBOM' } else { 'UTF8' }

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER InputObject
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
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

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER ReportData
Parameter description

.PARAMETER Format
Parameter description

.PARAMETER Path
Parameter description

.PARAMETER UseLocalChartJs
Parameter description

.PARAMETER ChartJsScriptPath
Parameter description

.PARAMETER HostName
Parameter description

.PARAMETER ExecFolder
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
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
        [string]$ChartJsScriptPath = "assets/chart.umd.min.js",
        [string]$HostName = "",
        [string]$ExecFolder = ""
    )

    switch ($Format) {
        'json' {
            $ReportData | ConvertTo-Json -Depth 10 | Set-Content -Path $Path -Encoding $script:_enc
        }
        'csv'  {
            $rows = @()
            if ($ReportData.PSObject.Properties.Name -contains 'tasks' -and $ReportData.tasks) {
                $rows = @($ReportData.tasks)
            } elseif ($ReportData -is [System.Collections.IEnumerable] -and -not ($ReportData -is [string])) {
                $rows = @($ReportData)
            } else {
                $rows = @($ReportData)
            }
            $rows | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
        }
        'html' {
            $score = if ($ReportData.PSObject.Properties.Name -contains 'score') { [int]$ReportData.score } else { 0 }
            $cpu = if ($ReportData.PSObject.Properties.Name -contains 'cpuScore') { [int]$ReportData.cpuScore } else { 0 }
            $memory = if ($ReportData.PSObject.Properties.Name -contains 'memoryScore') { [int]$ReportData.memoryScore } else { 0 }
            $disk      = if ($ReportData.PSObject.Properties.Name -contains 'diskScore') { [int]$ReportData.diskScore } else { 0 }
            $security  = if ($ReportData.PSObject.Properties.Name -contains 'securityScore') { [int]$ReportData.securityScore } else { 0 }
            $network   = if ($ReportData.PSObject.Properties.Name -contains 'networkScore') { [int]$ReportData.networkScore } else { 0 }
            $winUpdate = if ($ReportData.PSObject.Properties.Name -contains 'windowsUpdateScore') { [int]$ReportData.windowsUpdateScore } else { 0 }
            $startup   = if ($ReportData.PSObject.Properties.Name -contains 'startupScore') { [int]$ReportData.startupScore } else { 0 }
            $sysHealth = if ($ReportData.PSObject.Properties.Name -contains 'systemHealthScore') { [int]$ReportData.systemHealthScore } else { 0 }
            $scoreHistory = @()
            try {
                $scoreHistory = @(Update-ScoreHistory -ReportData $ReportData)
            } catch {
                $scoreHistory = @()
            }
            $trendLabels = ($scoreHistory | ForEach-Object { '"' + $_.runAt + '"' }) -join ','
            $trendScores = ($scoreHistory | ForEach-Object { $_.score }) -join ','
            $aiSection = ""
            if ($ReportData.PSObject.Properties.Name -contains 'aiDiagnosis' -and $ReportData.aiDiagnosis) {
                $ai = $ReportData.aiDiagnosis
                $aiHeadline = [System.Net.WebUtility]::HtmlEncode("$($ai.Headline)")
                $aiEval = [System.Net.WebUtility]::HtmlEncode("$($ai.Evaluation)")
                $aiSummary = if ($ai.PSObject.Properties.Name -contains 'Summary' -and $ai.Summary) { [System.Net.WebUtility]::HtmlEncode("$($ai.Summary)") } else { "" }
                $aiSource = [System.Net.WebUtility]::HtmlEncode("$($ai.Source)")
                $aiPromptVersion = if ($ai.PSObject.Properties.Name -contains 'PromptVersion' -and $ai.PromptVersion) { [System.Net.WebUtility]::HtmlEncode("$($ai.PromptVersion)") } else { "N/A" }
                $aiConfidence = if ($ai.PSObject.Properties.Name -contains 'Confidence' -and $null -ne $ai.Confidence) { [System.Net.WebUtility]::HtmlEncode(("{0:P0}" -f [double]$ai.Confidence)) } else { "N/A" }
                $aiDataTimestamp = if ($ai.PSObject.Properties.Name -contains 'DataTimestamp' -and $ai.DataTimestamp) { [System.Net.WebUtility]::HtmlEncode("$($ai.DataTimestamp)") } else { "N/A" }
                $aiFallback = if ($ai.PSObject.Properties.Name -contains 'FallbackReason' -and $ai.FallbackReason) { [System.Net.WebUtility]::HtmlEncode("$($ai.FallbackReason)") } else { "" }
                $aiNarrative = if ($ai.PSObject.Properties.Name -contains 'Narrative' -and $ai.Narrative) {
                    [System.Net.WebUtility]::HtmlEncode("$($ai.Narrative)")
                } else {
                    ""
                }
                $aiFindings = @($ai.Findings)
                $aiRecommendations = @($ai.Recommendations)
                $findingRows = if (@($aiFindings).Count -gt 0) {
                    (@($aiFindings | ForEach-Object { "<li>$([System.Net.WebUtility]::HtmlEncode("$_"))</li>" }) -join "`n")
                } else {
                    "<li>None</li>"
                }
                $actionRows = if (@($aiRecommendations).Count -gt 0) {
                    (@($aiRecommendations | ForEach-Object { "<li>$([System.Net.WebUtility]::HtmlEncode("$_"))</li>" }) -join "`n")
                } else {
                    "<li>None</li>"
                }
                $metricRows = ""
                if ($ai.PSObject.Properties.Name -contains 'InputMetrics' -and $ai.InputMetrics) {
                    $metricRows = @(
                        $ai.InputMetrics.PSObject.Properties |
                        ForEach-Object {
                            $k = [System.Net.WebUtility]::HtmlEncode("$($_.Name)")
                            $v = if ($null -eq $_.Value -or "$($_.Value)" -eq "") { "N/A" } else { [System.Net.WebUtility]::HtmlEncode("$($_.Value)") }
                            "<tr><td>$k</td><td>$v</td></tr>"
                        }
                    ) -join "`n"
                }
                $narrativeBlock = if ($aiNarrative) { "<div class='ai-narrative'>$aiNarrative</div>" } else { "" }
                $summaryBlock = if ($aiSummary) { "<p><strong>サマリー:</strong> $aiSummary</p>" } else { "" }
                $fallbackBlock = if ($aiFallback) { "<p><strong>フォールバック理由:</strong> $aiFallback</p>" } else { "" }
                $metricBlock = if ($metricRows) { "<h3>入力指標一覧</h3><table><tr><th>指標</th><th>値</th></tr>$metricRows</table>" } else { "" }
                $aiSection = @"
    <section>
      <h2>AI診断</h2>
      <p><strong>要約:</strong> $aiHeadline</p>
      $summaryBlock
      <p><strong>評価:</strong> $aiEval <span class="ai-source">ソース: $aiSource</span></p>
      <p><strong>Prompt Version:</strong> $aiPromptVersion</p>
      <p><strong>信頼度:</strong> $aiConfidence</p>
      <p><strong>対象データ時刻:</strong> $aiDataTimestamp</p>
      $fallbackBlock
      $narrativeBlock
      <div class="ai-grid">
        <div>
          <h3>根拠</h3>
          <ul>$findingRows</ul>
        </div>
        <div>
          <h3>推奨アクション</h3>
          <ol>$actionRows</ol>
        </div>
      </div>
      $metricBlock
    </section>
"@
            }

            $summaryTable = $ReportData | ConvertTo-Html -Fragment
            $chartScriptTag = if ($UseLocalChartJs) {
                "<script src=""$([System.Net.WebUtility]::HtmlEncode($ChartJsScriptPath))""></script>"
            } else {
                '<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>'
            }
            $pcName   = if ($HostName) { [System.Net.WebUtility]::HtmlEncode($HostName) } else { [System.Net.WebUtility]::HtmlEncode($env:COMPUTERNAME) }
            $execDir  = if ($ExecFolder) { [System.Net.WebUtility]::HtmlEncode($ExecFolder) } else { [System.Net.WebUtility]::HtmlEncode($PSScriptRoot) }
            $genTime  = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            $html = @"
<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>PC Health Report</title>
  $chartScriptTag
  <style>
    body { font-family: "Segoe UI", sans-serif; background:#f5f7fb; margin:0; }
    .wrap { max-width: 1100px; margin: 24px auto; background:#fff; padding:24px; border-radius:10px; box-shadow:0 2px 12px rgba(0,0,0,.08);}
    .top { display:flex; justify-content:space-between; align-items:baseline; gap:12px; flex-wrap:wrap; }
    .score { font-size:28px; font-weight:700; }
    h1,h2 { margin:0 0 12px; }
    .grid { display:grid; grid-template-columns: 1fr 1fr; gap:20px; }
    @media (max-width: 860px) { .grid { grid-template-columns: 1fr; } }
    .ai-grid { display:grid; grid-template-columns: 1fr 1fr; gap:20px; }
    @media (max-width: 860px) { .ai-grid { grid-template-columns: 1fr; } }
    .ai-source { color:#64748b; font-size:12px; margin-left:10px; }
    .ai-narrative { white-space:pre-wrap; border:1px solid #ddd; border-radius:8px; padding:10px; background:#fafafa; }
    table { width:100%; border-collapse: collapse; }
    th, td { border:1px solid #ddd; padding:8px; text-align:left; }
    th { background:#f1f4f8; }
    .meta-bar { background:#f1f4f8; border-radius:6px; padding:8px 12px; font-size:13px; color:#475569; display:flex; gap:20px; flex-wrap:wrap; margin-bottom:16px; }
    .print-btn { background:#3b82f6; color:#fff; border:none; border-radius:6px; padding:8px 18px; font-size:14px; cursor:pointer; }
    .print-btn:hover { background:#2563eb; }
    @media print {
      .print-btn { display:none; }
      body { background:#fff; }
      .wrap { box-shadow:none; margin:0; padding:12px; }
    }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="top">
      <h1>PC Health Report</h1>
      <div class="score">Score: $score / 100</div>
    </div>
    <div class="meta-bar">
      <span>🖥️ PC名: $pcName</span>
      <span>📁 実行フォルダ: $execDir</span>
      <span>🕐 生成日時: $genTime</span>
      <button class="print-btn" onclick="window.print()">🖨️ 印刷/PDF保存</button>
    </div>
    <div class="grid">
      <section>
        <h2>Summary</h2>
        $summaryTable
      </section>
      <section>
        <h2>Resource Chart</h2>
        <canvas id="resourceChart" height="220"></canvas>
        <div id="resourceChartFallback" style="display:none;margin-top:10px;">
          <table>
            <tr><th>CPU</th><th>Memory</th><th>Disk</th></tr>
            <tr><td>$cpu</td><td>$memory</td><td>$disk</td></tr>
          </table>
        </div>
      </section>
    </div>
    $aiSection
    <section style="margin-top:24px;">
      <h2>スコア推移</h2>
      <canvas id="scoreTrendChart" height="120"></canvas>
    </section>
    <div class="grid" style="margin-top:24px;">
      <section>
        <h2>カテゴリ別レーダー</h2>
        <canvas id="categoryRadarChart"></canvas>
      </section>
    </div>
  </div>
  <script>
    const ctx = document.getElementById('resourceChart');
    if (typeof Chart !== 'undefined') {
      new Chart(ctx, {
        type: 'bar',
        data: {
          labels: ['CPU', 'Memory', 'Disk'],
          datasets: [{
            label: 'Score',
            data: [$cpu, $memory, $disk],
            backgroundColor: ['#3b82f6','#16a34a','#f59e0b']
          }]
        },
        options: { scales: { y: { beginAtZero: true, max: 100 } } }
      });
    } else {
      document.getElementById('resourceChart').style.display = 'none';
      document.getElementById('resourceChartFallback').style.display = 'block';
    }
    if (typeof Chart !== 'undefined') {
      new Chart(document.getElementById('scoreTrendChart'), {
        type: 'line',
        data: {
          labels: [$trendLabels],
          datasets: [{
            label: '総合スコア',
            data: [$trendScores],
            borderColor: '#3b82f6',
            backgroundColor: 'rgba(59,130,246,0.1)',
            fill: true,
            tension: 0.3
          }]
        },
        options: { scales: { y: { beginAtZero: true, max: 100 } } }
      });
      new Chart(document.getElementById('categoryRadarChart'), {
        type: 'radar',
        data: {
          labels: ['CPU','Memory','Disk','Security','Network','WindowsUpdate','Startup','SystemHealth'],
          datasets: [{
            label: 'スコア',
            data: [$cpu, $memory, $disk, $security, $network, $winUpdate, $startup, $sysHealth],
            backgroundColor: 'rgba(59,130,246,0.2)',
            borderColor: '#3b82f6',
            pointBackgroundColor: '#3b82f6'
          }]
        },
        options: { scales: { r: { beginAtZero: true, max: 100 } } }
      });
    }
  </script>
</body>
</html>
"@
            Set-Content -Path $Path -Value $html -Encoding $script:_enc
        }
    }

    return $Path
}

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER ReportData
Parameter description

.PARAMETER HistoryPath
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function Update-ScoreHistory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][pscustomobject]$ReportData,
        [string]$HistoryPath = ""
    )

    if ($HistoryPath -eq "") {
        $reportsDir = Join-Path (Split-Path $PSScriptRoot -Parent) 'reports'
        $HistoryPath = Join-Path $reportsDir 'score_history.json'
    }

    $dir = Split-Path $HistoryPath -Parent
    if (-not (Test-Path $dir)) {
        [void](New-Item -ItemType Directory -Path $dir -Force)
    }

    $history = @()
    if (Test-Path $HistoryPath) {
        try {
            $parsed = Get-Content -Path $HistoryPath -Raw | ConvertFrom-Json
            if ($null -ne $parsed) {
                $history = @($parsed)
            }
        } catch {
            $history = @()
        }
    }

    $entry = [PSCustomObject]@{
        runAt              = (Get-Date).ToString('yyyy-MM-dd HH:mm')
        score              = if ($ReportData.PSObject.Properties.Name -contains 'score') { [int]$ReportData.score } else { 0 }
        cpuScore           = if ($ReportData.PSObject.Properties.Name -contains 'cpuScore') { [int]$ReportData.cpuScore } else { 0 }
        memoryScore        = if ($ReportData.PSObject.Properties.Name -contains 'memoryScore') { [int]$ReportData.memoryScore } else { 0 }
        diskScore          = if ($ReportData.PSObject.Properties.Name -contains 'diskScore') { [int]$ReportData.diskScore } else { 0 }
        securityScore      = if ($ReportData.PSObject.Properties.Name -contains 'securityScore') { [int]$ReportData.securityScore } else { 0 }
        networkScore       = if ($ReportData.PSObject.Properties.Name -contains 'networkScore') { [int]$ReportData.networkScore } else { 0 }
        windowsUpdateScore = if ($ReportData.PSObject.Properties.Name -contains 'windowsUpdateScore') { [int]$ReportData.windowsUpdateScore } else { 0 }
        startupScore       = if ($ReportData.PSObject.Properties.Name -contains 'startupScore') { [int]$ReportData.startupScore } else { 0 }
        systemHealthScore  = if ($ReportData.PSObject.Properties.Name -contains 'systemHealthScore') { [int]$ReportData.systemHealthScore } else { 0 }
    }

    $history = @($history) + @($entry)
    if ($history.Count -gt 30) {
        $history = @($history | Select-Object -Last 30)
    }

    $history | ConvertTo-Json -Depth 5 | Set-Content -Path $HistoryPath -Encoding $script:_enc

    return @($history)
}

Export-ModuleMember -Function New-OptimizerReportData,Export-OptimizerReport,Update-ScoreHistory

