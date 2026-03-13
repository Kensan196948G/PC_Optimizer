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

Export-ModuleMember -Function New-OptimizerReportData,Export-OptimizerReport,Update-ScoreHistory,Export-AgentTeamsHtmlTimeline

# ============================================================
# Agent Teams HTML タイムライン生成
# ============================================================

function Export-AgentTeamsHtmlTimeline {
    <#
    .SYNOPSIS
        Agent Teams 実行結果をインタラクティブ HTML タイムラインにエクスポートします。
    .PARAMETER AgentTeamsResult
        Invoke-AgentTeamsOrchestration の戻り値オブジェクト
    .PARAMETER OutputPath
        出力 HTML ファイルパス (省略時は reports/agent-teams/ 配下に自動生成)
    .PARAMETER ReportsDir
        出力先ルートディレクトリ (OutputPath 省略時に使用)
    .OUTPUTS
        生成した HTML ファイルの絶対パス
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][pscustomobject]$AgentTeamsResult,
        [string]$OutputPath = "",
        [string]$ReportsDir = ""
    )

    $runId = if ($AgentTeamsResult.PSObject.Properties["transactionId"] -and $AgentTeamsResult.transactionId) {
        "$($AgentTeamsResult.transactionId)"
    } else { [guid]::NewGuid().ToString("N") }

    if (-not $OutputPath) {
        $baseDir = if ($ReportsDir) { $ReportsDir } else { Join-Path $PSScriptRoot "..\reports" }
        $atDir = Join-Path $baseDir "agent-teams"
        if (-not (Test-Path $atDir)) { New-Item -ItemType Directory -Path $atDir -Force | Out-Null }
        $OutputPath = Join-Path $atDir ("AgentTeamsTimeline_${runId}.html")
    }

    # 会話ログ・ノード結果の取得
    $convLog     = @()
    $nodeResults = @()
    if ($AgentTeamsResult.PSObject.Properties["conversationLog"]) { $convLog = @($AgentTeamsResult.conversationLog) }
    if ($AgentTeamsResult.PSObject.Properties["nodeResults"])     { $nodeResults = @($AgentTeamsResult.nodeResults) }

    # --- JS データ構造生成 ---
    # 全フィールドを安全にJSONエスケープするヘルパー
    # (message は ConvertTo-Json が生成する JSON オブジェクトから取得してから処理)
    $convJs = ($convLog | ForEach-Object {
        $safeMsg      = "$($_.message)".Replace("\", "\\").Replace('"', '\"').Replace("/", "\/").Replace("`n", '\n').Replace("`r", '\r')
        $safeNodeId   = "$($_.nodeId)".Replace("\", "\\").Replace('"', '\"')
        $safeRole     = "$($_.role)".Replace("\", "\\").Replace('"', '\"')
        $safeStatus   = "$($_.status)".Replace("\", "\\").Replace('"', '\"')
        $safeIcon     = "$($_.statusIcon)".Replace("\", "\\").Replace('"', '\"')
        $safeRisk     = "$($_.risk)".Replace("\", "\\").Replace('"', '\"')
        $safeTs       = if ($_.PSObject.Properties["timestamp"]) { "$($_.timestamp)".Replace('"', '\"') } else { "" }
        $lv           = if ($_.PSObject.Properties["level"]) { [int]$_.level } else { 0 }
        '{"level":{0},"nodeId":"{1}","role":"{2}","status":"{3}","statusIcon":"{4}","risk":"{5}","message":"{6}","durationMs":{7},"timestamp":"{8}"}' -f `
            $lv, $safeNodeId, $safeRole, $safeStatus, $safeIcon, $safeRisk, $safeMsg, [int]$_.durationMs, $safeTs
    }) -join ","

    $roleColorMap = @{
        "planner"    = "#a855f7"
        "collector"  = "#f59e0b"
        "analyzer"   = "#06b6d4"
        "remediator" = "#22c55e"
        "reporter"   = "#3b82f6"
    }

    $genTime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

    $htmlContent = @"
<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Agent Teams 会話タイムライン</title>
  <style>
    *{box-sizing:border-box;margin:0;padding:0}
    body{font-family:'Segoe UI',sans-serif;background:#0f172a;color:#e2e8f0;min-height:100vh}
    .header{background:linear-gradient(135deg,#1e293b,#0f172a);padding:24px 32px;border-bottom:1px solid #334155}
    .header h1{font-size:22px;font-weight:700;color:#f1f5f9}
    .header .meta{font-size:12px;color:#94a3b8;margin-top:6px}
    .container{max-width:1200px;margin:0 auto;padding:24px 32px}
    .section-title{font-size:16px;font-weight:600;color:#94a3b8;text-transform:uppercase;letter-spacing:.05em;margin:24px 0 12px}
    /* DAG Timeline */
    .dag-container{background:#1e293b;border-radius:12px;padding:20px;border:1px solid #334155}
    .level-row{display:flex;align-items:flex-start;margin-bottom:8px;gap:12px}
    .level-label{width:80px;text-align:right;font-size:11px;color:#64748b;padding-top:10px;flex-shrink:0}
    .level-nodes{display:flex;flex-wrap:wrap;gap:8px;flex:1}
    .node-card{border-radius:8px;padding:12px 16px;min-width:180px;position:relative;border:1px solid transparent;transition:transform .15s}
    .node-card:hover{transform:translateY(-2px)}
    .node-card .nc-header{display:flex;align-items:center;gap:8px;margin-bottom:6px}
    .node-card .nc-icon{font-size:14px;font-weight:700;background:rgba(255,255,255,.1);padding:2px 6px;border-radius:4px;font-family:monospace}
    .node-card .nc-id{font-size:13px;font-weight:600}
    .node-card .nc-msg{font-size:11px;color:rgba(255,255,255,.75);line-height:1.4;margin-top:4px}
    .node-card .nc-footer{display:flex;justify-content:space-between;align-items:center;margin-top:8px;font-size:10px;color:rgba(255,255,255,.5)}
    .node-card .nc-status{font-weight:700;padding:1px 6px;border-radius:3px}
    .status-ok{background:rgba(34,197,94,.2);color:#4ade80}
    .status-ng{background:rgba(239,68,68,.2);color:#f87171}
    .status-other{background:rgba(100,116,139,.2);color:#94a3b8}
    .risk-high{color:#f87171}
    .risk-medium{color:#fbbf24}
    .risk-low{color:#4ade80}
    .arrow-row{text-align:center;color:#475569;font-size:18px;margin:4px 0}
    /* Conversation Log */
    .conv-log{background:#1e293b;border-radius:12px;padding:20px;border:1px solid #334155}
    .conv-entry{display:flex;gap:12px;padding:10px 0;border-bottom:1px solid #1e3a5f}
    .conv-entry:last-child{border-bottom:none}
    .conv-bubble{flex:1;background:#0f172a;border-radius:8px;padding:10px 14px}
    .conv-bubble .cb-header{display:flex;align-items:center;gap:8px;margin-bottom:6px}
    .conv-bubble .cb-icon{font-size:11px;font-weight:700;padding:2px 6px;border-radius:4px;font-family:monospace}
    .conv-bubble .cb-from{font-size:12px;font-weight:600}
    .conv-bubble .cb-time{font-size:10px;color:#64748b;margin-left:auto}
    .conv-bubble .cb-msg{font-size:12px;color:#cbd5e1;line-height:1.5}
    .conv-bubble .cb-dur{font-size:10px;color:#475569;margin-top:4px}
    .badge-level{background:#1e3a5f;color:#38bdf8;font-size:10px;padding:2px 8px;border-radius:9999px;flex-shrink:0;align-self:center}
    /* Summary bar */
    .summary-bar{display:flex;gap:20px;background:#1e293b;border-radius:10px;padding:16px 20px;border:1px solid #334155;flex-wrap:wrap;margin-bottom:24px}
    .summary-item{text-align:center}
    .summary-item .si-val{font-size:24px;font-weight:700;color:#f1f5f9}
    .summary-item .si-label{font-size:11px;color:#64748b;margin-top:2px}
    .si-success .si-val{color:#4ade80}
    .si-failed .si-val{color:#f87171}
  </style>
</head>
<body>
  <div class="header">
    <h1>🤖 Agent Teams 会話タイムライン</h1>
    <div class="meta">RunId: ${runId} &nbsp;|&nbsp; 生成日時: ${genTime}</div>
  </div>
  <div class="container">
    <div id="summary-bar" class="summary-bar"></div>
    <div class="section-title">DAG 実行フロー</div>
    <div class="dag-container" id="dag-container"></div>
    <div class="section-title" style="margin-top:32px">エージェント間会話ログ</div>
    <div class="conv-log" id="conv-log"></div>
  </div>
  <script>
    const convData = [$convJs];
    const roleColors = {
      planner: '#a855f7', collector: '#f59e0b', analyzer: '#06b6d4',
      remediator: '#22c55e', reporter: '#3b82f6'
    };
    const roleIcons = {
      planner: '[PLAN]', collector: '[COL]', analyzer: '[ANL]',
      remediator: '[REM]', reporter: '[RPT]'
    };
    function getRoleColor(r){ return roleColors[r] || '#94a3b8'; }
    function getRoleIcon(r){ return roleIcons[r] || '[AGT]'; }
    function statusClass(s){ return s==='Success'?'status-ok':s==='Failed'?'status-ng':'status-other'; }
    function riskClass(r){ return r==='High'?'risk-high':r==='Medium'?'risk-medium':'risk-low'; }
    function escHtml(s){ return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;'); }

    // Build DAG timeline by level
    const byLevel = {};
    convData.forEach(e => {
      if (!byLevel[e.level]) byLevel[e.level] = [];
      byLevel[e.level].push(e);
    });
    const dagContainer = document.getElementById('dag-container');
    const levels = Object.keys(byLevel).map(Number).sort((a,b)=>a-b);
    levels.forEach((lv, idx) => {
      const row = document.createElement('div');
      row.className = 'level-row';
      const label = document.createElement('div');
      label.className = 'level-label';
      label.textContent = 'Level ' + lv;
      row.appendChild(label);
      const nodes = document.createElement('div');
      nodes.className = 'level-nodes';
      byLevel[lv].forEach(e => {
        const c = document.createElement('div');
        c.className = 'node-card';
        c.style.background = getRoleColor(e.role) + '22';
        c.style.borderColor = getRoleColor(e.role) + '55';
        const shortMsg = e.message.length > 60 ? e.message.substring(0,60)+'...' : e.message;
        c.innerHTML = `
          <div class="nc-header">
            <span class="nc-icon" style="color:${getRoleColor(e.role)}">${escHtml(getRoleIcon(e.role))}</span>
            <span class="nc-id">${escHtml(e.nodeId)}</span>
          </div>
          <div class="nc-msg">${escHtml(shortMsg)}</div>
          <div class="nc-footer">
            <span class="nc-status ${statusClass(e.status)}">${escHtml(e.statusIcon)} ${escHtml(e.status)}</span>
            <span class="${riskClass(e.risk)}">${escHtml(e.risk)}</span>
            <span>${e.durationMs}ms</span>
          </div>`;
        nodes.appendChild(c);
      });
      row.appendChild(nodes);
      dagContainer.appendChild(row);
      if (idx < levels.length - 1) {
        const arr = document.createElement('div');
        arr.className = 'arrow-row';
        arr.innerHTML = '&#9660;';
        dagContainer.appendChild(arr);
      }
    });

    // Build conversation log
    const convLogEl = document.getElementById('conv-log');
    convData.forEach(e => {
      const entry = document.createElement('div');
      entry.className = 'conv-entry';
      const badge = document.createElement('div');
      badge.className = 'badge-level';
      badge.textContent = 'L' + e.level;
      const bubble = document.createElement('div');
      bubble.className = 'conv-bubble';
      bubble.style.borderLeft = '3px solid ' + getRoleColor(e.role);
      bubble.innerHTML = `
        <div class="cb-header">
          <span class="cb-icon" style="background:${getRoleColor(e.role)}33;color:${getRoleColor(e.role)}">${escHtml(getRoleIcon(e.role))}</span>
          <span class="cb-from" style="color:${getRoleColor(e.role)}">${escHtml(e.nodeId)}</span>
          <span class="cb-time">${escHtml(e.timestamp || '')}</span>
        </div>
        <div class="cb-msg">${escHtml(e.message)}</div>
        <div class="cb-dur">所要: ${e.durationMs}ms &nbsp;|&nbsp; ステータス: <span class="${statusClass(e.status)}">${escHtml(e.statusIcon)} ${escHtml(e.status)}</span> &nbsp;|&nbsp; リスク: <span class="${riskClass(e.risk)}">${escHtml(e.risk)}</span></div>`;
      entry.appendChild(badge);
      entry.appendChild(bubble);
      convLogEl.appendChild(entry);
    });

    // Summary bar
    const total = convData.length;
    const success = convData.filter(e=>e.status==='Success').length;
    const failed = total - success;
    const totalMs = convData.reduce((a,e)=>a+e.durationMs,0);
    document.getElementById('summary-bar').innerHTML = `
      <div class="summary-item si-success"><div class="si-val">${success}</div><div class="si-label">成功ノード</div></div>
      <div class="summary-item si-failed"><div class="si-val">${failed}</div><div class="si-label">失敗ノード</div></div>
      <div class="summary-item"><div class="si-val">${total}</div><div class="si-label">総ノード数</div></div>
      <div class="summary-item"><div class="si-val">${(totalMs/1000).toFixed(2)}s</div><div class="si-label">累計実行時間</div></div>
    `;
  </script>
</body>
</html>
"@
    Set-Content -Path $OutputPath -Value $htmlContent -Encoding $script:_enc
    return (Resolve-Path $OutputPath).Path
}

