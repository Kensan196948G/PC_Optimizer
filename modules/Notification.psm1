#Requires -Version 5.1
Set-StrictMode -Version Latest
$script:_enc = if ($PSVersionTable.PSVersion.Major -ge 7) { 'utf8NoBOM' } else { 'UTF8' }

function Send-SlackNotification {
    param(
        [Parameter(Mandatory)][string]$WebhookUrl,
        [Parameter(Mandatory)][string]$HostName,
        [Parameter(Mandatory)][int]$Score,
        [string]$Evaluation = "不明",
        [string]$TopRecommendation = "",
        [int]$TimeoutSec = 15
    )

    if ([string]::IsNullOrEmpty($WebhookUrl)) {
        Write-Warning "[Slack] WebhookUrl が未設定のため通知をスキップします。"
        return
    }

    $color = if ($Score -ge 80) { "#36a64f" } elseif ($Score -ge 50) { "#ffcc00" } else { "#ff0000" }

    $payload = [ordered]@{
        text        = "PC Health Report - Score: $Score/100"
        attachments = @(
            [ordered]@{
                color  = $color
                fields = @(
                    [ordered]@{ title = "ホスト名";       value = $HostName;         short = $true }
                    [ordered]@{ title = "評価";           value = $Evaluation;       short = $true }
                    [ordered]@{ title = "推奨アクション"; value = $TopRecommendation; short = $false }
                )
            }
        )
    }

    $json = $payload | ConvertTo-Json -Depth 10 -Compress
    $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($json)

    try {
        if ($PSVersionTable.PSVersion.Major -ge 7) {
            Invoke-RestMethod -Uri $WebhookUrl -Method Post -Body $bodyBytes -ContentType "application/json" -TimeoutSec $TimeoutSec | Out-Null
        } else {
            $wresp = Invoke-WebRequest -Uri $WebhookUrl -Method Post -Body $bodyBytes -ContentType "application/json" -TimeoutSec $TimeoutSec
        }
    } catch {
        Write-Warning "[Slack] 通知の送信に失敗しました: $_"
    }
}

function Send-TeamsNotification {
    param(
        [Parameter(Mandatory)][string]$WebhookUrl,
        [Parameter(Mandatory)][string]$HostName,
        [Parameter(Mandatory)][int]$Score,
        [string]$Evaluation = "不明",
        [string]$TopRecommendation = "",
        [int]$TimeoutSec = 15
    )

    if ([string]::IsNullOrEmpty($WebhookUrl)) {
        Write-Warning "[Teams] WebhookUrl が未設定のため通知をスキップします。"
        return
    }

    $themeColor = if ($Score -ge 80) { "00B050" } elseif ($Score -ge 50) { "FFC000" } else { "FF0000" }

    $payload = [ordered]@{
        "@type"      = "MessageCard"
        "@context"   = "http://schema.org/extensions"
        themeColor   = $themeColor
        summary      = "PC Health Report"
        sections     = @(
            [ordered]@{
                activityTitle    = "PC Health Report - Score: $Score/100"
                activitySubtitle = "ホスト名: $HostName"
                facts            = @(
                    [ordered]@{ name = "評価";           value = $Evaluation }
                    [ordered]@{ name = "推奨アクション"; value = $TopRecommendation }
                )
            }
        )
    }

    $json = $payload | ConvertTo-Json -Depth 10 -Compress
    $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($json)

    try {
        if ($PSVersionTable.PSVersion.Major -ge 7) {
            Invoke-RestMethod -Uri $WebhookUrl -Method Post -Body $bodyBytes -ContentType "application/json" -TimeoutSec $TimeoutSec | Out-Null
        } else {
            $wresp = Invoke-WebRequest -Uri $WebhookUrl -Method Post -Body $bodyBytes -ContentType "application/json" -TimeoutSec $TimeoutSec
        }
    } catch {
        Write-Warning "[Teams] 通知の送信に失敗しました: $_"
    }
}

function Send-McpProviderNotification {
    param(
        [Parameter(Mandatory)][object[]]$McpProviders,
        [Parameter(Mandatory)][string]$HostName,
        [Parameter(Mandatory)][int]$Score,
        [string]$Evaluation = "不明",
        [string]$TopRecommendation = "",
        [int]$TimeoutSec = 15
    )

    $commonParams = @{
        HostName          = $HostName
        Score             = $Score
        Evaluation        = $Evaluation
        TopRecommendation = $TopRecommendation
        TimeoutSec        = $TimeoutSec
    }

    foreach ($provider in $McpProviders) {
        if (-not $provider.enabled) { continue }

        $type = [string]$provider.type
        $url  = [string]$provider.webhookUrl

        try {
            switch ($type.ToLower()) {
                'slack' {
                    Send-SlackNotification -WebhookUrl $url @commonParams
                }
                'teams' {
                    Send-TeamsNotification -WebhookUrl $url @commonParams
                }
                default {
                    Write-Warning "[McpProvider] 未知のプロバイダータイプ '$type' をスキップします。"
                }
            }
        } catch {
            Write-Warning "[McpProvider] プロバイダー '$type' の通知処理中にエラーが発生しました: $_"
        }
    }
}

Export-ModuleMember -Function Send-SlackNotification, Send-TeamsNotification, Send-McpProviderNotification
