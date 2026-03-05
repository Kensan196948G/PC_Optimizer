#Requires -Version 5.1
Set-StrictMode -Version Latest
$script:_enc = if ($PSVersionTable.PSVersion.Major -ge 7) { 'utf8NoBOM' } else { 'UTF8' }

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER WebhookUrl
Parameter description

.PARAMETER HostName
Parameter description

.PARAMETER Score
Parameter description

.PARAMETER Evaluation
Parameter description

.PARAMETER TopRecommendation
Parameter description

.PARAMETER TimeoutSec
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function Send-SlackNotification {
    param(
        [string]$WebhookUrl = "",
        [string]$HostName = "",
        [int]$Score = 0,
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

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER WebhookUrl
Parameter description

.PARAMETER HostName
Parameter description

.PARAMETER Score
Parameter description

.PARAMETER Evaluation
Parameter description

.PARAMETER TopRecommendation
Parameter description

.PARAMETER TimeoutSec
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function Send-TeamsNotification {
    param(
        [string]$WebhookUrl = "",
        [string]$HostName = "",
        [int]$Score = 0,
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

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER McpProviders
Parameter description

.PARAMETER HostName
Parameter description

.PARAMETER Score
Parameter description

.PARAMETER Evaluation
Parameter description

.PARAMETER TopRecommendation
Parameter description

.PARAMETER TimeoutSec
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
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

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER InstanceUrl
Parameter description

.PARAMETER Table
Parameter description

.PARAMETER Token
Parameter description

.PARAMETER HostName
Parameter description

.PARAMETER Score
Parameter description

.PARAMETER Summary
Parameter description

.PARAMETER TopRecommendation
Parameter description

.PARAMETER TimeoutSec
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function Send-ServiceNowIncident {
    param(
        [string]$InstanceUrl = "",
        [string]$Table = "incident",
        [string]$Token = "",
        [Parameter(Mandatory)][string]$HostName,
        [Parameter(Mandatory)][int]$Score,
        [string]$Summary = "",
        [string]$TopRecommendation = "",
        [int]$TimeoutSec = 15
    )

    if ([string]::IsNullOrEmpty($InstanceUrl)) {
        Write-Warning "[ServiceNow] InstanceUrl が未設定のため起票をスキップします。"
        return $null
    }

    $urgency = if ($Score -lt 50) { "1" } elseif ($Score -lt 70) { "2" } else { "3" }
    $body = @{
        short_description = "PC Health Alert: $HostName Score=$Score/100"
        description       = "PC最適化ツールによる自動起票`nホスト名: $HostName`nスコア: $Score/100`n評価: $Summary`n推奨アクション: $TopRecommendation"
        urgency           = $urgency
    } | ConvertTo-Json

    $headers = @{ "Content-Type" = "application/json" }
    if (-not [string]::IsNullOrEmpty($Token)) { $headers["Authorization"] = "Bearer $Token" }
    $url = ("{0}/api/now/table/{1}" -f $InstanceUrl.TrimEnd("/"), $Table)

    try {
        if ($PSVersionTable.PSVersion.Major -ge 7) {
            $resp = Invoke-RestMethod -Uri $url -Method Post -Headers $headers -Body $body -ContentType "application/json" -TimeoutSec $TimeoutSec
        } else {
            $raw = Invoke-WebRequest -Uri $url -Method Post -Headers $headers -Body $body -ContentType "application/json" -TimeoutSec $TimeoutSec
            $resp = [System.Text.Encoding]::UTF8.GetString($raw.RawContentStream.ToArray()) | ConvertFrom-Json
        }
        $sysId = if ($resp -and $resp.result -and $resp.result.sys_id) { "$($resp.result.sys_id)" } else { $null }
        return [PSCustomObject]@{ status = "Created"; sysId = $sysId; url = $url }
    } catch {
        Write-Warning "[ServiceNow] インシデント起票に失敗しました: $_"
        return [PSCustomObject]@{ status = "Failed"; sysId = $null; url = $url; error = "$_" }
    }
}

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER JiraUrl
Parameter description

.PARAMETER ProjectKey
Parameter description

.PARAMETER IssueType
Parameter description

.PARAMETER UserEmail
Parameter description

.PARAMETER ApiToken
Parameter description

.PARAMETER HostName
Parameter description

.PARAMETER Score
Parameter description

.PARAMETER Summary
Parameter description

.PARAMETER TopRecommendation
Parameter description

.PARAMETER TimeoutSec
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function Send-JiraTask {
    param(
        [string]$JiraUrl = "",
        [string]$ProjectKey = "",
        [string]$IssueType = "Task",
        [string]$UserEmail = "",
        [string]$ApiToken = "",
        [Parameter(Mandatory)][string]$HostName,
        [Parameter(Mandatory)][int]$Score,
        [string]$Summary = "",
        [string]$TopRecommendation = "",
        [int]$TimeoutSec = 15
    )

    if ([string]::IsNullOrEmpty($JiraUrl) -or [string]::IsNullOrEmpty($ProjectKey)) {
        Write-Warning "[Jira] JiraUrl または ProjectKey が未設定のため起票をスキップします。"
        return $null
    }

    $body = @{
        fields = @{
            project     = @{ key = $ProjectKey }
            summary     = "PC Health Alert: $HostName Score=$Score/100"
            description = "PC最適化ツールによる自動起票`nホスト名: $HostName`nスコア: $Score/100`n評価: $Summary`n推奨アクション: $TopRecommendation"
            issuetype   = @{ name = if ($IssueType) { $IssueType } else { "Task" } }
        }
    } | ConvertTo-Json -Depth 8

    $headers = @{}
    if (-not [string]::IsNullOrEmpty($UserEmail) -and -not [string]::IsNullOrEmpty($ApiToken)) {
        $pair = "{0}:{1}" -f $UserEmail, $ApiToken
        $headers["Authorization"] = "Basic $([Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($pair)))"
    }
    $url = ("{0}/rest/api/2/issue" -f $JiraUrl.TrimEnd("/"))

    try {
        if ($PSVersionTable.PSVersion.Major -ge 7) {
            $resp = Invoke-RestMethod -Uri $url -Method Post -Headers $headers -Body $body -ContentType "application/json" -TimeoutSec $TimeoutSec
        } else {
            $raw = Invoke-WebRequest -Uri $url -Method Post -Headers $headers -Body $body -ContentType "application/json" -TimeoutSec $TimeoutSec
            $resp = [System.Text.Encoding]::UTF8.GetString($raw.RawContentStream.ToArray()) | ConvertFrom-Json
        }
        $key = if ($resp -and $resp.key) { "$($resp.key)" } else { $null }
        return [PSCustomObject]@{ status = "Created"; issueKey = $key; url = $url }
    } catch {
        Write-Warning "[Jira] タスク起票に失敗しました: $_"
        return [PSCustomObject]@{ status = "Failed"; issueKey = $null; url = $url; error = "$_" }
    }
}

Export-ModuleMember -Function Send-SlackNotification, Send-TeamsNotification, Send-McpProviderNotification, Send-ServiceNowIncident, Send-JiraTask
