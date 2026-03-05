# config仕様

対象ファイル: `config/config.json`

## 主要キー

| キー | 型 | 既定値 | 説明 |
|---|---|---|---|
| `logLevel` | string | `INFO` | ログ出力レベル。現時点は将来拡張用の設定値として保持。 |
| `reportFormat` | string | `HTML` | レポート形式の既定値。 |
| `cleanupBrowserCache` | bool | `true` | ブラウザキャッシュ関連処理の有効/無効フラグ。 |
| `diskOptimize` | bool | `true` | ディスク最適化処理の有効/無効フラグ。 |
| `enableSecurityCheck` | bool | `true` | セキュリティ診断処理の有効/無効フラグ。 |
| `failureMode` | string | `continue` | 失敗時動作の既定値（`continue` / `fail-fast`）。 |
| `defaultTasks` | string | `all` | 実行対象タスクの既定指定（`all` または `1,3,7` 形式）。 |

## WhatIf 削除候補マップ

`whatIfDeletedPathMap` は、`-WhatIf` 実行時に削除候補として出力するパスの定義です。

- キー: タスク番号（文字列）
- 値: パス配列（環境変数 `%TEMP%` 形式を利用可）

例:

```json
{
  "whatIfDeletedPathMap": {
    "1": [
      "%SystemRoot%\\Temp",
      "%TEMP%"
    ],
    "7": [
      "%LOCALAPPDATA%\\Google\\Chrome\\User Data\\Default\\Cache"
    ]
  }
}
```

## オプション機能の設定（ネットワーク通知）

以下の機能は `enabled: false`（デフォルト）のため**無効時はエラーなし・スキップ**されます。
利用する場合のみ値を設定し `enabled: true` に変更してください。

---

### Slack 通知（オプション）

PC 最適化完了後に Slack チャンネルへスコア・評価・推奨アクションを POST します。

```json
{
  "name": "slack-notify",
  "type": "slack",
  "enabled": true,
  "webhookUrl": "https://hooks.slack.com/services/XXXX/YYYY/ZZZZ"
}
```

取得方法: Slack App → Incoming Webhooks → URL をコピー

---

### Teams 通知（オプション）

PC 最適化完了後に Teams チャンネルへ MessageCard を POST します。

```json
{
  "name": "teams-notify",
  "type": "teams",
  "enabled": true,
  "webhookUrl": "https://xxxx.webhook.office.com/webhookb2/..."
}
```

取得方法: Teams チャンネル → コネクタ → Incoming Webhook

---

### ServiceNow インシデント自動起票（オプション）

PC スコアが `scoreThreshold`（既定値 70）未満の場合にインシデントを自動起票します。

```json
{
  "name": "servicenow-ticket",
  "type": "servicenow",
  "enabled": true,
  "instanceUrl": "https://yourinstance.service-now.com",
  "table": "incident",
  "token": "Bearer_トークン文字列",
  "scoreThreshold": 70
}
```

---

### Jira タスク自動起票（オプション）

PC スコアが `scoreThreshold` 未満の場合に Jira タスクを自動起票します。

```json
{
  "name": "jira-ticket",
  "type": "jira",
  "enabled": true,
  "url": "https://yourteam.atlassian.net",
  "projectKey": "OPS",
  "issueType": "Task",
  "userEmail": "your@email.com",
  "apiToken": "Jira_APIトークン"
}
```

`apiToken` 取得方法: Atlassian アカウント → セキュリティ → API トークンを作成

---

### SIEM 出力（オプション）

Hook イベントを JSONL / CEF / LEEF 形式でローカルファイルに書き出します。  
SIEM ツールや syslog フォワーダーと連携する際に有効化してください。

```json
"siem": {
  "enabled": true,
  "formats": [ "jsonl", "cef", "leef" ],
  "outputDir": "siem"
}
```

出力先: `logs/siem/HookEvents_{RunId}_{date}.{format}`

---

> **共通事項**: 3 機能はいずれも**ネットワーク接続済みの PC 上でのみ動作**します。
> 無効（`enabled: false`）の場合はスキップされ、最適化処理には一切影響しません。


