# PowerShell バージョン互換性

## 対応バージョン一覧

| PowerShell | 状態 | 備考 |
|---|---|---|
| 2.0 | 非対応 | `Get-CimInstance` が存在しない |
| 3.0 | 警告あり・動作可 | 最小サポートバージョン |
| 4.0 | 警告あり・動作可 | Windows 8.1 標準搭載 |
| 5.0 | 完全対応 | |
| **5.1** | **完全対応（推奨）** | Windows 10/11 標準搭載 |
| 6.x | 完全対応 | `Get-WmiObject` 削除済みだが本ツールは `Get-CimInstance` を使用 |
| **7.x** | **完全対応（推奨）** | 絵文字・ANSI カラー・Unicode ブロック進捗バーが有効 |

## バージョン別の機能差分

### 表示スタイル

| 機能 | PS 5.1 以下 | PS 7.x 以降 |
|---|---|---|
| システム情報のアイコン | `*` ASCII 記号 | 🖥️ 👤 🏷️ 💻 🧠 💾 絵文字 |
| 進捗バー文字 | `#---` ハッシュ記号 | `██░░` Unicode ブロック |
| 完了メッセージ | テキストのみ | ✅ 付き |

### コマンドの可用性

| コマンド / API | PS 3.0 | PS 5.1 | PS 7.x |
|---|---|---|---|
| `Get-CimInstance` | ✅ | ✅ | ✅ |
| `Get-WmiObject` | ✅ | ✅ | ❌ 削除済み |
| `Get-PhysicalDisk` | ✅* | ✅ | ✅ |
| `Optimize-Volume` | ✅* | ✅ | ✅ |
| `Add-Content` | ✅ | ✅ | ✅ |
| `Start-Process` | ✅ | ✅ | ✅ |
| `Read-Host` | ✅ | ✅ | ✅ |
| `wevtutil.exe` | ✅ | ✅ | ✅ |
| `sfc.exe` | ✅ | ✅ | ✅ |
| `DISM.exe` | ✅ | ✅ | ✅ |
| `powercfg.exe` | ✅ | ✅ | ✅ |

*Storage モジュールが必要（Windows 8.1+ に標準搭載）

### `Get-PhysicalDisk` の可用性チェック

本ツールは `Get-PhysicalDisk` を使用する前に以下のチェックを行います。

```powershell
if (Get-Command -Name Get-PhysicalDisk -ErrorAction SilentlyContinue) {
    # SSD/HDD 判定・SMART チェック
} else {
    # スキップしてログに記録
}
```

Storage モジュールが利用できない環境ではスキップされ、処理は継続します。

### `UsoClient.exe` の可用性チェック

```powershell
$usoPath = Join-Path $env:SystemRoot "System32\UsoClient.exe"
if (-not (Test-Path $usoPath)) {
    # Windows Update をスキップ
}
```

`UsoClient.exe` は Windows 10 以降にのみ存在します。

## 各バージョンの取得方法

### PowerShell 5.1 の確認（Windows 標準）

```powershell
$PSVersionTable.PSVersion
# 期待出力: Major:5, Minor:1
```

### PowerShell 7 のインストール

```powershell
# winget 経由（Windows 10/11）
winget install Microsoft.PowerShell

# または公式サイトからインストーラーをダウンロード
# https://github.com/PowerShell/PowerShell/releases
```

### インストール済みバージョンの確認

```powershell
# PowerShell 5.1 のパス
Get-Command powershell | Select-Object -ExpandProperty Source
# C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe

# PowerShell 7 のパス（インストール済みの場合）
Get-Command pwsh | Select-Object -ExpandProperty Source
# C:\Program Files\PowerShell\7\pwsh.exe
```

## 実行時のバージョン表示

スクリプト実行時にログに PowerShell バージョンが記録されます。

```
[PowerShell バージョン] 7.4.1
```

コンソールでも確認できます：

```powershell
$PSVersionTable.PSVersion.ToString()
```

## HTTP レスポンスの UTF-8 デコード

v4.0.1 で修正された PS5.1 の文字化けバグへの対処パターンです。

### 問題

PowerShell 5.1 の `Invoke-RestMethod` は HTTP レスポンスをシステムデフォルトエンコーディング（日本語 Windows では CP932/Shift-JIS）で解釈します。Anthropic API などの UTF-8 JSON レスポンスを受信すると日本語部分が文字化けします。

### 解決パターン（modules/Advanced.psm1・Notification.psm1 採用）

```powershell
if ($PSVersionTable.PSVersion.Major -ge 7) {
    # PS7: Invoke-RestMethod は UTF-8 を正しく処理する
    $response = Invoke-RestMethod -Uri $uri -Method Post -Body $bodyBytes -ContentType "application/json; charset=utf-8"
} else {
    # PS5.1: Invoke-WebRequest + 明示的 UTF-8 デコード
    $raw = Invoke-WebRequest -Uri $uri -Method Post -Body $bodyBytes -ContentType "application/json; charset=utf-8"
    $json = [System.Text.Encoding]::UTF8.GetString($raw.RawContentStream.ToArray())
    $response = $json | ConvertFrom-Json
}
```

| PowerShell | 推奨 API | 理由 |
|---|---|---|
| 5.1 | `Invoke-WebRequest` + `RawContentStream` + `UTF8.GetString` | `Invoke-RestMethod` が CP932 で誤デコード |
| 7.x | `Invoke-RestMethod` | UTF-8 を正しく処理する |

## エンコーディング統一パターン（$script:_enc）

v4.0.1 でモジュール全体に適用されたファイル書き込みエンコーディング統一パターンです。

```powershell
# モジュール冒頭に定義（関数定義より前に記述）
$script:_enc = if ($PSVersionTable.PSVersion.Major -ge 7) { 'utf8NoBOM' } else { 'UTF8' }

# 使用例（全ての Out-File / Set-Content / Add-Content で指定）
Set-Content -Path $path -Value $content -Encoding $script:_enc
Add-Content -Path $path -Value $line -Encoding $script:_enc
Out-File -FilePath $path -Encoding $script:_enc -Append
```

| 設定 | PS 5.1 値 | PS 7.x 値 |
|---|---|---|
| `$script:_enc` | `UTF8`（BOM付き） | `utf8NoBOM`（BOM無し） |

## StrictMode 対応パターン

`Set-StrictMode -Version Latest` 環境で PSObject プロパティにアクセスする際の安全なパターン（v4.0.1 で適用）。

```powershell
# NG: StrictMode で PropertyNotFoundException が発生する
$retryCount = if ($provider.retryCount) { $provider.retryCount } else { 3 }

# OK: PSObject.Properties ガードを使用
$retryCount = if ($provider.PSObject.Properties["retryCount"] -and $provider.retryCount) {
    [int]$provider.retryCount
} else { 3 }
```

## Notification.psm1 / Advanced.psm1 の外部 API 呼び出し互換性

| 機能 | 関数 | PS 5.1 実装 | PS 7.x 実装 |
|---|---|---|---|
| AI 診断 | `Invoke-AIDiagnosis` | `Invoke-WebRequest` + `RawContentStream` UTF-8 デコード | `Invoke-RestMethod` |
| Slack 通知 | `Send-SlackNotification` | `Invoke-WebRequest` | `Invoke-RestMethod` |
| Teams 通知 | `Send-TeamsNotification` | `Invoke-WebRequest` | `Invoke-RestMethod` |
| ServiceNow 起票 | `Send-ServiceNowIncident` | `Invoke-WebRequest` + `RawContentStream` UTF-8 デコード | `Invoke-RestMethod` |
| Jira 起票 | `Send-JiraTask` | `Invoke-WebRequest` + `RawContentStream` UTF-8 デコード | `Invoke-RestMethod` |

## コマンド可用性テーブル（追補）

以下のコマンドが v4.0 以降で追加利用されています。

| コマンド / API | PS 5.1 | PS 7.x | 備考 |
|---|---|---|---|
| `Invoke-RestMethod` | ✅（UTF-8 注意） | ✅ | 外部 API 呼び出し。PS5.1 は UTF-8 デコード要注意 |
| `Invoke-WebRequest` | ✅ | ✅ | PS5.1 での外部 API 呼び出しに使用 |
| `ConvertTo-Json -Depth` | ✅ | ✅ | `-Compress` フラグは PS3.0+ |
| `[System.Text.Encoding]::UTF8.GetString` | ✅ | ✅ | PS5.1 HTTP レスポンスデコードに使用 |
| `[Convert]::ToBase64String` | ✅ | ✅ | Jira Basic 認証 |
