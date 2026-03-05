# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## プロジェクト概要

Windows 10/11 向けの PC 最適化ツール。管理者権限で実行し、一時ファイル削除・ブラウザキャッシュ・ディスク最適化・SFC/DISM・Microsoft 365 / Windows Update（確認後適用）など 20 タスクを自動実行する。

## 前提

- **スタンドアロンローカルツール**: `D:\PC_Optimizer` フォルダを対象 PC の任意フォルダにコピーして、バッチを管理者権限で実行する形式。インストール不要。
- **対象 PC は常時ネット接続済み**: Slack / Teams / ServiceNow / Jira 等への外部 API POST 設計は有効（デフォルト無効、オプションで有効化）。
- **WinRM / 遠隔操作は設計外**: ローカル実行専用。リモート管理には対応しない。

## 実行方法

```batch
# 管理者として実行（BAT ファイル経由が推奨）
Run_PC_Optimizer.bat

# PowerShell から直接実行
powershell -NoProfile -ExecutionPolicy Bypass -File "PC_Optimizer.ps1"
```

BAT ファイルは PowerShell 7 (pwsh) を優先検出し、なければ PowerShell 5.1 にフォールバックする。管理者権限がない場合は UAC で昇格要求する。

テストは `tests/` フォルダ内のスクリプトを使用する:

```powershell
# 文字コード・構文チェック
powershell -NoProfile -ExecutionPolicy Bypass -File "tests\Test_Encoding.ps1"

# メインロジックテスト（93+ テスト）
powershell -NoProfile -ExecutionPolicy Bypass -File "tests\Test_PCOptimizer.ps1"

# Pester テスト（カバレッジ 30%）
Invoke-Pester "tests\PCOptimizer.Pester.Tests.ps1"
```

## アーキテクチャ

### ファイル構成

| ファイル / フォルダ | 役割 |
|---|---|
| `PC_Optimizer.ps1` | メイン最適化エンジン |
| `Run_PC_Optimizer.bat` | 管理者昇格 + PowerShell 検出ランチャー（Shift-JIS/CRLF） |
| `modules/` | 機能別 PowerShell モジュール群（下記参照） |
| `config/` | 設定ファイル群（JSON） |
| `reports/` | HTML/CSV/JSON レポート出力先（自動生成） |
| `logs/` | 実行ログ・SIEM ログの出力先（自動生成） |
| `tests/` | テストスクリプト群 |
| `docs/` | ドキュメント群 |

### モジュール構成（modules/）

| モジュール | 役割 |
|---|---|
| `Common.psm1` | `$script:_enc` 定義・`Write-StructuredLog`・`Invoke-GuardedStep` |
| `Cleanup.psm1` | ファイル削除系タスク |
| `Diagnostics.psm1` | SFC / DISM / SSD 診断 |
| `Performance.psm1` | ディスク最適化・電源プラン |
| `Security.psm1` | セキュリティ診断 |
| `Network.psm1` | DNS / ネットワーク最適化 |
| `Update.psm1` | Windows Update / Microsoft 365 更新 |
| `Report.psm1` | HTML/CSV/JSON レポート生成・`Update-ScoreHistory`・Chart.js グラフ |
| `Advanced.psm1` | AI 診断エンジン（Anthropic API、PS5.1 UTF-8 修正済み） |
| `Orchestration.psm1` | Agent Teams DAG 実行・Hook ディスパッチ・MCP 管理・SIEM 出力 |
| `Notification.psm1` | Slack / Teams / ServiceNow / Jira 通知（全てデフォルト `enabled:false`） |
| `agents/` | カスタムエージェント格納先 |

### 実行フロー

```
Run_PC_Optimizer.bat
  → PowerShell 検出 & 管理者昇格
    → PC_Optimizer.ps1
      → モジュールロード（modules/*.psm1）
      → システム情報収集・ログ初期化（logs/ フォルダに出力）
      → 初期ディスク空き容量記録
      → 20 タスクを Invoke-GuardedStep で順次実行
      → ディスク解放量の比較表示
      → Report.psm1 でレポート生成（reports/ フォルダ）
      → Agent Teams DAG 実行（Orchestration.psm1）
      → 再起動要否チェック → 終了
```

#### Agent Teams DAG（Orchestration.psm1）

```
planner
  → [collector.security, collector.network, collector.update]  ← 並行実行
    → analyzer.*
      → analyzer.aggregate
        → remediator
          → reporter
            → on_report フック自動発火（hooks 設定に従い SIEM 出力・audit-report-log 等）
            → Invoke-McpProviders（Slack/Teams/ServiceNow/Jira dispatch）
```

### 主要な関数

- `Write-StructuredLog` / `Write-Log` / `Write-ErrorLog` — ログ書き込み（`logs/` フォルダ）
- `Invoke-GuardedStep`（旧 `Try-Step`）— 各タスクを try-catch でラップし、失敗しても次タスクへ継続する
- `Show` — 色付きコンソール出力
- `Progress-Bar` — 進捗表示（PS7: Unicode ブロック / PS5: ASCII）
- `Test-PendingReboot` — レジストリで再起動保留を確認
- `Run-WindowsUpdate` — UsoClient.exe 経由で Windows Update を実行
- `Update-ScoreHistory` — スコア履歴を JSON に追記し Chart.js グラフを更新

### オプション機能（デフォルト無効）

| 機能 | 設定キー | 備考 |
|---|---|---|
| Slack 通知 | `mcpProviders[].type = "slack"` / `enabled: false` | Webhook URL 設定で有効化 |
| Teams 通知 | `mcpProviders[].type = "teams"` / `enabled: false` | Webhook URL 設定で有効化 |
| ServiceNow 起票 | `mcpProviders[].type = "servicenow"` / `enabled: false` | `scoreThreshold: 70` 以下で自動起票 |
| Jira 起票 | `mcpProviders[].type = "jira"` / `enabled: false` | `scoreThreshold: 70` 以下で自動起票 |
| SIEM 出力 | `hooks.siem` | `logs/siem/` フォルダに JSONL/CEF/LEEF 形式で出力 |
| AI 診断 | `Advanced.psm1` | Anthropic API キー設定で有効化 |

### ログ出力

スクリプト実行ごとにタイムスタンプ付きのログファイルを `logs/` サブフォルダに 2 種類生成する：
- `logs/PC_Optimizer_Log_YYYYMMDDHHMM.txt` — 通常ログ
- `logs/PC_Optimizer_Error_YYYYMMDDHHMM.txt` — エラーログ

### エラー処理パターン

- `Try-Step` ラッパーにより、各タスクは失敗しても後続タスクを継続する（非ブロッキング）
- ファイル削除には `-ErrorAction SilentlyContinue` を使用してロック中ファイルを安全にスキップ
- 外部コマンド（`wevtutil`, `powercfg` 等）は `$LASTEXITCODE` で成否判定

### メディア種別の自動検出

SSD と HDD を自動判別し、SSD では TRIM (`Optimize-Volume -ReTrim`)、HDD ではデフラグを実行する。判定は `Get-PhysicalDisk` の `MediaType` プロパティによる。

## 規約（必読）

| ファイル | 内容 |
|---|---|
| `docs/文字コード規約.md` | **文字コード統一規約** — `Out-File` 等への `-Encoding` 明示義務・禁止事項・CI 検査項目 |
| `docs/実装憲法.md` | **実装憲法** — 開発フェーズ・エラー処理・セルフチェックリスト |

### エンコーディング実装パターン（必須）

```powershell
# バージョン検出は関数定義より前に記述すること
$psver       = $PSVersionTable.PSVersion.Major
$isPS7Plus   = ($psver -ge 7)
$logEncoding = if ($isPS7Plus) { 'utf8NoBOM' } else { 'UTF8' }

# モジュール内では $script:_enc 変数を使用する（Common.psm1 で定義）
# $script:_enc = if ($isPS7Plus) { 'utf8NoBOM' } else { 'UTF8' }

# ファイル書き込みは必ず -Encoding を指定する
Add-Content -Path $path -Value $msg -Encoding $logEncoding
# または（モジュール内）
Add-Content -Path $path -Value $msg -Encoding $script:_enc
```

### BAT ファイルの保存（重要）

BAT ファイルは **Shift-JIS (CP932) + CRLF** で保存すること。
LF のみだと `cmd.exe` の `if (...) (...)` ブロック解析が失敗する。

```powershell
$sjis = [System.Text.Encoding]::GetEncoding(932)
$content = $lines -join "`r`n"
[System.IO.File]::WriteAllBytes($batPath, $sjis.GetBytes($content))
```

## ドキュメント

詳細ドキュメントは `docs/` フォルダに格納されています。

| ファイル | 内容 |
|---|---|
| `docs/アーキテクチャ.md` | 実行フロー・関数依存関係・拡張ポイント |
| `docs/関数リファレンス.md` | 全関数の仕様 |
| `docs/PowerShell互換性.md` | バージョン別の機能差分・コマンド可用性表 |
| `docs/削除対象パス.md` | 削除対象パスの完全一覧 |
| `docs/ログ仕様.md` | ログファイル仕様 |
| `docs/トラブルシューティング.md` | よくある問題と対処 |
| `docs/セキュリティ.md` | 権限・ログのセキュリティ考慮事項 |
| `docs/使い方.md` | 詳細な使い方・タスク一覧 |
| `docs/インストール手順.md` | インストール・前提条件 |
| `docs/変更履歴.md` | 変更履歴 |
| `docs/文字コード規約.md` | 文字コード統一規約 |
| `docs/実装憲法.md` | ClaudeCode 実装憲法 |

## 開発上の注意点

- スクリプトは **管理者権限必須**。権限なしに実行しても大半のタスクは失敗する。
- Windows Update タスク (`Run-WindowsUpdate`) は COM API で更新一覧を表示してから Y/N 確認を求める。`Y` を押すと `UsoClient.exe` 経由で実際の更新インストールが走る。テスト時は `N` を選択すること。
- Microsoft 365 タスクは `OfficeC2RClient.exe` を検出した場合のみ Y/N 確認プロンプトを表示する。テスト時は `N` を選択すること。
- ログファイルは `logs/` サブフォルダに実行のたびに生成される。定期的なクリーンアップが必要。
- ハードコードされたパス (`C:\Windows\Temp`、`%TEMP%` など) を変更する場合は `docs/削除対象パス.md` と整合させること。
- SFC (`sfc /scannow`) と DISM (`/ScanHealth`) は完了まで数分〜10 分以上かかる場合がある。
- Hook の `command` タイプは **廃止済み**（v4.0.1 でセキュリティ修正）。`webhook` / `file` タイプのみサポート。
- `Invoke-Expression` はコードベース内で使用禁止。Hook 実行には `Invoke-WebRequest` または `Add-Content` を使用する。
