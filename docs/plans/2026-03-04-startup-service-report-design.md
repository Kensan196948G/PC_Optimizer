# タスク20 スタートアップ・サービスレポート 設計ドキュメント

**作成日**: 2026-03-04
**バージョン**: 1.0
**ステータス**: 承認済み

---

## 概要

PC最適化ツール v3.1.0 に**タスク20**として「スタートアップ・サービスレポート」を追加する。
**表示のみ（変更なし）** の読み取り専用機能として実装し、ユーザーがシステム起動時の負荷を把握できるようにする。

---

## 要件

| # | 要件 | 優先度 |
|---|---|---|
| R1 | スタートアップ登録アプリを一覧表示する | 必須 |
| R2 | 不要サービスの稼働状況を確認・表示する | 必須 |
| R3 | システムへの変更は一切行わない（読み取り専用） | 必須 |
| R4 | PS3.0 / PS5.1 / PS7.x の全バージョンで動作する | 必須 |
| R5 | 既存の `Try-Step` パターンに従い実装する | 必須 |
| R6 | 結果をログファイル（Write-Log）に記録する | 必須 |

---

## アーキテクチャ

### 実行フロー

```
Try-Step "スタートアップ・サービスレポート" {
    ┌─ フェーズ1: スタートアップアプリ一覧 ────────────────────┐
    │  Get-CimInstance Win32_StartupCommand                   │
    │  → Name / Command / Location / User を取得              │
    │  → 表示件数上限 30件（ログには全件記録）                  │
    │  → Show() でカラー表示 + Write-Log() に記録              │
    └──────────────────────────────────────────────────────┘
    ┌─ フェーズ2: 不要サービスチェック ─────────────────────────┐
    │  Get-Service で Running状態 全サービスを取得              │
    │  → 事前定義リストと照合（8サービス）                      │
    │  → 該当するものを警告色で表示 + Write-Log() に記録        │
    └──────────────────────────────────────────────────────┘
}
```

### データフロー

```
Get-CimInstance Win32_StartupCommand
        ↓
 フィルタ・整形（Name/Command/Location/User）
        ↓
 Sort-Object Location, Name
        ↓
 Show() カラー表示 + Write-Log() 記録
        ↓
Get-Service | Where-Object Status -eq 'Running'
        ↓
 不要サービスリスト（$unnecessaryServices）との照合
        ↓
 該当サービスを Yellow/Red で表示 + Write-Log() 記録
```

---

## コンポーネント詳細

### スタートアップアプリ取得

```powershell
$startupApps = Get-CimInstance -ClassName Win32_StartupCommand |
    Select-Object Name, Command, Location, User |
    Sort-Object Location, Name
```

- `Location` の例: `HKU\<SID>\SOFTWARE\Microsoft\Windows\CurrentVersion\Run`、`Startup`
- 表示上限: 30件（超過時は「他 N 件はログを参照」と表示）
- エラー時: 警告メッセージを表示してフェーズをスキップ（`Try-Step` で補足）

### 不要サービス照合リスト

```powershell
$unnecessaryServices = @(
    'Fax',         # FAXサービス（多くの環境で不使用）
    'RetailDemo',  # 小売デモモード（一般ユーザー不要）
    'XblGameSave', # Xbox ゲームセーブ（非ゲーマー不要）
    'XboxGipSvc',  # Xbox 周辺機器インターフェース
    'WpcMonSvc',   # 保護者機能モニタ（Windows 10 のみ）
    'DiagTrack',   # Connected User Experiences and Telemetry
    'MapsBroker',  # ダウンロード済み地図マネージャ
    'lfsvc'        # 位置情報サービス
)
```

### エラー処理

| 状況 | 処理 |
|---|---|
| `Win32_StartupCommand` 取得失敗 | 警告をコンソール表示 + `Write-ErrorLog` してフェーズスキップ |
| `Get-Service` 取得失敗 | 警告をコンソール表示 + `Write-ErrorLog` してフェーズスキップ |
| スタートアップアプリ0件 | 「登録なし」を表示してログ記録 |
| 不要サービス該当なし | 「問題なし」を表示してログ記録 |

---

## PS バージョン互換性

| 機能 | PS3.0 | PS5.1 | PS7.x |
|---|---|---|---|
| `Get-CimInstance` | ✓ | ✓ | ✓ |
| `Get-Service` | ✓ | ✓ | ✓ |
| 絵文字（🚀 📋） | ✗ → ASCII fallback | ✗ → ASCII fallback | ✓ |
| ANSI 太文字（`$B` / `$R`） | ✗ → 空文字 | ✗ → 空文字 | ✓ |

既存の `$isPS7Plus` フラグを使って絵文字の有無を分岐する。

---

## 実装箇所

| ファイル | 変更種別 | 内容 |
|---|---|---|
| `PC_Optimizer.ps1` | 追記のみ | タスク20のコードブロックを末尾（再起動チェック前）に追加 |
| `docs/変更履歴.md` | 追記 | v3.2.0 としてタスク20追加を記録 |
| `docs/関数リファレンス.md` | 追記 | 新規タスク20の説明を追加 |
| `tests/Test_PCOptimizer.ps1` | 追記 | タスク20用のテストケースを追加 |

---

## 非機能要件

- **変更禁止**: レジストリ・サービス設定・スタートアップ登録の変更は一切行わない
- **権限**: 既存の管理者権限セッションで動作する（追加昇格不要）
- **実行時間**: 5秒以内を目標（`Get-Service` は高速）
- **ログ**: `Write-Log` で全情報をタイムスタンプ付きで記録

---

## テスト計画

| テストケース | 確認内容 |
|---|---|
| TC-01 | `Win32_StartupCommand` の取得結果が表示される |
| TC-02 | ログファイルにスタートアップ情報が記録される |
| TC-03 | 不要サービスリストの照合ロジックが正しく動作する |
| TC-04 | 不要サービスが Running の場合に警告色で表示される |
| TC-05 | `Get-CimInstance` 失敗時に警告を出してスキップする |
| TC-06 | PS5.1 環境でも正しく動作する（絵文字なし fallback） |

---

## 変更履歴

| 日付 | バージョン | 変更内容 |
|---|---|---|
| 2026-03-04 | 1.0 | 初版作成・承認済み |
