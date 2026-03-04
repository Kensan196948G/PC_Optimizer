# HTMLレポート出力機能 — 設計文書

**作成日**: 2026-03-04
**対象バージョン**: v3.3.0（予定）
**ステータス**: 設計承認済み・実装待ち

---

## 概要

PC最適化完了後、全タスクの成否・ディスク解放量・システム情報・スタートアップ一覧を
まとめたHTMLレポートを自動生成し、既定ブラウザで開く機能を追加する。

---

## Section 1: アーキテクチャ変更

### 変更ファイル
- `PC_Optimizer.ps1` のみ（単一ファイル完結、外部依存なし）

### 追加グローバル変数
```powershell
$script:taskResults = @()   # タスク結果収集配列
$script:scriptStartTime = Get-Date  # 実行開始時刻
```

### Try-Step 関数への修正
- 内部で `Stopwatch` を使い所要時間を計測
- 成否と所要時間を `$script:taskResults` に追記

### 新規関数 New-HtmlReport
- 引数: `-Results`, `-DiskBefore`, `-DiskAfter`, `-SystemInfo`, `-StartupInfo`
- 出力: `logs/PC_Optimizer_Report_YYYYMMDDHHMM.html`
- 保存: `[IO.File]::WriteAllText()` で UTF-8（BOMなし）
- 完了後: `Start-Process $reportPath` で既定ブラウザ起動

---

## Section 2: HTMLレポートのデザイン・構成

### レイアウト構成
```
[ヘッダー] ─── グラデーション背景 + タイトル + 実行日時
[システム情報カード] ─── ホスト名 / OS / CPU / RAM / Disk
[ディスク解放サマリー] ─── 最適化前後のディスク空き容量 + 解放量
[タスク結果テーブル] ─── 20タスク全件：結果（✅/❌）+ 所要時間 + エラー詳細
[スタートアップ・サービス一覧] ─── タスク20の出力を転記
```

### スタイル方針
- 外部CDN不使用（完全オフライン動作）
- CSS変数で色管理（ダークヘッダー + カード形式）
- 成否表示: `✅` 成功 / `❌` 失敗 / `⚠️` 要確認

---

## Section 3: データフロー

```
PC_Optimizer.ps1 起動
  ↓
$script:scriptStartTime = Get-Date
$script:taskResults = @()
$diskBefore = (Get-PSDrive C).Free

  ↓（Try-Step 各タスク実行）

Try-Step "タスク名" { ... }
  成功 → $script:taskResults += @{ Name; Status="OK"; Duration; Error="" }
  失敗 → $script:taskResults += @{ Name; Status="NG"; Duration; Error=$_.Message }

  ↓（20タスク全完了後）

$diskAfter = (Get-PSDrive C).Free
New-HtmlReport -Results $script:taskResults ...
  → HTML生成（ヒアドキュメント）
  → logs/ フォルダへ保存
  → Start-Process で既定ブラウザ起動
```

### スコープ注意点
- `$script:taskResults` はスクリプトスコープで管理
  （`Try-Step` 内の子スコープから書き込めるようにするため）
- `Stopwatch` は `Try-Step` 内部でのみ使用（ローカルスコープ可）

---

## 非機能要件

| 項目 | 要件 |
|---|---|
| エンコーディング | UTF-8 BOMなし（`[Text.Encoding]::UTF8` 使用） |
| オフライン動作 | 外部リソース（CDN）不使用 |
| PS互換性 | PS5.1 / PS7+ 両対応 |
| エラー耐性 | HTMLレポート生成失敗は警告のみ（メイン処理は継続） |
| 保存先 | `logs/PC_Optimizer_Report_YYYYMMDDHHMM.html` |

---

## テスト計画

- [ ] `$script:taskResults` が全20タスク分のデータを持つことを確認（静的解析）
- [ ] `New-HtmlReport` 関数定義が存在することを確認
- [ ] `Start-Process $reportPath` によるブラウザ起動コードが存在することを確認
- [ ] 生成されたHTMLが有効な構造（`<html>`, `<head>`, `<body>`）を持つことを確認
