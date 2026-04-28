# GUI 画面仕様

## 画面概要

GUI は単一メインウィンドウ構成とする。
PowerShell 7.x を優先実行エンジンとし、画像ヘッダーと色分けされた操作ボタンを持つポップな画面構成とする。

| 領域 | 内容 |
|---|---|
| 最上部 | ヒーローヘッダー、実行エンジン情報、権限状態、モード要約 |
| 上部 | 実行構成サマリー、主要オプション |
| 左側 | タスク選択チェックリスト |
| 右上 | コマンドプレビュー、実行ステータス |
| 右下 | 実行ログビュー |
| 下部 | 実行、停止、フォルダを開く各種操作ボタン |

## コントロール一覧

### 実行設定

| コントロール | 種別 | 説明 |
|---|---|---|
| `Mode` | ComboBox | `repair` / `diagnose` |
| `ExecutionProfile` | ComboBox | `classic` / `agent-teams` |
| `FailureMode` | ComboBox | `continue` / `fail-fast` |
| `ConfigPath` | TextBox | 任意の設定 JSON パス |
| `ExportDeletedPathsPath` | TextBox | WhatIf 出力先フォルダ |
| `Config 参照` | Button | `ConfigPath` 選択 |
| `Export 参照` | Button | `ExportDeletedPathsPath` 選択 |

### オプション

| コントロール | 種別 | 説明 |
|---|---|---|
| `WhatIf` | CheckBox | 副作用なし実行 |
| `UseLocalChartJs` | CheckBox | ローカル Chart.js を使用 |
| `ExportPowerBIJson` | CheckBox | Power BI JSON を生成 |
| `UseAnthropicAI` | CheckBox | Anthropic AI を有効化 |
| `ExportDeletedPaths` | CheckBox | 削除候補を出力 |
| `ExportDeletedPathsFormat` | ComboBox | `json` / `csv` |

### タスク選択

- 20 タスクをチェックリストで表示する
- 初期状態は全選択とする
- 補助ボタンを提供する
  - 全選択
  - 全解除
  - 診断向け選択
- タスク 18 / 19 は選択可能だが、GUI 実行では更新適用は自動スキップされることを画面上で説明する

### 実行状態

| 項目 | 内容 |
|---|---|
| 状態ラベル | `待機中` / `実行中` / `完了` / `失敗` / `停止` |
| 現在タスク | 出力から推定した実行中タスク名 |
| 経過時間 | 起動からの経過秒数 |
| 終了コード | CLI 戻り値 |
| 進捗バー | 選択タスク数に対する完了率 |
| 最新 HTML | 最新 HTML レポートのパス表示と起動ボタン |
| 最新 JSON | 最新 JSON レポートのパス表示と起動ボタン |

### ログビュー

- 読み取り専用の複数行テキストボックス
- 標準出力 / 標準エラーを時系列で追記表示
- ANSI 制御シーケンスを除去したテキストを表示する
- 自動スクロールを有効にする

## 画面サイズ

- 初期クライアントサイズ: 1360 x 900
- 最小サイズ: 1200 x 900
- リサイズ時はログビューが優先的に拡張されること
