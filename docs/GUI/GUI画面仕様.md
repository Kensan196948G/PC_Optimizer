# GUI 画面仕様

## 画面概要

GUI は単一メインウィンドウ構成とする。
PowerShell 5.1 互換を優先するため、GUI 内のラベルやボタンに絵文字アイコンは使用しない。

| 領域 | 内容 |
|---|---|
| 上部 | 実行エンジン情報、スクリプト配置、主要オプション |
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

### 実行状態

| 項目 | 内容 |
|---|---|
| 状態ラベル | `待機中` / `実行中` / `完了` / `失敗` / `停止` |
| 現在タスク | 出力から推定した実行中タスク名 |
| 経過時間 | 起動からの経過秒数 |
| 終了コード | CLI 戻り値 |
| 進捗バー | 選択タスク数に対する完了率 |

### ログビュー

- 読み取り専用の複数行テキストボックス
- 標準出力 / 標準エラーを時系列で追記表示
- 自動スクロールを有効にする

## 画面サイズ

- 初期サイズ: 1200 x 780
- 最小サイズ: 1024 x 700
- リサイズ時はログビューが優先的に拡張されること
