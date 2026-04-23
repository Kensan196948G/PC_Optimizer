# GUI 要件定義

## 目的

`PC_Optimizer.ps1` の既存 CLI 機能に加えて、非 PowerShell 利用者でも操作しやすい Windows 向け GUI を提供する。

## 対象ユーザー

- PowerShell 引数に不慣れな一般利用者
- 実行前にモード、タスク、出力オプションを視覚的に確認したい運用担当者
- レポートやログを GUI から開きたい保守担当者

## 対象範囲

GUI v1 の対象は以下とする。

- 実行モード選択: `repair` / `diagnose`
- 実行プロファイル選択: `classic` / `agent-teams`
- タスク選択: 1〜20 の個別選択
- 主要実行オプション選択:
  - `WhatIf`
  - `FailureMode`
  - `UseLocalChartJs`
  - `ExportPowerBIJson`
  - `UseAnthropicAI`
  - `ExportDeletedPaths`
- 実行ログのリアルタイム表示
- 実行進捗の可視化
- `logs/` `reports/` `docs/GUI/` への導線

## 非対象

GUI v1 では以下は対象外とする。

- CLI 本体ロジックの GUI への再実装
- タスク 18 / 19 の対話式 Y/N 確認を GUI 内ダイアログへ変換すること
- CLI で生成する HTML / JSON / CSV レポートの GUI 内埋め込み表示
- 複数 PC の一括管理画面

## 前提条件

- Windows 10 / 11
- PowerShell 5.1 以上
- GUI スクリプトは Windows Forms を利用する
- バックエンド実行対象の `PC_Optimizer.ps1` がリポジトリルートに存在する

## 機能要件

### FR-01 起動

- GUI は `GUI/PC_Optimizer_GUI.ps1` から起動できること
- `GUI/Run_PC_Optimizer_GUI.bat` からも起動できること
- 管理者権限が不足する場合、GUI 自身が昇格再起動を試行すること

### FR-02 設定入力

- モード、実行プロファイル、失敗モードをコンボボックスで選択できること
- タスク 1〜20 をチェックリストで選択できること
- `ConfigPath` と `ExportDeletedPathsPath` をテキストボックスで指定できること

### FR-03 実行

- GUI は選択内容から CLI 引数を組み立てること
- 実行時に `-NonInteractive -NoRebootPrompt` を付与すること
- バックエンドは別プロセスとして起動すること
- 実行中に標準出力・標準エラー出力を GUI 上に表示すること

### FR-04 進捗

- 選択タスク数に対する完了数から進捗バーを更新すること
- 実行中タスク名、経過秒数、終了コードを表示すること

### FR-05 実行後操作

- `reports/` を開けること
- `logs/` を開けること
- `docs/GUI/` を開けること
- 最新レポートパスを GUI 上で確認できること

## 制約

- GUI v1 は CLI の標準出力を利用した進捗可視化であり、CLI 内部の厳密なパーセント値とは独立する
- GUI 実行では `-NonInteractive` 固定のため、タスク 18 / 19 の更新適用は既定応答 `N` となりスキップされる
- `ExportDeletedPaths` は安全のため `WhatIf` 併用を前提とする
- PowerShell 5.1 互換を優先するため、GUI 内の状態表示に絵文字アイコンは使用しない。アイコンが必要な場合は `.ico` / `.png` などの画像アセットを使う

## 品質要件

- PowerShell 5.1 / 7.x の両方で起動可能であること
- GUI 実装は単独フォルダ `GUI/` に閉じること
- レポジトリ内既存 CLI ロジックへの侵襲を最小にすること
- 実行失敗時も GUI が落ちずに終了コードとエラー内容を表示すること
