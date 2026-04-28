# GUI

PowerShell GUI フロントエンドです。  
`PC_Optimizer.ps1` をバックエンドとして起動し、CLI 引数を画面操作で指定できます。

## ファイル

| ファイル | 役割 |
|---|---|
| `PC_Optimizer_GUI.ps1` | GUI 本体 |
| `Run_PC_Optimizer_GUI.bat` | GUI 起動用ランチャー |
| `PC2026.jpg` | ヘッダー表示用イメージ |

## 起動方法

### バッチ起動

`Run_PC_Optimizer_GUI.bat` を実行します。  
ランチャーは `pwsh.exe` (PowerShell 7) を優先し、見つからない場合のみ `powershell.exe` にフォールバックします。

### PowerShell 直接起動

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\GUI\PC_Optimizer_GUI.ps1
```

## GUI の特徴

- PowerShell 7 系優先の WinForms GUI
- ポップなカラーテーマとヘッダー画像
- コマンドプレビュー、進捗バー、リアルタイムログ表示
- 最新 HTML / JSON レポートをワンクリックで開ける
- `reports/` `logs/` `docs/GUI/` `repo root` へのショートカットを用意

## 注意

- GUI は内部で `PC_Optimizer.ps1` を別プロセス起動します
- GUI 実行時は `-NonInteractive -NoRebootPrompt` を付与します
- `ExportDeletedPaths` は `WhatIf` と併用してください
- タスク 18 / 19 は選択可能ですが、GUI 実行では既定応答 `N` により更新適用はスキップされます
- 詳細仕様は `docs/GUI/` を参照してください
