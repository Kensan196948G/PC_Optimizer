# GUI

PowerShell GUI フロントエンドです。  
`PC_Optimizer.ps1` をバックエンドとして起動し、CLI 引数を画面操作で指定できます。

## ファイル

| ファイル | 役割 |
|---|---|
| `PC_Optimizer_GUI.ps1` | GUI 本体 |
| `Run_PC_Optimizer_GUI.bat` | GUI 起動用ランチャー |

## 起動方法

### バッチ起動

`Run_PC_Optimizer_GUI.bat` を実行します。

### PowerShell 直接起動

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\GUI\PC_Optimizer_GUI.ps1
```

## 注意

- GUI は内部で `PC_Optimizer.ps1` を別プロセス起動します
- GUI 実行時は `-NonInteractive -NoRebootPrompt` を付与します
- GUI は PowerShell 5.1 互換を優先し、絵文字アイコンではなく通常文字と必要に応じた画像アイコンを前提にします
- タスク 18 / 19 は GUI v1 では既定応答によりスキップされます
- 詳細仕様は `docs/GUI/` を参照してください
