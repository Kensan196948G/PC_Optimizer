# GUI

`gui/PCOptimizer.Gui.ps1` は、既存の `PC_Optimizer.ps1` をそのまま実行エンジンとして利用する WPF GUI です。

## 概要

- GUI と CUI は同じ PowerShell エンジンを共有します
- GUI は管理者権限で再起動し、`-NonInteractive` と `-NoRebootPrompt` を付けて実行します
- 実行ログ、進捗、状態、最新レポートを GUI 上で確認できます
- 実行完了後は GUI からレポートとログを開けます

## 起動方法

推奨:

```powershell
.\Run_PC_Optimizer_GUI.bat
```

直接起動:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -STA -File .\gui\PCOptimizer.Gui.ps1
```

`pwsh` がない環境では Windows PowerShell 5.1 でも起動できます。
