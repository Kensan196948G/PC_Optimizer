# GUI 運用手順

## 起動方法

### バッチ起動

`GUI/Run_PC_Optimizer_GUI.bat` を右クリックして実行する。

### PowerShell 直接起動

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\GUI\PC_Optimizer_GUI.ps1
```

## 基本操作

1. 実行モードを選ぶ
2. 実行プロファイルを選ぶ
3. 対象タスクを選ぶ
4. 必要なら `WhatIf` や出力オプションを有効化する
5. `実行` を押す
6. ログビューで進行を確認する
7. 完了後に `レポート` または `ログ` を開く

## 推奨運用

- 初回は `WhatIf` を有効にして挙動確認する
- `ExportDeletedPaths` を使う場合は `WhatIf` と併用する
- `diagnose` モードでは task 20 を中心に使う
- AI 診断を使う場合は `.env` または設定ファイルを事前に確認する

## 既知制約

- GUI 実行は `-NonInteractive` 固定のため、更新確認タスク 18 / 19 は既定応答 `N` でスキップされる
- GUI は CLI の標準出力から進捗を推定するため、内部処理の実時間と完全一致はしない
- レポート閲覧自体は既定アプリで開く方式であり、GUI 内埋め込みではない

## 障害時対応

- スクリプトが起動しない場合:
  - `PC_Optimizer.ps1` の配置を確認する
  - 管理者権限で起動しているか確認する
- ログが出ない場合:
  - 実行エンジンが `pwsh` / `powershell` のどちらか確認する
  - `logs/` へファイルが生成されているか確認する
- エラー終了した場合:
  - GUI のログビューを確認する
  - `logs/PC_Optimizer_Error_*.txt` を確認する
  - `reports/Audit_Run_*.json` を確認する

