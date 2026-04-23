# GUI ドキュメント

PowerShell ベースの GUI フロントエンドに関する仕様書群です。  
GUI は `PC_Optimizer.ps1` をバックエンド実行エンジンとして利用し、CLI と同じ引数体系を可能な範囲で画面操作に落とし込みます。

## ドキュメント一覧

| ファイル | 内容 |
|---|---|
| [GUI要件定義.md](GUI要件定義.md) | GUI の目的、対象範囲、前提条件、制約 |
| [GUI画面仕様.md](GUI画面仕様.md) | 画面構成、各コントロールの役割、表示ルール |
| [GUIイベント仕様.md](GUIイベント仕様.md) | ボタン、入力変更、実行時イベント、バリデーション |
| [GUIアーキテクチャ.md](GUIアーキテクチャ.md) | 実装構成、実行フロー、CLI 連携方式 |
| [GUI運用手順.md](GUI運用手順.md) | 起動方法、操作手順、既知制約、障害時対応 |

## 実装配置

- 実装本体: `GUI/PC_Optimizer_GUI.ps1`
- GUI ランチャー: `GUI/Run_PC_Optimizer_GUI.bat`
- GUI 利用案内: `GUI/README.md`

## 静的検証コマンド

GUI の静的検証は `tests/Test_GUI.ps1` で実行します。GUI の起動確認や品質ゲートが変わる場合は、ルート `README.md` とこのドキュメントも更新してください。

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test_GUI.ps1
pwsh -NoProfile -ExecutionPolicy Bypass -File tests\Test_GUI.ps1
```

GUI 本体を直接起動する場合:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File GUI\PC_Optimizer_GUI.ps1
pwsh -NoProfile -ExecutionPolicy Bypass -File GUI\PC_Optimizer_GUI.ps1
```

## 設計方針

- GUI は CLI を置き換えず、`PC_Optimizer.ps1` のフロントエンドとして動作する
- 実行時は `-NonInteractive -NoRebootPrompt` を付与し、GUI が入力待ちで停止しないようにする
- 進捗は標準出力の監視と選択タスク数に基づいて GUI 側で可視化する
- レポート生成やログ出力は既存 CLI 実装の責務をそのまま再利用する
