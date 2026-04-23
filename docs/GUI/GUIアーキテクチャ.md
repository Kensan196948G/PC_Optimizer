# GUI アーキテクチャ

## 構成

```text
GUI/PC_Optimizer_GUI.ps1
  ├─ Windows Forms UI
  ├─ 入力値バリデーション
  ├─ CLI 引数生成
  ├─ 子プロセス起動
  ├─ 標準出力/標準エラー監視
  ├─ 進捗推定
  └─ レポート/ログ導線

PC_Optimizer.ps1
  ├─ 実最適化処理
  ├─ ログ生成
  ├─ レポート生成
  └─ Agent Teams 実行
```

## 設計原則

- GUI は UI 層に限定し、最適化処理は既存 CLI を再利用する
- GUI から CLI への入力は引数に正規化して受け渡す
- 実行状態の可視化は標準出力と終了コードを用いる
- 実行生成物は既存の `logs/` `reports/` をそのまま使う

## 子プロセス実行方式

- GUI 本体は UI 応答性維持のため、`PC_Optimizer.ps1` を別プロセス起動する
- `UseShellExecute = false`
- `RedirectStandardOutput = true`
- `RedirectStandardError = true`
- `CreateNoWindow = true`

## 進捗推定方式

- GUI は既知タスク名 20 件を内部に保持する
- 出力行にタスク識別子または状態語が含まれると現在タスク候補とする
- 出力行に `完了` `失敗` `スキップ` `skip` が含まれる場合、重複カウントを避けて完了済みとしてカウントする
- 進捗率 = `完了済み選択タスク数 / 選択タスク数`

## 互換性方針

- PowerShell 5.1 / 7.x の双方で動く Windows Forms を採用する
- GUI 内で使う入出力エンコーディングは CLI 側の既存ログ方針に依存しない
- AI / Notification / Agent Teams の詳細ロジックは CLI 側責務とする
- GUI ランチャーとバックエンド実行は Windows PowerShell 5.1 を優先し、絵文字表示に依存しない
- 視覚的なアイコンが必要な場合は Unicode 絵文字ではなく、Windows Forms の `Icon` / `ImageList` / `PictureBox` で画像を表示する
