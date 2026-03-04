# config仕様

対象ファイル: `config/config.json`

## 主要キー

| キー | 型 | 既定値 | 説明 |
|---|---|---|---|
| `logLevel` | string | `INFO` | ログ出力レベル。現時点は将来拡張用の設定値として保持。 |
| `reportFormat` | string | `HTML` | レポート形式の既定値。 |
| `cleanupBrowserCache` | bool | `true` | ブラウザキャッシュ関連処理の有効/無効フラグ。 |
| `diskOptimize` | bool | `true` | ディスク最適化処理の有効/無効フラグ。 |
| `enableSecurityCheck` | bool | `true` | セキュリティ診断処理の有効/無効フラグ。 |
| `failureMode` | string | `continue` | 失敗時動作の既定値（`continue` / `fail-fast`）。 |
| `defaultTasks` | string | `all` | 実行対象タスクの既定指定（`all` または `1,3,7` 形式）。 |

## WhatIf 削除候補マップ

`whatIfDeletedPathMap` は、`-WhatIf` 実行時に削除候補として出力するパスの定義です。

- キー: タスク番号（文字列）
- 値: パス配列（環境変数 `%TEMP%` 形式を利用可）

例:

```json
{
  "whatIfDeletedPathMap": {
    "1": [
      "%SystemRoot%\\Temp",
      "%TEMP%"
    ],
    "7": [
      "%LOCALAPPDATA%\\Google\\Chrome\\User Data\\Default\\Cache"
    ]
  }
}
```

## 注意事項

- 無効な JSON は起動時エラーになります。
- パスは実行環境で展開されるため、ユーザー/端末差分を考慮して定義してください。
