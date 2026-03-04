# 高性能PC最適化ツール

Windows 10/11 向けのワンクリック PC メンテナンスツールです。
一時ファイル削除・ディスク最適化・Windows Update など 20 タスクを管理者権限で自動実行します。

## クイックスタート

1. `PC_Optimizer.ps1` と `Run_PC_Optimizer.bat` を同じフォルダに配置する
2. `Run_PC_Optimizer.bat` を**右クリック → 管理者として実行**
3. UAC プロンプトで **[はい]** を選択
4. カラフルな進捗表示で自動処理が進む
5. 完了後、`Enter` キーで終了

## 対応環境

| 項目 | 要件 |
|---|---|
| OS | Windows 10 / 11 |
| PowerShell | 3.0 以上（5.1 / 7.x 推奨） |
| 権限 | 管理者権限必須 |

## 実行される主な処理（20 タスク）

| # | タスク |
|---|---|
| 1 | Windows / ユーザー 一時ファイル削除 |
| 2 | Prefetch・Windows Update・Delivery Optimization キャッシュ削除 |
| 3 | 配信最適化キャッシュ削除 |
| 4 | Windows Update キャッシュ削除 |
| 5 | エラーレポート・CBS ログ削除 |
| 6 | OneDrive / Teams / Office キャッシュ削除 |
| 7 | ブラウザキャッシュ削除（Chrome / Edge / Firefox / Brave / Opera / Vivaldi） |
| 8 | サムネイルキャッシュ削除 |
| 9 | Microsoft Store キャッシュ削除 |
| 10 | ゴミ箱の空化 |
| 11 | DNS キャッシュクリア |
| 12 | Windows イベントログのクリア（Application / System） |
| 13 | ディスク最適化（SSD: TRIM / HDD: デフラグ 自動判定） |
| 14 | SSD SMART 健康診断 |
| 15 | システムファイルの整合性チェック・修復（SFC） |
| 16 | Windows コンポーネントストアの診断（DISM） |
| 17 | 電源プランの最適化（ノートPC: バランス / デスクトップ: 高パフォーマンス） |
| 18 | **Microsoft 365 の更新確認・適用**（Y/N 確認あり） |
| 19 | **Windows Update の確認・適用**（更新一覧表示 → Y/N 確認あり） |
| 20 | **スタートアップ・サービスレポート**（読み取り専用） |

実行ログは `logs/` サブフォルダに自動保存されます。

## 運用ポリシー

Git 管理対象・除外対象の基準は [リポジトリ運用方針.md](リポジトリ運用方針.md) を参照してください。
`release/v3.3.0` の保守専用ルールも同ドキュメントに記載しています。
CI artifact（テストログ）の標準保持期間は 14 日です。

## ドキュメント一覧

| ファイル | 内容 |
|---|---|
| [インストール手順.md](インストール手順.md) | インストール・前提条件 |
| [使い方.md](使い方.md) | 詳細な使い方 |
| [アーキテクチャ.md](アーキテクチャ.md) | コード設計・アーキテクチャ |
| [関数リファレンス.md](関数リファレンス.md) | 関数リファレンス |
| [ログ仕様.md](ログ仕様.md) | ログファイル仕様 |
| [削除対象パス.md](削除対象パス.md) | 削除対象パス一覧 |
| [PowerShell互換性.md](PowerShell互換性.md) | PowerShell バージョン互換性 |
| [トラブルシューティング.md](トラブルシューティング.md) | トラブルシューティング |
| [セキュリティ.md](セキュリティ.md) | セキュリティ考慮事項 |
| [変更履歴.md](変更履歴.md) | 変更履歴 |
| [文字コード規約.md](文字コード規約.md) | 文字コード統一規約 |
| [実装憲法.md](実装憲法.md) | ClaudeCode 実装憲法 |
| [リポジトリ運用方針.md](リポジトリ運用方針.md) | Git 追跡対象・除外対象の方針 |

## 開発計画・スキーマ

- [plans/2026-03-04-v3.3.2-cli-foundation.md](plans/2026-03-04-v3.3.2-cli-foundation.md)
- [plans/2026-03-04-v3.3.2-pr-split.md](plans/2026-03-04-v3.3.2-pr-split.md)
- [schemas/pc-optimizer-report-v1.schema.json](schemas/pc-optimizer-report-v1.schema.json)

## v4.0 再設計

- [plans/2026-03-04-v4.0-rearchitecture.md](plans/2026-03-04-v4.0-rearchitecture.md)
- [plans/2026-03-04-v4.0-cli-spec.md](plans/2026-03-04-v4.0-cli-spec.md)
- [plans/2026-03-04-v4.0-pr-plan.md](plans/2026-03-04-v4.0-pr-plan.md)
- [plans/2026-03-04-v4.0-test-strategy.md](plans/2026-03-04-v4.0-test-strategy.md)
- [plans/2026-03-04-v4.0-migration-plan.md](plans/2026-03-04-v4.0-migration-plan.md)
