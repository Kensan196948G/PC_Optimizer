# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## プロジェクト概要

Windows 10/11 向けの PC 最適化ツール。管理者権限で実行し、一時ファイル削除・ブラウザキャッシュ・ディスク最適化・SFC/DISM・Microsoft 365 / Windows Update（確認後適用）など 19 タスクを自動実行する。

## 実行方法

```batch
# 管理者として実行（BAT ファイル経由が推奨）
Run_PC_Optimizer.bat

# PowerShell から直接実行
powershell -NoProfile -ExecutionPolicy Bypass -File "PC_Optimizer.ps1"
```

BAT ファイルは PowerShell 7 (pwsh) を優先検出し、なければ PowerShell 5.1 にフォールバックする。管理者権限がない場合は UAC で昇格要求する。

テストは `Test_Encoding.ps1` を使用する（文字コード・構文チェック）:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "Test_Encoding.ps1"
```

## アーキテクチャ

### ファイル構成

| ファイル / フォルダ | 役割 |
|---|---|
| `PC_Optimizer.ps1` | メイン最適化エンジン |
| `Run_PC_Optimizer.bat` | 管理者昇格 + PowerShell 検出ランチャー（Shift-JIS/CRLF） |
| `Test_Encoding.ps1` | 文字コード・日本語 UI・構文の検証テスト |
| `logs/` | 実行ログの出力先（自動生成） |
| `docs/` | ドキュメント群 |

### 実行フロー

```
Run_PC_Optimizer.bat
  → PowerShell 検出 & 管理者昇格
    → PC_Optimizer.ps1
      → システム情報収集・ログ初期化（logs/ フォルダに出力）
      → 初期ディスク空き容量記録
      → 19 タスクを Try-Step で順次実行
      → ディスク解放量の比較表示
      → 再起動要否チェック → 終了
```

### 主要な関数（PC_Optimizer.ps1）

- `Write-Log` / `Write-ErrorLog` — メインログ・エラーログへの書き込み（`logs/` フォルダ）
- `Show` — 色付きコンソール出力
- `Progress-Bar` — 進捗表示（PS7: Unicode ブロック / PS5: ASCII）
- `Try-Step` — 各最適化タスクを try-catch でラップし、失敗しても次タスクへ継続する
- `Test-PendingReboot` — レジストリで再起動保留を確認
- `Run-WindowsUpdate` — UsoClient.exe 経由で Windows Update を実行

### ログ出力

スクリプト実行ごとにタイムスタンプ付きのログファイルを `logs/` サブフォルダに 2 種類生成する：
- `logs/PC_Optimizer_Log_YYYYMMDDHHMM.txt` — 通常ログ
- `logs/PC_Optimizer_Error_YYYYMMDDHHMM.txt` — エラーログ

### エラー処理パターン

- `Try-Step` ラッパーにより、各タスクは失敗しても後続タスクを継続する（非ブロッキング）
- ファイル削除には `-ErrorAction SilentlyContinue` を使用してロック中ファイルを安全にスキップ
- 外部コマンド（`wevtutil`, `powercfg` 等）は `$LASTEXITCODE` で成否判定

### メディア種別の自動検出

SSD と HDD を自動判別し、SSD では TRIM (`Optimize-Volume -ReTrim`)、HDD ではデフラグを実行する。判定は `Get-PhysicalDisk` の `MediaType` プロパティによる。

## 規約（必読）

| ファイル | 内容 |
|---|---|
| `docs/文字コード規約.md` | **文字コード統一規約** — `Out-File` 等への `-Encoding` 明示義務・禁止事項・CI 検査項目 |
| `docs/実装憲法.md` | **実装憲法** — 開発フェーズ・エラー処理・セルフチェックリスト |

### エンコーディング実装パターン（必須）

```powershell
# バージョン検出は関数定義より前に記述すること
$psver       = $PSVersionTable.PSVersion.Major
$isPS7Plus   = ($psver -ge 7)
$logEncoding = if ($isPS7Plus) { 'utf8NoBOM' } else { 'UTF8' }

# ファイル書き込みは必ず -Encoding を指定する
Add-Content -Path $path -Value $msg -Encoding $logEncoding
```

### BAT ファイルの保存（重要）

BAT ファイルは **Shift-JIS (CP932) + CRLF** で保存すること。
LF のみだと `cmd.exe` の `if (...) (...)` ブロック解析が失敗する。

```powershell
$sjis = [System.Text.Encoding]::GetEncoding(932)
$content = $lines -join "`r`n"
[System.IO.File]::WriteAllBytes($batPath, $sjis.GetBytes($content))
```

## ドキュメント

詳細ドキュメントは `docs/` フォルダに格納されています。

| ファイル | 内容 |
|---|---|
| `docs/アーキテクチャ.md` | 実行フロー・関数依存関係・拡張ポイント |
| `docs/関数リファレンス.md` | 全関数の仕様 |
| `docs/PowerShell互換性.md` | バージョン別の機能差分・コマンド可用性表 |
| `docs/削除対象パス.md` | 削除対象パスの完全一覧 |
| `docs/ログ仕様.md` | ログファイル仕様 |
| `docs/トラブルシューティング.md` | よくある問題と対処 |
| `docs/セキュリティ.md` | 権限・ログのセキュリティ考慮事項 |
| `docs/使い方.md` | 詳細な使い方・タスク一覧 |
| `docs/インストール手順.md` | インストール・前提条件 |
| `docs/変更履歴.md` | 変更履歴 |
| `docs/文字コード規約.md` | 文字コード統一規約 |
| `docs/実装憲法.md` | ClaudeCode 実装憲法 |

## 開発上の注意点

- スクリプトは **管理者権限必須**。権限なしに実行しても大半のタスクは失敗する。
- Windows Update タスク (`Run-WindowsUpdate`) は COM API で更新一覧を表示してから Y/N 確認を求める。`Y` を押すと `UsoClient.exe` 経由で実際の更新インストールが走る。テスト時は `N` を選択すること。
- Microsoft 365 タスクは `OfficeC2RClient.exe` を検出した場合のみ Y/N 確認プロンプトを表示する。テスト時は `N` を選択すること。
- ログファイルは `logs/` サブフォルダに実行のたびに生成される。定期的なクリーンアップが必要。
- ハードコードされたパス (`C:\Windows\Temp`、`%TEMP%` など) を変更する場合は `docs/削除対象パス.md` と整合させること。
- SFC (`sfc /scannow`) と DISM (`/ScanHealth`) は完了まで数分〜10 分以上かかる場合がある。
