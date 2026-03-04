# PowerShell バージョン互換性

## 対応バージョン一覧

| PowerShell | 状態 | 備考 |
|---|---|---|
| 2.0 | 非対応 | `Get-CimInstance` が存在しない |
| 3.0 | 警告あり・動作可 | 最小サポートバージョン |
| 4.0 | 警告あり・動作可 | Windows 8.1 標準搭載 |
| 5.0 | 完全対応 | |
| **5.1** | **完全対応（推奨）** | Windows 10/11 標準搭載 |
| 6.x | 完全対応 | `Get-WmiObject` 削除済みだが本ツールは `Get-CimInstance` を使用 |
| **7.x** | **完全対応（推奨）** | 絵文字・ANSI カラー・Unicode ブロック進捗バーが有効 |

## バージョン別の機能差分

### 表示スタイル

| 機能 | PS 5.1 以下 | PS 7.x 以降 |
|---|---|---|
| システム情報のアイコン | `*` ASCII 記号 | 🖥️ 👤 🏷️ 💻 🧠 💾 絵文字 |
| 進捗バー文字 | `#---` ハッシュ記号 | `██░░` Unicode ブロック |
| 完了メッセージ | テキストのみ | ✅ 付き |

### コマンドの可用性

| コマンド / API | PS 3.0 | PS 5.1 | PS 7.x |
|---|---|---|---|
| `Get-CimInstance` | ✅ | ✅ | ✅ |
| `Get-WmiObject` | ✅ | ✅ | ❌ 削除済み |
| `Get-PhysicalDisk` | ✅* | ✅ | ✅ |
| `Optimize-Volume` | ✅* | ✅ | ✅ |
| `Add-Content` | ✅ | ✅ | ✅ |
| `Start-Process` | ✅ | ✅ | ✅ |
| `Read-Host` | ✅ | ✅ | ✅ |
| `wevtutil.exe` | ✅ | ✅ | ✅ |
| `sfc.exe` | ✅ | ✅ | ✅ |
| `DISM.exe` | ✅ | ✅ | ✅ |
| `powercfg.exe` | ✅ | ✅ | ✅ |

*Storage モジュールが必要（Windows 8.1+ に標準搭載）

### `Get-PhysicalDisk` の可用性チェック

本ツールは `Get-PhysicalDisk` を使用する前に以下のチェックを行います。

```powershell
if (Get-Command -Name Get-PhysicalDisk -ErrorAction SilentlyContinue) {
    # SSD/HDD 判定・SMART チェック
} else {
    # スキップしてログに記録
}
```

Storage モジュールが利用できない環境ではスキップされ、処理は継続します。

### `UsoClient.exe` の可用性チェック

```powershell
$usoPath = Join-Path $env:SystemRoot "System32\UsoClient.exe"
if (-not (Test-Path $usoPath)) {
    # Windows Update をスキップ
}
```

`UsoClient.exe` は Windows 10 以降にのみ存在します。

## 各バージョンの取得方法

### PowerShell 5.1 の確認（Windows 標準）

```powershell
$PSVersionTable.PSVersion
# 期待出力: Major:5, Minor:1
```

### PowerShell 7 のインストール

```powershell
# winget 経由（Windows 10/11）
winget install Microsoft.PowerShell

# または公式サイトからインストーラーをダウンロード
# https://github.com/PowerShell/PowerShell/releases
```

### インストール済みバージョンの確認

```powershell
# PowerShell 5.1 のパス
Get-Command powershell | Select-Object -ExpandProperty Source
# C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe

# PowerShell 7 のパス（インストール済みの場合）
Get-Command pwsh | Select-Object -ExpandProperty Source
# C:\Program Files\PowerShell\7\pwsh.exe
```

## 実行時のバージョン表示

スクリプト実行時にログに PowerShell バージョンが記録されます。

```
[PowerShell バージョン] 7.4.1
```

コンソールでも確認できます：

```powershell
$PSVersionTable.PSVersion.ToString()
```
