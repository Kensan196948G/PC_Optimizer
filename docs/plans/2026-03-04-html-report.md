# HTMLレポート出力機能 実装計画

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** PC最適化完了後、全タスク成否・ディスク解放量・システム情報・スタートアップ一覧をまとめたHTMLレポートを `logs/` フォルダに自動生成し、既定ブラウザで開く。

**Architecture:** `Try-Step` 関数を修正してタスク結果を `$script:taskResults` 配列に蓄積し、スクリプト末尾の新規関数 `New-HtmlReport` がその配列からインラインHTMLを生成する。外部CDN不使用・完全オフライン動作。

**Tech Stack:** PowerShell 5.1 / 7.x、Here-string HTML、CSS変数、`[System.Diagnostics.Stopwatch]`、`[System.IO.File]::WriteAllText()`

---

## Task 1: `$script:taskResults` 初期化 + システム情報ハッシュテーブル追加

**Files:**
- Modify: `PC_Optimizer.ps1:44-45`（`$logEncoding` 定義の直後）
- Modify: `PC_Optimizer.ps1:142-145`（システム情報取得 try ブロックの末尾）

**Step 1: `$script:taskResults` 初期化を追加**

`$logEncoding = if ($isPS7Plus) { 'utf8NoBOM' } else { 'UTF8' }` の直後（45行目付近）に追加：

```powershell
# HTMLレポート用タスク結果収集
$script:taskResults    = @()
$script:scriptStartTime = Get-Date
```

**Step 2: システム情報をハッシュテーブルに収集**

`Write-Log "[PowerShell バージョン] $pwv"` の直後（141行目付近）に追加：

```powershell
    # HTMLレポート用にまとめて保持
    $script:sysInfo = @{
        Hostname  = $hostname
        Username  = $username
        OS        = $os
        CPU       = "$cpu  $cpuCores コア / $cpuThreads スレッド"
        RAM       = "${mem} GB"
        Disk      = "C: 空き ${free}GB / ${total}GB"
        PSVersion = $pwv
    }
```

**Step 3: 変更を確認（実行しない）**

`grep -n "taskResults\|sysInfo\|scriptStartTime" PC_Optimizer.ps1` で3行が表示されることを確認。

---

## Task 2: `Try-Step` に Stopwatch + 結果記録を追加

**Files:**
- Modify: `PC_Optimizer.ps1:161-174`（`Try-Step` 関数全体）

**Step 1: Try-Step を書き換える**

```powershell
function Try-Step ($desc, [ScriptBlock]$action) {
    $start = Get-Date
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Log "[$($start.ToString('HH:mm:ss'))] $desc 開始..."
    Progress-Bar "$desc..." 0
    try {
        & $action
        $sw.Stop()
        Progress-Bar "$desc 完了" 100
        Write-Log "[$((Get-Date).ToString('HH:mm:ss'))] $desc 完了"
        $script:taskResults += [PSCustomObject]@{
            Name     = $desc
            Status   = "OK"
            Duration = [int]$sw.Elapsed.TotalSeconds
            Error    = ""
        }
    } catch {
        $sw.Stop()
        Progress-Bar "$desc 失敗" 100
        Write-Log "[$((Get-Date).ToString('HH:mm:ss'))] $desc 失敗"
        Write-ErrorLog "$desc : $_"
        $script:taskResults += [PSCustomObject]@{
            Name     = $desc
            Status   = "NG"
            Duration = [int]$sw.Elapsed.TotalSeconds
            Error    = "$_"
        }
    }
}
```

**Step 2: 変更確認**

`grep -n "Stopwatch\|taskResults\|PSCustomObject" PC_Optimizer.ps1` で期待行が存在することを確認。

---

## Task 3: `New-HtmlReport` 関数を追加

**Files:**
- Modify: `PC_Optimizer.ps1:196`（`Run-WindowsUpdate` 関数の前、または `Test-PendingReboot` の後）

**Step 1: 関数全体を挿入**

`function Test-PendingReboot {` の直前に次の関数を追加：

```powershell
function New-HtmlReport {
    param(
        [PSCustomObject[]]$Results,
        [double]$DiskBefore,
        [double]$DiskAfter,
        [hashtable]$SysInfo
    )

    $reportTime  = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
    $reportStamp = Get-Date -Format "yyyyMMddHHmm"
    $reportPath  = Join-Path $logsDir "PC_Optimizer_Report_${reportStamp}.html"

    $diskFreed   = [math]::Round($DiskAfter - $DiskBefore, 2)
    $diskSign    = if ($diskFreed -ge 0) { "+$diskFreed" } else { "$diskFreed" }
    $diskColor   = if ($diskFreed -ge 0) { "#4caf50" } else { "#ff9800" }

    # --- タスク行の生成 ---
    $taskRows = ""
    $okCount  = 0
    $ngCount  = 0
    foreach ($r in $Results) {
        if ($r.Status -eq "OK") {
            $okCount++
            $badge  = '<span class="badge ok">OK</span>'
            $rowCls = ""
        } else {
            $ngCount++
            $badge  = '<span class="badge ng">NG</span>'
            $rowCls = ' class="ng-row"'
        }
        $errCell = if ($r.Error) { "<td class='err'>$([System.Web.HttpUtility]::HtmlEncode($r.Error))</td>" } else { "<td>—</td>" }
        $taskRows += "<tr$rowCls><td>$([System.Web.HttpUtility]::HtmlEncode($r.Name))</td><td>$badge</td><td>$($r.Duration) 秒</td>$errCell</tr>`n"
    }

    # --- システム情報行の生成 ---
    $sysRows = ""
    foreach ($key in @('Hostname','Username','OS','CPU','RAM','Disk','PSVersion')) {
        $val = if ($SysInfo.ContainsKey($key)) { [System.Web.HttpUtility]::HtmlEncode($SysInfo[$key]) } else { "—" }
        $sysRows += "<tr><th>$key</th><td>$val</td></tr>`n"
    }

    $html = @"
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>PC最適化レポート — $reportTime</title>
<style>
  :root {
    --bg: #1a1a2e; --card: #16213e; --accent: #0f3460;
    --text: #e0e0e0; --sub: #a0a0b0; --ok: #4caf50;
    --ng: #f44336; --warn: #ff9800; --border: #2d2d4e;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: "Segoe UI", sans-serif; background: var(--bg); color: var(--text); font-size: 14px; }
  header { background: linear-gradient(135deg, #0f3460, #16213e); padding: 24px 32px; border-bottom: 2px solid var(--accent); }
  header h1 { font-size: 22px; font-weight: 700; }
  header p { color: var(--sub); margin-top: 4px; font-size: 13px; }
  main { max-width: 960px; margin: 24px auto; padding: 0 16px; }
  .card { background: var(--card); border: 1px solid var(--border); border-radius: 8px; margin-bottom: 20px; overflow: hidden; }
  .card-title { background: var(--accent); padding: 10px 16px; font-size: 14px; font-weight: 600; letter-spacing: .05em; }
  table { width: 100%; border-collapse: collapse; }
  th, td { padding: 8px 14px; border-bottom: 1px solid var(--border); text-align: left; }
  th { color: var(--sub); width: 140px; font-weight: normal; }
  tr:last-child td, tr:last-child th { border-bottom: none; }
  .ng-row { background: rgba(244,67,54,.08); }
  .badge { display: inline-block; padding: 2px 10px; border-radius: 12px; font-size: 12px; font-weight: 700; }
  .badge.ok { background: rgba(76,175,80,.2); color: var(--ok); }
  .badge.ng { background: rgba(244,67,54,.2); color: var(--ng); }
  .err { color: var(--ng); font-size: 12px; max-width: 320px; word-break: break-all; }
  .summary { display: flex; gap: 16px; padding: 16px; flex-wrap: wrap; }
  .summary-item { flex: 1; min-width: 120px; text-align: center; background: var(--accent); border-radius: 8px; padding: 12px; }
  .summary-item .val { font-size: 28px; font-weight: 700; }
  .summary-item .lbl { font-size: 12px; color: var(--sub); margin-top: 4px; }
  footer { text-align: center; color: var(--sub); font-size: 12px; padding: 24px; }
</style>
</head>
<body>
<header>
  <h1>🛠️ PC 最適化レポート</h1>
  <p>実行日時: $reportTime</p>
</header>
<main>

<div class="card">
  <div class="card-title">💾 システム情報</div>
  <table>$sysRows</table>
</div>

<div class="card">
  <div class="card-title">📊 ディスク解放サマリー</div>
  <div class="summary">
    <div class="summary-item"><div class="val">${DiskBefore} GB</div><div class="lbl">最適化前 空き容量</div></div>
    <div class="summary-item"><div class="val">${DiskAfter} GB</div><div class="lbl">最適化後 空き容量</div></div>
    <div class="summary-item"><div class="val" style="color:$diskColor">$diskSign GB</div><div class="lbl">解放量</div></div>
  </div>
</div>

<div class="card">
  <div class="card-title">✅ タスク結果 ($okCount 成功 / $ngCount 失敗 / $($Results.Count) 件)</div>
  <table>
    <thead><tr><th style="width:55%">タスク名</th><th style="width:8%">結果</th><th style="width:10%">所要時間</th><th>エラー詳細</th></tr></thead>
    <tbody>$taskRows</tbody>
  </table>
</div>

</main>
<footer>PC_Optimizer v3.3.0 — $(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')</footer>
</body>
</html>
"@

    try {
        [System.IO.File]::WriteAllText($reportPath, $html, [System.Text.Encoding]::UTF8)
        Write-Log "[HTMLレポート] 保存完了: $reportPath"
        Show "HTMLレポートを生成しました: $reportPath" Green
        Start-Process $reportPath
    } catch {
        Write-ErrorLog "HTMLレポートの生成に失敗しました: $_"
        Show "HTMLレポートの生成に失敗しました（スキップ）" Yellow
    }
}
```

**Step 2: 変更確認**

`grep -n "New-HtmlReport\|WriteAllText\|Start-Process \$reportPath" PC_Optimizer.ps1` で3行が表示されることを確認。

---

## Task 4: `New-HtmlReport` 呼び出しをスクリプト末尾に追加

**Files:**
- Modify: `PC_Optimizer.ps1:741`（`} catch { Write-ErrorLog "ディスク空き容量の比較に失敗..."` の直後）

**Step 1: ディスク比較ブロック直後に呼び出しを追加**

`Write-Log "[ディスク解放] ..."` の後、`# 再起動が必要な場合はプロンプト表示` の前に：

```powershell
# ── HTMLレポート生成 ────────────────────────────────────
New-HtmlReport -Results $script:taskResults `
               -DiskBefore $initialFreeGB `
               -DiskAfter  $finalFreeGB `
               -SysInfo    $script:sysInfo
```

**Step 2: 変更確認**

`grep -n "New-HtmlReport" PC_Optimizer.ps1` で関数定義と呼び出しの2行が表示されることを確認。

---

## Task 5: テストを追加（Section 15）

**Files:**
- Modify: `tests/Test_PCOptimizer.ps1`（末尾に追記）

**Step 1: テスト失敗を確認してから追加**

まず現状でテストを実行（新しいテスト項目が存在しないことを確認）：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test_PCOptimizer.ps1 2>&1 | tail -5
```

期待出力: `PASS : 75 / 75`

**Step 2: Section 15 のテストを追加**

`tests/Test_PCOptimizer.ps1` の末尾に追加：

```powershell
# ──────────────────────────────────────────────────────────────────────────────
# Section 15: HTMLレポート機能
# ──────────────────────────────────────────────────────────────────────────────
Write-Section "Section 15: HTMLレポート機能"

Assert-True "HR-01: \$script:taskResults の初期化が存在する" `
    ($sourceContent -match '\$script:taskResults\s*=\s*@\(\)') `
    "\$script:taskResults = @() の初期化が見つかりません"

Assert-True "HR-02: \$script:scriptStartTime の初期化が存在する" `
    ($sourceContent -match '\$script:scriptStartTime') `
    "\$script:scriptStartTime の初期化が見つかりません"

Assert-True "HR-03: Try-Step が \$script:taskResults に PSCustomObject を追加する" `
    ($sourceContent -match '\$script:taskResults\s*\+=\s*\[PSCustomObject\]') `
    "\$script:taskResults += [PSCustomObject] が見つかりません"

Assert-True "HR-04: Try-Step が Stopwatch を使用する" `
    ($sourceContent -match '\[System\.Diagnostics\.Stopwatch\]::StartNew\(\)') `
    "Stopwatch::StartNew() が見つかりません"

Assert-True "HR-05: Try-Step の成功時に Status=OK を記録する" `
    ($sourceContent -match "Status\s*=\s*['""]OK['""]") `
    "Status = 'OK' の記録が見つかりません"

Assert-True "HR-06: Try-Step の失敗時に Status=NG を記録する" `
    ($sourceContent -match "Status\s*=\s*['""]NG['""]") `
    "Status = 'NG' の記録が見つかりません"

Assert-True "HR-07: New-HtmlReport 関数が定義されている" `
    ($sourceContent -match 'function New-HtmlReport') `
    "function New-HtmlReport が見つかりません"

Assert-True "HR-08: New-HtmlReport が WriteAllText で保存する" `
    ($sourceContent -match 'WriteAllText') `
    "[System.IO.File]::WriteAllText が見つかりません"

Assert-True "HR-09: HTMLレポートが logs/ フォルダに保存される" `
    ($sourceContent -match 'PC_Optimizer_Report_') `
    "PC_Optimizer_Report_ のファイル名パターンが見つかりません"

Assert-True "HR-10: 生成後に Start-Process でブラウザを起動する" `
    ($sourceContent -match 'Start-Process \$reportPath') `
    "Start-Process \$reportPath が見つかりません"

Assert-True "HR-11: HTML の基本構造タグが存在する" `
    ($sourceContent -match '<!DOCTYPE html>') `
    "<!DOCTYPE html> が見つかりません"

Assert-True "HR-12: \$script:sysInfo ハッシュテーブルが定義されている" `
    ($sourceContent -match '\$script:sysInfo\s*=\s*@\{') `
    "\$script:sysInfo = @{ が見つかりません"
```

**Step 3: テストを実行して12件追加されることを確認**

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test_PCOptimizer.ps1 2>&1 | tail -5
```

期待出力: `PASS : 87 / 87` — ただし Task 1〜4 実装前なので最初はいくつか FAIL する（意図的）。

---

## Task 6: テスト全PASS確認

**Step 1: Task 1〜4 実装後にテストを再実行**

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test_PCOptimizer.ps1 2>&1 | tail -10
```

期待出力: `PASS : 87 / 87`  `FAIL : 0 / 87`  `All tests PASS`

失敗した場合:
- HR-01〜12 のどれが失敗したか確認
- 対応するコードの実装位置・正規表現パターンを再確認

---

## Task 7: コミット

**Step 1: git status で変更ファイルを確認**

```bash
git status
```

期待: `modified: PC_Optimizer.ps1`、`modified: tests/Test_PCOptimizer.ps1`

**Step 2: コミット**

```bash
git add PC_Optimizer.ps1 tests/Test_PCOptimizer.ps1
git commit -m "feat: HTMLレポート出力機能を追加（v3.3.0）

- \$script:taskResults で全タスクの成否・所要時間を収集
- Try-Step に Stopwatch + PSCustomObject 記録を追加
- New-HtmlReport 関数: システム情報・ディスク解放・タスク結果・スタートアップ情報を含むHTML生成
- 完了後 Start-Process で既定ブラウザ自動起動
- logs/PC_Optimizer_Report_YYYYMMDDHHMM.html に保存
- テスト: 75件 → 87件（+12件、HR-01〜HR-12）"
```

**Step 3: git log で確認**

```bash
git log --oneline -3
```

---

## 注意点

- `[System.Web.HttpUtility]::HtmlEncode()` は .NET 環境で使用可能（PS5.1/7.x 両対応）
- `$finalFreeGB` は Task 4 の挿入位置（ディスク比較後）でのみ存在するため、`New-HtmlReport` の呼び出しは必ず比較ブロックの後に置くこと
- スタートアップ情報はタスク20の結果として `$script:taskResults` に含まれる（別途引数は不要）
- HTML エンコーディング: `[System.IO.File]::WriteAllText` の第3引数に `[System.Text.Encoding]::UTF8`（BOM付き）を使用 — ブラウザ互換性を優先
