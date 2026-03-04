# タスク20 スタートアップ・サービスレポート 実装計画

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** PC_Optimizer.ps1 にタスク20「スタートアップ・サービスレポート」を追加し、スタートアップ登録アプリと不要サービスの稼働状況を読み取り専用で一覧表示する。

**Architecture:** 既存の `Try-Step` ラッパーパターンを踏襲し、`Get-CimInstance Win32_StartupCommand` でスタートアップアプリを、`Get-Service` で不要サービスを取得・表示する。システムへの変更は一切行わない（読み取り専用）。PS3.0/5.1/7.x 全対応。

**Tech Stack:** PowerShell 3.0〜7.x / CIM (Win32_StartupCommand) / Get-Service / 既存 Show・Write-Log・Try-Step 関数

---

### Task 1: テストケースの追加（TDD: 先にテストを書く）

**Files:**
- Modify: `tests/Test_PCOptimizer.ps1`（末尾に追記）

**Step 1: 既存テストファイルの末尾を確認する**

`tests/Test_PCOptimizer.ps1` を開き、現在のテスト数と末尾の構造を確認する。
`$unnecessaryServices` リストのテスト対象は `Fax`, `RetailDemo`, `XblGameSave`, `XboxGipSvc`, `WpcMonSvc`, `DiagTrack`, `MapsBroker`, `lfsvc` の8件。

**Step 2: テストセクションを追加する**

`tests/Test_PCOptimizer.ps1` の末尾（`Write-Host "TOTAL"` などの集計行の直前）に以下を追加する：

```powershell
# ── Task 20: スタートアップ・サービスレポートのロジックテスト ─────
Write-Section "Task 20: Startup / Service Report Logic"

# TC-01: $unnecessaryServices リストが定義されている（PC_Optimizer.ps1 に含まれる）
Assert-True "TC-01: PC_Optimizer.ps1 に unnecessaryServices リストが存在する" `
    ($content -match '\$unnecessaryServices') `
    "スクリプト内に `\$unnecessaryServices` が見つかりません"

# TC-02: リストに Fax が含まれる
Assert-True "TC-02: unnecessaryServices に 'Fax' が含まれる" `
    ($content -match "'Fax'") `
    "'Fax' がリストに見つかりません"

# TC-03: リストに DiagTrack が含まれる
Assert-True "TC-03: unnecessaryServices に 'DiagTrack' が含まれる" `
    ($content -match "'DiagTrack'") `
    "'DiagTrack' がリストに見つかりません"

# TC-04: スタートアップアプリ取得コードが存在する
Assert-True "TC-04: Win32_StartupCommand の呼び出しが存在する" `
    ($content -match 'Win32_StartupCommand') `
    "Win32_StartupCommand が見つかりません"

# TC-05: Get-Service の呼び出しが存在する
Assert-True "TC-05: Get-Service の呼び出しが存在する" `
    ($content -match 'Get-Service') `
    "Get-Service が見つかりません"

# TC-06: タスク20が Try-Step で囲まれている
Assert-True "TC-06: タスク20が Try-Step ラッパーで実装されている" `
    ($content -match 'Try-Step.*スタートアップ') `
    "Try-Step 'スタートアップ...' が見つかりません"

# TC-07: 読み取り専用確認（Set-Service / Stop-Service / Disable-Service が存在しない）
Assert-True "TC-07: Set-Service (無効化コマンド) が存在しない" `
    ($content -notmatch 'Set-Service') `
    "Set-Service が検出されました（変更系コマンドは禁止）"

Assert-True "TC-08: Stop-Service が存在しない" `
    ($content -notmatch 'Stop-Service') `
    "Stop-Service が検出されました（変更系コマンドは禁止）"
```

**Step 3: テストを実行して失敗を確認する（TDD: Red フェーズ）**

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "tests\Test_PCOptimizer.ps1"
```

期待される結果：TC-01〜TC-08 のうち TC-01, TC-02, TC-03, TC-04, TC-05, TC-06 が `[FAIL]`（まだ実装していないため）。TC-07, TC-08 は `[PASS]`（Set-Service / Stop-Service は存在しない）。

**Step 4: コミット（テストのみ）**

```powershell
git add tests/Test_PCOptimizer.ps1
git commit -m "test: Task20 スタートアップ・サービスレポートのテストケースを追加（Red フェーズ）"
```

---

### Task 2: PC_Optimizer.ps1 にタスク20を実装する

**Files:**
- Modify: `PC_Optimizer.ps1`（タスク19の後、再起動チェック前に追記）

**Step 1: 挿入箇所を確認する**

`PC_Optimizer.ps1` を開き、タスク19（`Run-WindowsUpdate` の呼び出し）の終わりと、再起動チェック（`Test-PendingReboot`）の呼び出しの間を探す。
Grep で確認：

```powershell
grep -n "Test-PendingReboot\|Run-WindowsUpdate\|タスク 19\|タスク19" PC_Optimizer.ps1
```

**Step 2: タスク20のコードブロックを追記する**

`Test-PendingReboot` の呼び出し直前（空行を1行入れて）に以下を追記する：

```powershell
# ── タスク 20: スタートアップ・サービスレポート ──────────────────────
Try-Step "スタートアップ・サービスレポート" {
    # ── フェーズ 1: スタートアップ登録アプリ一覧 ──────────────────
    if ($isPS7Plus) {
        Show "📋  ${B}スタートアップ登録アプリ一覧${RB}" Cyan
    } else {
        Show "* スタートアップ登録アプリ一覧" Cyan
    }
    Write-Log "---- [タスク20] スタートアップ登録アプリ一覧 ----"

    try {
        $startupApps = Get-CimInstance -ClassName Win32_StartupCommand -ErrorAction Stop |
            Select-Object Name, Command, Location, User |
            Sort-Object Location, Name

        if (-not $startupApps) {
            Show "  登録なし（スタートアップアプリは見つかりませんでした）" Gray
            Write-Log "  登録なし"
        } else {
            # 表示上限 30 件（ログには全件記録）
            $displayLimit  = 30
            $displayCount  = 0
            $totalCount    = @($startupApps).Count

            foreach ($app in $startupApps) {
                Write-Log ("  名前: {0} | コマンド: {1} | 場所: {2} | ユーザー: {3}" -f `
                    $app.Name, $app.Command, $app.Location, $app.User)

                if ($displayCount -lt $displayLimit) {
                    Show ("  [{0}] {1}" -f $app.Location, $app.Name) White
                    $displayCount++
                }
            }

            if ($totalCount -gt $displayLimit) {
                Show ("  ... 他 {0} 件はログファイルを参照してください" -f ($totalCount - $displayLimit)) Gray
            }
            Write-Log ("  合計: {0} 件" -f $totalCount)
        }
    } catch {
        Show "  スタートアップ情報の取得に失敗しました（スキップ）" Yellow
        Write-ErrorLog "タスク20 スタートアップ取得失敗: $_"
    }

    Show "-----------------------------------------------------" Gray

    # ── フェーズ 2: 不要サービスチェック ──────────────────────────
    if ($isPS7Plus) {
        Show "🔍  ${B}不要サービス稼働チェック${RB}" Cyan
    } else {
        Show "* 不要サービス稼働チェック" Cyan
    }
    Write-Log "---- [タスク20] 不要サービス稼働チェック ----"

    # 照合リスト（読み取り専用チェックのみ。変更・停止は行わない）
    $unnecessaryServices = @(
        'Fax',         # FAXサービス（多くの環境で不使用）
        'RetailDemo',  # 小売デモモード（一般ユーザー不要）
        'XblGameSave', # Xbox ゲームセーブ（非ゲーマー不要）
        'XboxGipSvc',  # Xbox 周辺機器インターフェース
        'WpcMonSvc',   # 保護者機能モニタ（Windows 10 のみ）
        'DiagTrack',   # Connected User Experiences and Telemetry
        'MapsBroker',  # ダウンロード済み地図マネージャ
        'lfsvc'        # 位置情報サービス
    )

    try {
        $runningServices = Get-Service -ErrorAction Stop |
            Where-Object { $_.Status -eq 'Running' }

        $foundCount = 0
        foreach ($svcName in $unnecessaryServices) {
            $match = $runningServices | Where-Object { $_.Name -eq $svcName }
            if ($match) {
                $foundCount++
                Show ("  [確認推奨] {0} が実行中です" -f $svcName) Yellow
                Write-Log ("  [確認推奨] {0} ({1})" -f $svcName, $match.DisplayName)
            }
        }

        if ($foundCount -eq 0) {
            Show "  問題なし（確認推奨サービスの稼働は検出されませんでした）" Green
            Write-Log "  確認推奨サービス: 0 件"
        } else {
            Show ("  合計 {0} 件の確認推奨サービスが実行中です（変更は手動で行ってください）" -f $foundCount) Yellow
            Write-Log ("  確認推奨サービス: {0} 件" -f $foundCount)
        }
    } catch {
        Show "  サービス情報の取得に失敗しました（スキップ）" Yellow
        Write-ErrorLog "タスク20 サービスチェック失敗: $_"
    }
}
```

**Step 3: テストを実行して全件 PASS を確認する（TDD: Green フェーズ）**

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "tests\Test_PCOptimizer.ps1"
```

期待される結果：TC-01〜TC-08 が全て `[PASS]`。

**Step 4: 文字コードテストも実行する**

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "Test_Encoding.ps1"
```

期待される結果：全テスト `PASS`（既存テスト群が引き続き通ること）。

**Step 5: コミット（実装）**

```powershell
git add PC_Optimizer.ps1
git commit -m "feat: タスク20 スタートアップ・サービスレポートを追加（読み取り専用）"
```

---

### Task 3: ドキュメントの更新

**Files:**
- Modify: `docs/変更履歴.md`（先頭に v3.2.0 を追記）
- Modify: `docs/関数リファレンス.md`（タスク20の説明を追記）

**Step 1: 変更履歴に v3.2.0 を追記する**

`docs/変更履歴.md` の先頭（`## [3.1.0]` の直前）に以下を追記する：

```markdown
## [3.2.0] - 2026-03-04

### 追加

- **タスク20: スタートアップ・サービスレポート（表示のみ）**
  - `Win32_StartupCommand` でスタートアップ登録アプリを一覧表示（最大30件、全件はログ参照）
  - 不要サービス8件（Fax / RetailDemo / XblGameSave 等）の稼働チェックと表示
  - 読み取り専用：システムへの変更は一切行わない
  - PS3.0 / PS5.1 / PS7.x 全対応（PS7+ で絵文字表示）

---
```

**Step 2: 関数リファレンスにタスク20を追記する**

`docs/関数リファレンス.md` のタスク一覧セクションにタスク20の説明を追加する。

**Step 3: コミット（ドキュメント）**

```powershell
git add docs/変更履歴.md docs/関数リファレンス.md
git commit -m "docs: v3.2.0 タスク20のドキュメントを更新"
```

---

### Task 4: 最終確認・動作確認チェックリスト

**Step 1: 実装コードのセルフチェック**

以下を `PC_Optimizer.ps1` で目視確認する：

- [ ] `$unnecessaryServices` 変数が `Try-Step` ブロック内に定義されている
- [ ] `Set-Service`, `Stop-Service`, `Disable-Service` が存在しない（読み取り専用確認）
- [ ] `Write-Log` が全情報を記録している
- [ ] `Write-ErrorLog` がエラー時に呼ばれる
- [ ] `-ErrorAction Stop` が `try-catch` と組み合わせて使われている
- [ ] `$isPS7Plus` フラグで絵文字/ASCIIが分岐している

**Step 2: テスト全実行**

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "tests\Test_PCOptimizer.ps1"
powershell -NoProfile -ExecutionPolicy Bypass -File "Test_Encoding.ps1"
```

両方とも全テスト `PASS` であることを確認する。

**Step 3: 設計ドキュメントのコミット**

```powershell
git add docs/plans/
git commit -m "docs: タスク20 設計ドキュメントと実装計画を追加"
```

---

## 実装チェックリスト（完了基準）

- [ ] `tests/Test_PCOptimizer.ps1` に TC-01〜TC-08 が追加され全 PASS
- [ ] `PC_Optimizer.ps1` にタスク20のコードブロックが追加されている
- [ ] `Win32_StartupCommand` でスタートアップ一覧が取得・表示される
- [ ] `$unnecessaryServices` リスト（8件）が定義されている
- [ ] `Get-Service` で照合し Running のもののみ警告表示される
- [ ] `Set-Service` / `Stop-Service` / `Disable-Service` が存在しない
- [ ] `Write-Log` / `Write-ErrorLog` が適切に使われている
- [ ] `Test_Encoding.ps1` 全 PASS
- [ ] `docs/変更履歴.md` に v3.2.0 が記録されている
- [ ] `docs/関数リファレンス.md` にタスク20が記録されている
- [ ] git commit が Task1〜Task4 で計4コミット作成されている
