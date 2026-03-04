# Test_Encoding.ps1 — 文字コード検証テスト
# 対象:
#   1. Run_PC_Optimizer.bat  — Shift-JIS (CP932)・BOM なし・日本語UI文字列
#   2. PC_Optimizer.ps1      — UTF-8 BOM 付き・日本語UI文字列・構文チェック

$pass = 0
$fail = 0

function Test-Assert ($label, $result, $detail = "") {
    if ($result) {
        Write-Host "  [PASS] $label" -ForegroundColor Green
        $script:pass++
    } else {
        Write-Host "  [FAIL] $label$(if ($detail) { ' — ' + $detail })" -ForegroundColor Red
        $script:fail++
    }
}

# ============================================================
#  BAT ファイル検証
# ============================================================
$batPath = Join-Path $PSScriptRoot "Run_PC_Optimizer.bat"

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  [1/2] Run_PC_Optimizer.bat 文字コード検証" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  対象: $batPath"
Write-Host ""

if (-not (Test-Path $batPath)) {
    Write-Host "  [ERROR] ファイルが見つかりません: $batPath" -ForegroundColor Red
    exit 1
}

$batBytes = [System.IO.File]::ReadAllBytes($batPath)
Write-Host "  ファイルサイズ: $($batBytes.Length) bytes"
Write-Host ""

# Test BAT-1: BOM チェック
Write-Host "[ BAT Test 1 ] BOM チェック" -ForegroundColor Yellow
$hasUtf8Bom = ($batBytes.Length -ge 3 -and $batBytes[0] -eq 0xEF -and $batBytes[1] -eq 0xBB -and $batBytes[2] -eq 0xBF)
$hasUtf16Le = ($batBytes.Length -ge 2 -and $batBytes[0] -eq 0xFF -and $batBytes[1] -eq 0xFE)
$hasUtf16Be = ($batBytes.Length -ge 2 -and $batBytes[0] -eq 0xFE -and $batBytes[1] -eq 0xFF)
Test-Assert "UTF-8 BOM なし (EF BB BF)"  (-not $hasUtf8Bom)
Test-Assert "UTF-16 LE BOM なし (FF FE)" (-not $hasUtf16Le)
Test-Assert "UTF-16 BE BOM なし (FE FF)" (-not $hasUtf16Be)
Write-Host ""

# Test BAT-1b: CRLF 行末コードの確認（cmd.exe 必須要件）
Write-Host "[ BAT Test 1b ] 行末コードの確認（CRLF 必須）" -ForegroundColor Yellow
$batText    = [System.Text.Encoding]::GetEncoding(932).GetString($batBytes)
$crlfCount  = ([regex]::Matches($batText, "`r`n")).Count
$lfOnlyCount = ([regex]::Matches($batText, "(?<!`r)`n")).Count
Test-Assert "CRLF 行が存在する（cmd.exe 互換）" ($crlfCount -gt 0)         "CRLF 行数: $crlfCount"
Test-Assert "LF のみ行がない（LF のみは cmd.exe でブロック解析失敗）" ($lfOnlyCount -eq 0) "LF のみ行数: $lfOnlyCount"
Write-Host ""

# Test BAT-2: Shift-JIS デコード
Write-Host "[ BAT Test 2 ] Shift-JIS (CP932) デコード" -ForegroundColor Yellow
$sjis       = [System.Text.Encoding]::GetEncoding(932)
$batDecoded = $sjis.GetString($batBytes)
Test-Assert "CP932 デコード成功（例外なし）" ($batDecoded -ne $null -and $batDecoded.Length -gt 0)
Write-Host ""

# Test BAT-3: 日本語 UI 文字列の存在
Write-Host "[ BAT Test 3 ] 日本語 UI 文字列の存在確認" -ForegroundColor Yellow
@(
    "管理者権限で再起動しています...",
    "管理者権限がなければ",
    "PowerShell 7 が存在すれば優先使用",
    "スクリプトのパスを取得",
    "PowerShell スクリプトを実行"
) | ForEach-Object { Test-Assert "「$_」が含まれる" ($batDecoded -match [regex]::Escape($_)) }
Write-Host ""

# Test BAT-4: 旧英語 UI 文字列の除去確認
Write-Host "[ BAT Test 4 ] 旧英語 UI 文字列が削除されていること" -ForegroundColor Yellow
@(
    "Relaunching with Administrator privileges",
    "Auto-detect PowerShell path",
    "Relaunch as Administrator if not elevated",
    "Run the PowerShell script",
    "Prefer PowerShell 7 if available",
    "Get the batch file path"
) | ForEach-Object { Test-Assert "「$_」が存在しない" ($batDecoded -notmatch [regex]::Escape($_)) }
Write-Host ""

# Test BAT-5: Shift-JIS バイト列の直接検証
Write-Host "[ BAT Test 5 ] Shift-JIS バイト列の直接検証" -ForegroundColor Yellow
$testStr  = "管理者権限で再起動しています..."
$testHex  = ($sjis.GetBytes($testStr) | ForEach-Object { $_.ToString('X2') }) -join ' '
$fileHex  = ($batBytes | ForEach-Object { $_.ToString('X2') }) -join ' '
Test-Assert "「管理者権限で再起動しています...」の SJIS バイト列が存在する" ($fileHex.Contains($testHex)) "期待バイト: $testHex"
Write-Host ""

# ============================================================
#  PS1 ファイル検証
# ============================================================
$ps1Path = Join-Path $PSScriptRoot "PC_Optimizer.ps1"

Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  [2/2] PC_Optimizer.ps1 文字コード検証" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  対象: $ps1Path"
Write-Host ""

if (-not (Test-Path $ps1Path)) {
    Write-Host "  [ERROR] ファイルが見つかりません: $ps1Path" -ForegroundColor Red
    exit 1
}

$ps1Bytes = [System.IO.File]::ReadAllBytes($ps1Path)
Write-Host "  ファイルサイズ: $($ps1Bytes.Length) bytes"
Write-Host ""

# Test PS1-1: UTF-8 BOM の存在確認
Write-Host "[ PS1 Test 1 ] UTF-8 BOM 付き確認（PS5.1 互換要件）" -ForegroundColor Yellow
$ps1HasBom    = ($ps1Bytes.Length -ge 3 -and $ps1Bytes[0] -eq 0xEF -and $ps1Bytes[1] -eq 0xBB -and $ps1Bytes[2] -eq 0xBF)
$ps1HasUtf16  = ($ps1Bytes.Length -ge 2 -and ($ps1Bytes[0] -eq 0xFF -or $ps1Bytes[0] -eq 0xFE))
Test-Assert "UTF-8 BOM あり (EF BB BF)" $ps1HasBom  "先頭バイト: 0x$($ps1Bytes[0].ToString('X2')) 0x$($ps1Bytes[1].ToString('X2')) 0x$($ps1Bytes[2].ToString('X2'))"
Test-Assert "UTF-16 でないこと"         (-not $ps1HasUtf16)
Write-Host ""

# Test PS1-2: UTF-8 デコード
Write-Host "[ PS1 Test 2 ] UTF-8 デコード" -ForegroundColor Yellow
$utf8        = [System.Text.Encoding]::UTF8
$ps1Decoded  = $utf8.GetString($ps1Bytes)
Test-Assert "UTF-8 デコード成功（例外なし）" ($ps1Decoded -ne $null -and $ps1Decoded.Length -gt 0)
Write-Host ""

# Test PS1-3: 日本語 UI 文字列の存在
Write-Host "[ PS1 Test 3 ] 日本語 UI 文字列の存在確認" -ForegroundColor Yellow
$requiredUiStrings = @(
    "警告: PowerShell 3.0 以上を推奨します。",
    "現在のバージョン:",
    "一部の機能が正常に動作しない場合があります。",
    "ホスト:",
    "ユーザー:",
    "メモリ:",
    "ディスク",
    "進捗:",
    "開始...",
    "完了",
    "失敗",
    "一時ファイルの削除",
    "Prefetch・更新キャッシュの削除",
    "配信最適化キャッシュの削除",
    "Windows Update キャッシュの削除",
    "エラーレポート・ログ・不要キャッシュの削除",
    "OneDrive / Teams / Office キャッシュの削除",
    "ごみ箱を空にする",
    "DNS キャッシュのクリア",
    "ディスクの最適化（SSD: TRIM / HDD: デフラグ）",
    "SSD ヘルスチェック（SMART）",
    "SSD が検出されませんでした。",
    "UsoClient.exe が見つかりません。Windows Update をスキップします。",
    "Windows Update の実行",
    "サムネイルキャッシュの削除",
    "Microsoft Store キャッシュのクリア",
    "Windows イベントログのクリア",
    "システムファイルの整合性チェック・修復（SFC）",
    "Windows コンポーネントストアの診断（DISM）",
    "電源プランの最適化",
    "Microsoft 365 の更新確認・適用",
    "Microsoft 365 を検出しました。",
    "Microsoft 365 の更新を確認・適用しますか？",
    "利用可能な更新",
    "Windows Update を確認中",
    "これらの更新を適用しますか？",
    "ディスク解放結果:",
    "最適化前:",
    "解放容量:",
    "再起動が必要です。今すぐ再起動しますか？",
    "PC 最適化が完了しました。",
    "Enter キーを押して終了"
)

$requiredUiStrings | ForEach-Object { Test-Assert "「$_」が含まれる" ($ps1Decoded -match [regex]::Escape($_)) }

# 文言の完全一致ではなく、機能要件（6ブラウザ対応）を検証する
Test-Assert "「ブラウザキャッシュの削除（...）」見出しが含まれる" ($ps1Decoded -match [regex]::Escape("ブラウザキャッシュの削除（"))
@("Chrome", "Edge", "Firefox", "Brave", "Opera", "Vivaldi") | ForEach-Object {
    Test-Assert "ブラウザ名「$_」が含まれる" ($ps1Decoded -match [regex]::Escape($_))
}
Write-Host ""

# Test PS1-4: 旧英語 UI 文字列の除去確認
Write-Host "[ PS1 Test 4 ] 旧英語 UI 文字列が削除されていること" -ForegroundColor Yellow
@(
    "Warning: PowerShell 3.0 or higher is recommended.",
    "Current version:",
    "Some features may not work correctly.",
    "Host:",
    "Memory:",
    "Progress:",
    "start...",
    " completed",
    " failed",
    "Delete temp files",
    "Delete Prefetch",
    "Delete Delivery Optimization cache",
    "Delete Windows Update cache",
    "Delete error reports",
    "Delete OneDrive",
    "Empty Recycle Bin",
    "Clear DNS cache",
    "Optimize disk",
    "SSD health check",
    "No SSD detected.",
    "UsoClient.exe not found.",
    "Run Windows Update",
    "A restart is required. Restart now?",
    "PC optimization completed.",
    "Press Enter to exit"
) | ForEach-Object { Test-Assert "「$_」が存在しない" ($ps1Decoded -notmatch [regex]::Escape($_)) }
Write-Host ""

# Test PS1-5: PowerShell 構文チェック
Write-Host "[ PS1 Test 5 ] PowerShell 構文チェック" -ForegroundColor Yellow
try {
    $tokens = $null
    $errors = $null
    $null = [System.Management.Automation.Language.Parser]::ParseFile(
        $ps1Path,
        [ref]$tokens,
        [ref]$errors
    )
    Test-Assert "構文エラーなし（パースOK）" ($errors.Count -eq 0) "エラー数: $($errors.Count)"
} catch {
    Test-Assert "構文チェック実行成功" $false "例外: $_"
}
Write-Host ""

# ============================================================
#  結果サマリー
# ============================================================
Write-Host "================================================================" -ForegroundColor Cyan
$total = $pass + $fail
if ($fail -eq 0) {
    Write-Host "  結果: 全 $total テスト PASS" -ForegroundColor Green
} else {
    Write-Host "  結果: $pass / $total PASS  |  $fail FAIL" -ForegroundColor Red
}
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

exit $fail
