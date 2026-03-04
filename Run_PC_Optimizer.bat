@echo off
:: スクリプトのパスを取得
set "SCRIPT_DIR=%~dp0"
set "PS_SCRIPT=%SCRIPT_DIR%PC_Optimizer.ps1"

:: PowerShell のパスを自動検出（Win10/11 デフォルト）
set "PWSH_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
:: PowerShell 7 が存在すれば優先使用
where pwsh >nul 2>nul
if %ERRORLEVEL%==0 (
    set "PWSH_EXE=pwsh"
)

:: 管理者権限がなければ UAC で昇格して再起動
net session >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo 管理者権限で再起動しています...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

:: PowerShell スクリプトを実行
"%PWSH_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%"
pause
