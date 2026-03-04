@echo off
setlocal EnableExtensions EnableDelayedExpansion

:: スクリプトのパスを取得
set "SCRIPT_DIR=%~dp0"
set "PS_SCRIPT=%SCRIPT_DIR%PC_Optimizer.ps1"

if /I "%~1"=="--elevated" shift

set "FORWARD_ARGS="
:buildArgs
if "%~1"=="" goto argsDone
set "ARG=%~1"
set "ARG=!ARG:\"=\\\"!"
set "FORWARD_ARGS=!FORWARD_ARGS! "!ARG!""
shift
goto buildArgs
:argsDone

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
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -ArgumentList '--elevated!FORWARD_ARGS!' -Verb RunAs"
    exit /b
)

:: PowerShell スクリプトを実行
"%PWSH_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" %FORWARD_ARGS%
pause