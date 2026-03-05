@echo off
setlocal EnableExtensions EnableDelayedExpansion

:: スクリプトのパスを取得
set "SCRIPT_DIR=%~dp0"
set "PS_SCRIPT=%SCRIPT_DIR%PC_Optimizer.ps1"
set "ENV_FILE=%SCRIPT_DIR%.env"

:: .env ファイル読み込み（API キー・設定の引き継ぎ用）
if exist "%ENV_FILE%" (
    for /f "usebackq tokens=1,* delims==" %%A in ("%ENV_FILE%") do (
        set "ENV_KEY=%%~A"
        set "ENV_VAL=%%~B"
        for /f "tokens=* delims= " %%K in ("!ENV_KEY!") do set "ENV_KEY=%%~K"
        for /f "tokens=* delims= " %%V in ("!ENV_VAL!") do set "ENV_VAL=%%~V"
        if /I "!ENV_KEY:~0,7!"=="export " set "ENV_KEY=!ENV_KEY:~7!"
        if /I "!ENV_KEY:~0,4!"=="set " set "ENV_KEY=!ENV_KEY:~4!"
        if defined ENV_KEY (
            if not "!ENV_KEY:~0,1!"=="#" (
                if not "!ENV_KEY:~0,1!"==";" (
                    if "!ENV_VAL:~0,1!"==""" if "!ENV_VAL:~-1!"==""" set "ENV_VAL=!ENV_VAL:~1,-1!"
                    set "!ENV_KEY!=!ENV_VAL!"
                )
            )
        )
    )
)

set "_ELEVATED=0"
if /I "%~1"=="--elevated" (
    set "_ELEVATED=1"
    shift
)

:: PowerShell 7 が存在すれば優先使用
set "PWSH_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
where pwsh >nul 2>nul
if %ERRORLEVEL%==0 (
    set "PWSH_EXE=pwsh"
)

:: 管理者権限がなければ UAC で昇格要求
net session >nul 2>&1
if %ERRORLEVEL% NEQ 0 if "%_ELEVATED%"=="0" (
    echo 管理者権限で再起動しています...
    powershell -NoProfile -Command "Start-Process -FilePath 'cmd.exe' -ArgumentList '/c','""%~f0" --elevated %*"' -Verb RunAs"
    exit /b
)
if %ERRORLEVEL% NEQ 0 if "%_ELEVATED%"=="1" (
    echo 管理者権限がなければ実行できません。
    exit /b 1
)

:: PowerShell スクリプトを実行
"%PWSH_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" %*
pause
