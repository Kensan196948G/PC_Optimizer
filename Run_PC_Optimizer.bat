@echo off
setlocal EnableExtensions EnableDelayedExpansion

:: Resolve script directory and target PowerShell script path
set "SCRIPT_DIR=%~dp0"
set "PS_SCRIPT=%SCRIPT_DIR%PC_Optimizer.ps1"
set "ENV_FILE=%SCRIPT_DIR%.env"

:: Load .env (KEY=VALUE) so elevated run can reuse API keys and config.
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
                    if "!ENV_VAL:~0,1!"=="^"" if "!ENV_VAL:~-1!"=="^"" set "ENV_VAL=!ENV_VAL:~1,-1!"
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

:: Prefer Windows PowerShell 5.1 (default on Win10/11)
set "PWSH_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
:: Use PowerShell 7 if installed
where pwsh >nul 2>nul
if %ERRORLEVEL%==0 (
    set "PWSH_EXE=pwsh"
)

:: Elevate once via UAC when not running as administrator
net session >nul 2>&1
if %ERRORLEVEL% NEQ 0 if "%_ELEVATED%"=="0" (
    echo Requesting administrator privileges...
    powershell -NoProfile -Command "Start-Process -FilePath 'cmd.exe' -ArgumentList '/c','\"\"%~f0\" --elevated %*\"' -Verb RunAs"
    exit /b
)
if %ERRORLEVEL% NEQ 0 if "%_ELEVATED%"=="1" (
    echo Failed to acquire administrator privileges.
    exit /b 1
)

:: Run the PowerShell script
"%PWSH_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" %*
pause
