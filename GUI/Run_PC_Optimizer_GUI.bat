@echo off
setlocal EnableExtensions

set "SCRIPT_DIR=%~dp0"
set "GUI_SCRIPT=%SCRIPT_DIR%PC_Optimizer_GUI.ps1"
set "POWERSHELL_EXE="

if not exist "%GUI_SCRIPT%" (
  echo [ERROR] GUI script not found: "%GUI_SCRIPT%"
  exit /b 1
)

if exist "%ProgramFiles%\PowerShell\7\pwsh.exe" (
  set "POWERSHELL_EXE=%ProgramFiles%\PowerShell\7\pwsh.exe"
) else if exist "%ProgramFiles%\PowerShell\7-preview\pwsh.exe" (
  set "POWERSHELL_EXE=%ProgramFiles%\PowerShell\7-preview\pwsh.exe"
) else (
  for /f "delims=" %%I in ('where pwsh 2^>nul') do (
    set "POWERSHELL_EXE=%%I"
    goto :run_gui
  )
  set "POWERSHELL_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
)

:run_gui
if not exist "%POWERSHELL_EXE%" (
  echo [ERROR] PowerShell executable not found: "%POWERSHELL_EXE%"
  exit /b 1
)

"%POWERSHELL_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%GUI_SCRIPT%" %*
