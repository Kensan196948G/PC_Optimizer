@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "GUI_SCRIPT=%SCRIPT_DIR%gui\PCOptimizer.Gui.ps1"

where pwsh >nul 2>nul
if %ERRORLEVEL%==0 (
    pwsh -NoProfile -ExecutionPolicy Bypass -STA -File "%GUI_SCRIPT%"
) else (
    powershell -NoProfile -ExecutionPolicy Bypass -STA -File "%GUI_SCRIPT%"
)
