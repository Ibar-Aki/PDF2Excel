@echo off
setlocal
chcp 65001 >nul
set "SCRIPT_DIR=%~dp0"
if "%~1"=="" (
    powershell -NoProfile -ExecutionPolicy RemoteSigned -File "%SCRIPT_DIR%scripts\run_pdf2excel_menu_v2.ps1" -ForceMenu
) else (
    powershell -NoProfile -ExecutionPolicy RemoteSigned -File "%SCRIPT_DIR%scripts\run_pdf2excel_menu_v2.ps1" %*
)
exit /b %ERRORLEVEL%
