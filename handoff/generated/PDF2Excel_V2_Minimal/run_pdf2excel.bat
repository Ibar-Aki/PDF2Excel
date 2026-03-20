@echo off
setlocal
chcp 65001 >nul
set "SCRIPT_DIR=%~dp0"
if "%~1"=="" (
    powershell -NoProfile -ExecutionPolicy RemoteSigned -File "%SCRIPT_DIR%scripts\run_pdf2excel_menu.ps1" -VersionMode v2 -RunScriptPath "%SCRIPT_DIR%scripts\run_pdf2excel.ps1" -PassCoreDefaults -DefaultSecurityMode Secure -DefaultProfileName construction_transfer_poc -ForceMenu
) else (
    powershell -NoProfile -ExecutionPolicy RemoteSigned -File "%SCRIPT_DIR%scripts\run_pdf2excel_menu.ps1" -VersionMode v2 -RunScriptPath "%SCRIPT_DIR%scripts\run_pdf2excel.ps1" -PassCoreDefaults -DefaultSecurityMode Secure -DefaultProfileName construction_transfer_poc %*
)
exit /b %ERRORLEVEL%
