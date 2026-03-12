@echo off
setlocal

set "SCRIPT_DIR=%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%scripts\run_pdf2excel.ps1" %*
set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" (
  echo PDF2Excel failed with exit code %EXIT_CODE%.
)

exit /b %EXIT_CODE%
