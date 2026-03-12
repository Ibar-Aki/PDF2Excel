@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "MANUAL_PATH=%SCRIPT_DIR%docs\user-manual.md"
set "INTERACTIVE_MODE=1"

if not "%~1"=="" (
  set "INTERACTIVE_MODE=0"
  goto RUN_DIRECT
)

:MENU
cls
echo ==========================================
echo   PDF2Excel - Batch Table Converter
echo ==========================================
echo.
echo  Convert PDF tables to Excel in one batch.
echo  If no output path is specified, the file is saved under output.
echo.
echo  [1] Select PDF files and convert
echo  [2] Select a PDF folder and convert
echo  [3] Open the user manual
echo  [4] Open the output folder
echo  [5] Exit
echo.
choice /c 12345 /n /m "Choose an option: "

if errorlevel 5 goto END
if errorlevel 4 goto OPEN_OUTPUT
if errorlevel 3 goto OPEN_MANUAL
if errorlevel 2 goto RUN_FOLDER_DIALOG
if errorlevel 1 goto RUN_DEFAULT

:RUN_DEFAULT
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%scripts\run_pdf2excel.ps1" -PromptForOutputFile
set "EXIT_CODE=%ERRORLEVEL%"
goto SHOW_RESULT

:RUN_FOLDER_DIALOG
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%scripts\run_pdf2excel.ps1" -SelectInputFolder -PromptForOutputFile
set "EXIT_CODE=%ERRORLEVEL%"
goto SHOW_RESULT

:OPEN_MANUAL
if exist "%MANUAL_PATH%" (
  start "" "%MANUAL_PATH%"
) else (
  echo User manual was not found: %MANUAL_PATH%
  pause
)
goto MENU

:OPEN_OUTPUT
if exist "%SCRIPT_DIR%output" (
  start "" "%SCRIPT_DIR%output"
) else (
  echo Output folder was not found: %SCRIPT_DIR%output
  pause
)
goto MENU

:RUN_DIRECT
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%scripts\run_pdf2excel.ps1" %*
set "EXIT_CODE=%ERRORLEVEL%"

:SHOW_RESULT

if not "%EXIT_CODE%"=="0" (
  echo.
  echo PDF2Excel failed. Exit code: %EXIT_CODE%
  echo Check the logs under "%SCRIPT_DIR%logs"
  if "%INTERACTIVE_MODE%"=="1" pause
  goto END
)

echo.
echo Conversion completed.
echo Check the xlsx file and logs under the output and logs folders if needed.
if "%INTERACTIVE_MODE%"=="1" pause

:END
exit /b %EXIT_CODE%


