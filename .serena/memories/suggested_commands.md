# Suggested commands
- Git status: `git status --short --branch`
- Run unit tests: `powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\run_unit_tests.ps1`
- Run full integration tests: `powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\run_integration_tests.ps1`
- Run one integration case: `powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\run_integration_tests.ps1 -CaseName 'V2 2ページ同一列の変換'`
- Run V1: `powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\run_pdf2excel_v1.ps1 ...`
- Run V2: `powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\run_pdf2excel_v2.ps1 ...`
- Rebuild templates: `powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\build_excel_template.ps1 -TemplatePath .\template\PDF2Excel_V1_Converter.xlsm -TemplateVariant v1` and same for v2.
- Rebuild samples: `powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\build_sample_pdfs.ps1`