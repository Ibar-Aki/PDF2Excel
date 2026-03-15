# PDF2Excel overview
- Purpose: local Windows tool to convert text PDF tables into Excel using PowerShell, Excel M365 COM, and Power Query.
- Versions: V1 is standard stable conversion; V2 is construction-site raw transfer with Review sheet and multi-page same-column handling.
- Core entrypoints: run_pdf2excel.bat (legacy -> V1), run_pdf2excel_v1.bat, run_pdf2excel_v2.bat, scripts/run_pdf2excel.ps1.
- Key dirs: config/profiles/v1 and v2, samples/v1 and v2, scripts, template, tests, handoff.
- Important constraints: avoid breaking V1, keep BAT ASCII-only, move Japanese UI text into PowerShell, sync handoff/template/profile changes.