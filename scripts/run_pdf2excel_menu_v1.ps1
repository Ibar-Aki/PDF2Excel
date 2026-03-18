param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$ForwardArgs
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$menuScript = Join-Path $PSScriptRoot 'run_pdf2excel_menu.ps1'
& powershell -NoProfile -ExecutionPolicy RemoteSigned -File $menuScript -VersionMode v1 @ForwardArgs
exit $LASTEXITCODE
