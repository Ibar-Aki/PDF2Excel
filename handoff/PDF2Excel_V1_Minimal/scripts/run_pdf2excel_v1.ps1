[CmdletBinding()]
param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$ForwardArgs
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$coreScript = Join-Path $PSScriptRoot 'run_pdf2excel.ps1'
& powershell -NoProfile -ExecutionPolicy Bypass -File $coreScript -VersionMode v1 @ForwardArgs
exit $LASTEXITCODE
