[CmdletBinding()]
param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$ForwardArgs
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$coreScript = Join-Path $PSScriptRoot 'run_pdf2excel.ps1'
$effectiveArgs = @('-VersionMode', 'v2', '-SecurityMode', 'Secure')
if (-not ($ForwardArgs -contains '-ProfileName') -and -not ($ForwardArgs -contains '-ProfilePath')) {
    $effectiveArgs += @('-ProfileName', 'construction_transfer_poc')
}
$effectiveArgs += $ForwardArgs
& powershell -NoProfile -ExecutionPolicy RemoteSigned -File $coreScript @effectiveArgs
exit $LASTEXITCODE
