[CmdletBinding()]
param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$ForwardArgs
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoCoreScript = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..\scripts\run_pdf2excel.ps1'))
$packageCoreScript = Join-Path $PSScriptRoot 'run_pdf2excel.ps1'
$coreScript = if (Test-Path -LiteralPath $repoCoreScript) { $repoCoreScript } else { $packageCoreScript }

$profilePathCandidates = @(
    [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\config\profiles\default.json')),
    [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\config\profiles\v1\default.json'))
)
$defaultProfilePath = $profilePathCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1

$effectiveArgs = @('-VersionMode', 'v1')
$joinedArgs = ' ' + ($ForwardArgs -join ' ') + ' '
if ($defaultProfilePath -and $joinedArgs -notmatch '\s-Profile(Name|Path)\s') {
    $effectiveArgs += @('-ProfilePath', $defaultProfilePath)
}
$effectiveArgs += $ForwardArgs

& powershell -NoProfile -ExecutionPolicy RemoteSigned -File $coreScript @effectiveArgs
exit $LASTEXITCODE
