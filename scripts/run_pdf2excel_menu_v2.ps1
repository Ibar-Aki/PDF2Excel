param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$ForwardArgs,
    [switch]$ForceMenu
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$menuScript = Join-Path $PSScriptRoot 'run_pdf2excel_menu.ps1'
$invokeArgs = @('-VersionMode', 'v2')
if ($ForceMenu) {
    $invokeArgs += '-ForceMenu'
}
$invokeArgs += $ForwardArgs
& powershell -NoProfile -ExecutionPolicy RemoteSigned -File $menuScript @invokeArgs
exit $LASTEXITCODE
