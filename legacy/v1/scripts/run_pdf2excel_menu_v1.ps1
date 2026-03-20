param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$ForwardArgs
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoMenuScript = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..\scripts\run_pdf2excel_menu.ps1'))
$packageMenuScript = Join-Path $PSScriptRoot 'run_pdf2excel_menu.ps1'
$menuScript = if (Test-Path -LiteralPath $repoMenuScript) { $repoMenuScript } else { $packageMenuScript }

$repoRunScript = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..\scripts\run_pdf2excel.ps1'))
$packageRunScript = Join-Path $PSScriptRoot 'run_pdf2excel.ps1'
$runScript = if (Test-Path -LiteralPath $repoRunScript) { $repoRunScript } else { $packageRunScript }

$profileDirCandidates = @(
    [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\config\profiles\v1')),
    [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\config\profiles'))
)
$profileDir = $profileDirCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1

$manualPathCandidates = @(
    [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..\docs\maintenance-guide.md')),
    [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\HANDOFF_README.md'))
)
$manualPath = $manualPathCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1

$invokeArgs = @(
    '-VersionMode', 'v1',
    '-RunScriptPath', $runScript,
    '-PassCoreDefaults',
    '-SystemLabelOverride', 'PDF2Excel VER1 Legacy - 保守用標準変換',
    '-CompletionSheetMessageOverride', 'Result / Errors / Summary を確認してください。'
)
if ($profileDir) {
    $invokeArgs += @('-ProfileDirOverride', $profileDir, '-ProfileBaseDirOverride', $profileDir)
}
if ($manualPath) {
    $invokeArgs += @('-ManualPathOverride', $manualPath)
}
$invokeArgs += $ForwardArgs

& powershell -NoProfile -ExecutionPolicy RemoteSigned -File $menuScript @invokeArgs
exit $LASTEXITCODE
