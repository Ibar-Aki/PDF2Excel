param(
    [ValidateSet('all', 'v1', 'v2')]
    [string]$TargetVersion = 'all'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
. (Join-Path $PSScriptRoot 'pdf2excel.common.ps1')

$handoffRoot = Join-Path $projectRoot 'handoff'
Ensure-Directory -Path $handoffRoot

$packageDefinitions = @(
    [pscustomobject]@{
        VersionMode = 'v1'
        PackageRoot = Join-Path $handoffRoot 'PDF2Excel_V1_Minimal'
        ZipPath = Join-Path $handoffRoot 'PDF2Excel_V1_Minimal.zip'
        TemplateFile = 'PDF2Excel_V1_Converter.xlsm'
        RunBat = 'run_pdf2excel_v1.bat'
        MenuScript = 'run_pdf2excel_menu_v1.ps1'
        RunScript = 'run_pdf2excel_v1.ps1'
    },
    [pscustomobject]@{
        VersionMode = 'v2'
        PackageRoot = Join-Path $handoffRoot 'PDF2Excel_V2_Minimal'
        ZipPath = Join-Path $handoffRoot 'PDF2Excel_V2_Minimal.zip'
        TemplateFile = 'PDF2Excel_V2_Converter.xlsm'
        RunBat = 'run_pdf2excel_v2.bat'
        MenuScript = 'run_pdf2excel_menu_v2.ps1'
        RunScript = 'run_pdf2excel_v2.ps1'
    }
)

if ($TargetVersion -ne 'all') {
    $packageDefinitions = @($packageDefinitions | Where-Object VersionMode -eq $TargetVersion)
}

foreach ($package in $packageDefinitions) {
    Reset-Directory -Path $package.PackageRoot

    foreach ($relativeDir in @('scripts', 'template', "config\profiles\$($package.VersionMode)", 'input', 'output', 'output\runtime', 'logs')) {
        Ensure-Directory -Path (Join-Path $package.PackageRoot $relativeDir)
    }

    $filesToCopy = @(
        @{ Source = Join-Path $projectRoot $package.RunBat; Destination = Join-Path $package.PackageRoot 'run_pdf2excel.bat' },
        @{ Source = Join-Path $projectRoot 'scripts\run_pdf2excel.ps1'; Destination = Join-Path $package.PackageRoot 'scripts\run_pdf2excel.ps1' },
        @{ Source = Join-Path $projectRoot ("scripts\{0}" -f $package.RunScript); Destination = Join-Path $package.PackageRoot ("scripts\{0}" -f $package.RunScript) },
        @{ Source = Join-Path $projectRoot ("scripts\{0}" -f $package.MenuScript); Destination = Join-Path $package.PackageRoot ("scripts\{0}" -f $package.MenuScript) },
        @{ Source = Join-Path $projectRoot 'scripts\run_pdf2excel_menu.ps1'; Destination = Join-Path $package.PackageRoot 'scripts\run_pdf2excel_menu.ps1' },
        @{ Source = Join-Path $projectRoot 'scripts\pdf2excel.common.ps1'; Destination = Join-Path $package.PackageRoot 'scripts\pdf2excel.common.ps1' },
        @{ Source = Join-Path $projectRoot ("template\{0}" -f $package.TemplateFile); Destination = Join-Path $package.PackageRoot ("template\{0}" -f $package.TemplateFile) },
        @{ Source = Join-Path $projectRoot 'handoff\HANDOFF_README_SOURCE.md'; Destination = Join-Path $package.PackageRoot 'HANDOFF_README.md' }
    )

    foreach ($file in $filesToCopy) {
        Copy-Item -LiteralPath $file.Source -Destination $file.Destination -Force
    }

    Get-ChildItem -LiteralPath (Join-Path $projectRoot ("config\profiles\{0}" -f $package.VersionMode)) -Filter '*.json' -File | ForEach-Object {
        Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $package.PackageRoot ("config\profiles\{0}\{1}" -f $package.VersionMode, $_.Name)) -Force
    }

    foreach ($gitkeepPath in @(
        (Join-Path $package.PackageRoot 'input\.gitkeep'),
        (Join-Path $package.PackageRoot 'output\.gitkeep'),
        (Join-Path $package.PackageRoot 'output\runtime\.gitkeep'),
        (Join-Path $package.PackageRoot 'logs\.gitkeep')
    )) {
        New-EmptyFile -Path $gitkeepPath
    }

    if (Test-Path -LiteralPath $package.ZipPath) {
        Remove-Item -LiteralPath $package.ZipPath -Force
    }

    Compress-Archive -Path (Join-Path $package.PackageRoot '*') -DestinationPath $package.ZipPath
    Write-Host ("配布フォルダ ({0}): {1}" -f (Get-VersionDisplayName -VersionMode $package.VersionMode), $package.PackageRoot)
    Write-Host ("配布ZIP ({0}): {1}" -f (Get-VersionDisplayName -VersionMode $package.VersionMode), $package.ZipPath)
}
