param(
    [ValidateSet('all', 'v1', 'v2')]
    [string]$TargetVersion = 'v2'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
. (Join-Path $PSScriptRoot 'pdf2excel.common.ps1')

[Console]::InputEncoding = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding = [Console]::OutputEncoding

function Sync-VbaModuleEncodingMirror {
    param(
        [Parameter(Mandatory = $true)][string]$Utf8Path,
        [Parameter(Mandatory = $true)][string]$ShiftJisPath
    )

    if (-not (Test-Path -LiteralPath $Utf8Path)) {
        throw "VBA モジュールファイルが見つかりません: $Utf8Path"
    }

    $moduleText = Get-Content -LiteralPath $Utf8Path -Raw -Encoding UTF8
    $encoding = [System.Text.Encoding]::GetEncoding(932)
    [System.IO.File]::WriteAllBytes($ShiftJisPath, $encoding.GetBytes($moduleText))
}

$handoffRoot = Join-Path $projectRoot 'handoff'
Ensure-Directory -Path $handoffRoot

Sync-VbaModuleEncodingMirror -Utf8Path (Join-Path $projectRoot 'template\vba\PDF2ExcelMacros.bas') -ShiftJisPath (Join-Path $projectRoot 'template\vba\PDF2ExcelMacros.sjis.bas')
Sync-VbaModuleEncodingMirror -Utf8Path (Join-Path $projectRoot 'template\vba\PDF2ExcelTemplateBuilder.bas') -ShiftJisPath (Join-Path $projectRoot 'template\vba\PDF2ExcelTemplateBuilder.sjis.bas')

$packageDefinitions = @(
    [pscustomobject]@{
        VersionMode = 'v1'
        PackageRoot = Join-Path $handoffRoot 'PDF2Excel_V1_Minimal'
        ZipPath = Join-Path $handoffRoot 'PDF2Excel_V1_Minimal.zip'
        TemplateFile = 'PDF2Excel_V1_Converter.xlsm'
        RunBat = 'run_pdf2excel_v1.bat'
        MenuScript = 'run_pdf2excel_menu_v1.ps1'
        RunScript = 'run_pdf2excel_v1.ps1'
        ReadmeSource = 'handoff\HANDOFF_README_SOURCE.md'
    },
    [pscustomobject]@{
        VersionMode = 'v2'
        PackageRoot = Join-Path $handoffRoot 'PDF2Excel_V2_Minimal'
        ZipPath = Join-Path $handoffRoot 'PDF2Excel_V2_Minimal.zip'
        TemplateFile = 'PDF2Excel_V2_Converter.xlsm'
        RunBat = 'run_pdf2excel_v2.bat'
        MenuScript = $null
        RunScript = $null
        ReadmeSource = 'handoff\HANDOFF_README_V2_SOURCE.md'
    }
)

if ($TargetVersion -ne 'all') {
    $packageDefinitions = @($packageDefinitions | Where-Object VersionMode -eq $TargetVersion)
}

foreach ($package in $packageDefinitions) {
    Reset-Directory -Path $package.PackageRoot

    foreach ($relativeDir in @('scripts', 'template', 'template\vba', "config\profiles\$($package.VersionMode)", 'input', 'output', 'output\runtime', 'logs')) {
        Ensure-Directory -Path (Join-Path $package.PackageRoot $relativeDir)
    }

    $filesToCopy = @(
        @{ Source = Join-Path $projectRoot $package.RunBat; Destination = Join-Path $package.PackageRoot 'run_pdf2excel.bat' },
        @{ Source = Join-Path $projectRoot 'scripts\run_pdf2excel.ps1'; Destination = Join-Path $package.PackageRoot 'scripts\run_pdf2excel.ps1' },
        @{ Source = Join-Path $projectRoot 'scripts\run_pdf2excel_menu.ps1'; Destination = Join-Path $package.PackageRoot 'scripts\run_pdf2excel_menu.ps1' },
        @{ Source = Join-Path $projectRoot 'scripts\new_profile_scaffold.ps1'; Destination = Join-Path $package.PackageRoot 'scripts\new_profile_scaffold.ps1' },
        @{ Source = Join-Path $projectRoot 'scripts\pdf2excel.common.ps1'; Destination = Join-Path $package.PackageRoot 'scripts\pdf2excel.common.ps1' },
        @{ Source = Join-Path $projectRoot ("template\{0}" -f $package.TemplateFile); Destination = Join-Path $package.PackageRoot ("template\{0}" -f $package.TemplateFile) },
        @{ Source = Join-Path $projectRoot 'template\vba\PDF2ExcelMacros.bas'; Destination = Join-Path $package.PackageRoot 'template\vba\PDF2ExcelMacros.bas' },
        @{ Source = Join-Path $projectRoot 'template\vba\PDF2ExcelMacros.sjis.bas'; Destination = Join-Path $package.PackageRoot 'template\vba\PDF2ExcelMacros.sjis.bas' },
        @{ Source = Join-Path $projectRoot 'template\vba\PDF2ExcelTemplateBuilder.bas'; Destination = Join-Path $package.PackageRoot 'template\vba\PDF2ExcelTemplateBuilder.bas' },
        @{ Source = Join-Path $projectRoot 'template\vba\PDF2ExcelTemplateBuilder.sjis.bas'; Destination = Join-Path $package.PackageRoot 'template\vba\PDF2ExcelTemplateBuilder.sjis.bas' },
        @{ Source = Join-Path $projectRoot $package.ReadmeSource; Destination = Join-Path $package.PackageRoot 'HANDOFF_README.md' }
    )

    if (-not [string]::IsNullOrWhiteSpace($package.RunScript)) {
        $filesToCopy += @{
            Source = Join-Path $projectRoot ("scripts\{0}" -f $package.RunScript)
            Destination = Join-Path $package.PackageRoot ("scripts\{0}" -f $package.RunScript)
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($package.MenuScript)) {
        $filesToCopy += @{
            Source = Join-Path $projectRoot ("scripts\{0}" -f $package.MenuScript)
            Destination = Join-Path $package.PackageRoot ("scripts\{0}" -f $package.MenuScript)
        }
    }

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
