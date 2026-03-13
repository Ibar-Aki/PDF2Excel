param(
    [string]$OutputRoot = (Join-Path (Join-Path $PSScriptRoot '..') 'handoff\PDF2Excel_Minimal'),
    [string]$ZipPath = (Join-Path (Join-Path $PSScriptRoot '..') 'handoff\PDF2Excel_Minimal.zip')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))

function Ensure-Directory {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Reset-Directory {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Recurse -Force
    }
    New-Item -ItemType Directory -Path $Path -Force | Out-Null
}

function New-EmptyFile {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType File -Path $Path | Out-Null
    }
}

$handoffParent = Split-Path -Path $OutputRoot -Parent
Ensure-Directory -Path $handoffParent
Reset-Directory -Path $OutputRoot

$directories = @(
    'scripts',
    'template',
    'config\profiles',
    'input',
    'output',
    'output\runtime',
    'logs'
)

foreach ($relativeDir in $directories) {
    Ensure-Directory -Path (Join-Path $OutputRoot $relativeDir)
}

$filesToCopy = @(
    @{ Source = Join-Path $projectRoot 'run_pdf2excel.bat'; Destination = Join-Path $OutputRoot 'run_pdf2excel.bat' },
    @{ Source = Join-Path $projectRoot 'scripts\run_pdf2excel.ps1'; Destination = Join-Path $OutputRoot 'scripts\run_pdf2excel.ps1' },
    @{ Source = Join-Path $projectRoot 'template\PDF2Excel_Converter.xlsm'; Destination = Join-Path $OutputRoot 'template\PDF2Excel_Converter.xlsm' },
    @{ Source = Join-Path $projectRoot 'config\profiles\default.json'; Destination = Join-Path $OutputRoot 'config\profiles\default.json' },
    @{ Source = Join-Path $projectRoot 'handoff\HANDOFF_README_SOURCE.md'; Destination = Join-Path $OutputRoot 'HANDOFF_README.md' }
)

foreach ($file in $filesToCopy) {
    Copy-Item -LiteralPath $file.Source -Destination $file.Destination -Force
}

foreach ($gitkeepPath in @(
    (Join-Path $OutputRoot 'input\.gitkeep'),
    (Join-Path $OutputRoot 'output\.gitkeep'),
    (Join-Path $OutputRoot 'output\runtime\.gitkeep'),
    (Join-Path $OutputRoot 'logs\.gitkeep')
)) {
    New-EmptyFile -Path $gitkeepPath
}

if (Test-Path -LiteralPath $ZipPath) {
    Remove-Item -LiteralPath $ZipPath -Force
}

Compress-Archive -Path (Join-Path $OutputRoot '*') -DestinationPath $ZipPath

Write-Host "Handoff folder: $OutputRoot"
Write-Host "Handoff zip: $ZipPath"
