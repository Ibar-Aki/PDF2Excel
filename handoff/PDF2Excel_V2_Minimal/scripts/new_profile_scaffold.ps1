param(
    [ValidateSet('v1', 'v2')]
    [string]$VersionMode,
    [string]$ProfileName,
    [string]$DisplayName,
    [string]$OutputPath,
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'pdf2excel.common.ps1')

[Console]::InputEncoding = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding = [Console]::OutputEncoding

$projectRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))

function Read-VersionMode {
    while ($true) {
        $selection = Read-Host '対象バージョンを選んでください (1=VER1, 2=VER2)'
        switch ($selection) {
            '1' { return 'v1' }
            '2' { return 'v2' }
            default { Write-Host '1 または 2 を入力してください。' }
        }
    }
}

function Read-ProfileName {
    param([string]$DefaultValue)

    while ($true) {
        $prompt = if ([string]::IsNullOrWhiteSpace($DefaultValue)) { '内部名/ファイル名に使うプロファイル名を入力してください (ASCII のみ)' } else { "内部名/ファイル名に使うプロファイル名を入力してください (Enter で $DefaultValue)" }
        $value = Read-Host $prompt
        if ([string]::IsNullOrWhiteSpace($value)) {
            $value = $DefaultValue
        }

        if ([string]::IsNullOrWhiteSpace($value)) {
            Write-Host 'プロファイル名は必須です。'
            continue
        }

        if ($value -notmatch '^[A-Za-z0-9_-]+$') {
            Write-Host 'プロファイル名は英数字、ハイフン、アンダースコアのみ使用できます。'
            continue
        }

        return $value
    }
}

function Read-OptionalValue {
    param(
        [Parameter(Mandatory = $true)][string]$Prompt,
        [string]$DefaultValue
    )

    $value = Read-Host $(if ([string]::IsNullOrWhiteSpace($DefaultValue)) { $Prompt } else { "$Prompt (Enter で $DefaultValue)" })
    if ([string]::IsNullOrWhiteSpace($value)) {
        return $DefaultValue
    }

    return $value
}

function New-ProfileScaffoldObject {
    param(
        [Parameter(Mandatory = $true)][string]$ResolvedVersionMode,
        [Parameter(Mandatory = $true)][string]$ResolvedProfileName,
        [Parameter(Mandatory = $true)][string]$ResolvedDisplayName
    )

    $profile = [ordered]@{
        name                      = $ResolvedProfileName
        displayName               = $ResolvedDisplayName
        description               = if ($ResolvedVersionMode -eq 'v2') { '生データ転記向けのサンプルプロファイルです。列定義と review 列を調整して使います。' } else { '標準変換向けのサンプルプロファイルです。列数やヘッダー行数を調整して使います。' }
        expectedColumns           = 10
        headerRowsToSkip          = 1
        targetRowCount            = 100
        allowMoreColumns          = $false
        preferredTableKinds       = @('Table')
        preferredTableNameContains = @()
        preferredTableIdContains  = @()
        sourceFileColumnName      = if ($ResolvedVersionMode -eq 'v2') { '元ファイル名' } else { 'SourceFile' }
        dataColumnPrefix          = if ($ResolvedVersionMode -eq 'v2') { '項目' } else { 'Column' }
    }

    if ($ResolvedVersionMode -eq 'v2') {
        $profile.multiPageMergeMode = 'single'
        $profile.normalizedTimeColumns = @()
        $profile.reviewPersonColumn = $null
        $profile.reviewSiteColumn = $null
        $profile.reviewInTimeColumn = $null
        $profile.reviewOutTimeColumn = $null
    }

    return $profile
}

if ([string]::IsNullOrWhiteSpace($VersionMode)) {
    $VersionMode = Read-VersionMode
}

$ProfileName = Read-ProfileName -DefaultValue $ProfileName
$defaultDisplayName = if ($VersionMode -eq 'v2') { '生データ転記サンプルプロファイル' } else { "$ProfileName プロファイル" }
if ([string]::IsNullOrWhiteSpace($DisplayName)) {
    $DisplayName = Read-OptionalValue -Prompt '表示名を入力してください' -DefaultValue $defaultDisplayName
}

$defaultOutputPath = Join-Path (Join-Path $projectRoot ("config\profiles\{0}" -f $VersionMode)) ("{0}.json" -f $ProfileName)
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Read-OptionalValue -Prompt '保存先を入力してください' -DefaultValue $defaultOutputPath
}
$OutputPath = [System.IO.Path]::GetFullPath($OutputPath)
$outputDir = Split-Path -Path $OutputPath -Parent
Ensure-Directory -Path $outputDir

if ((Test-Path -LiteralPath $OutputPath) -and -not $Force) {
    throw "プロファイルは既に存在します。上書きする場合は -Force を指定してください: $OutputPath"
}

$profileObject = New-ProfileScaffoldObject -ResolvedVersionMode $VersionMode -ResolvedProfileName $ProfileName -ResolvedDisplayName $DisplayName
$json = $profileObject | ConvertTo-Json -Depth 6
[System.IO.File]::WriteAllText($OutputPath, $json + [Environment]::NewLine, [System.Text.UTF8Encoding]::new($true))

Write-Host ''
Write-Host ("{0} のプロファイル雛形を作成しました。" -f (Get-VersionDisplayName -VersionMode $VersionMode))
Write-Host ("内部ID: {0}" -f $ProfileName)
Write-Host ("表示名: {0}" -f $DisplayName)
Write-Host ("保存先: {0}" -f $OutputPath)
