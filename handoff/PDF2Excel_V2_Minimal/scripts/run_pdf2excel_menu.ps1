param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$ForwardArgs,
    [ValidateSet('v1', 'v2')]
    [string]$VersionMode = 'v1',
    [string]$RunScriptPath,
    [switch]$PassCoreDefaults,
    [ValidateSet('Standard', 'Secure')]
    [string]$DefaultSecurityMode,
    [string]$DefaultProfileName,
    [switch]$ForceMenu
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

[Console]::InputEncoding = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding = [Console]::OutputEncoding

$scriptRoot = $PSScriptRoot
. (Join-Path $scriptRoot 'pdf2excel.common.ps1')
$projectRoot = [System.IO.Path]::GetFullPath((Join-Path $scriptRoot '..'))
$runScript = if ([string]::IsNullOrWhiteSpace($RunScriptPath)) { Join-Path $scriptRoot ("run_pdf2excel_{0}.ps1" -f $VersionMode) } else { $RunScriptPath }
$profileScaffoldScript = Join-Path $scriptRoot 'new_profile_scaffold.ps1'
$manualPath = Join-Path $projectRoot 'docs\user-manual.md'
$profileDir = Join-Path $projectRoot ("config\profiles\{0}" -f $VersionMode)
$outputDir = Join-Path $projectRoot 'output'
$isSecureDefault = ($VersionMode -eq 'v2' -and $DefaultSecurityMode -eq 'Secure')
$logsDir = if ($isSecureDefault) {
    $secureLogsDir = Get-LocalAppDataPdf2ExcelPath -ChildPath 'logs'
    if ([string]::IsNullOrWhiteSpace($secureLogsDir)) {
        Join-Path $projectRoot 'logs'
    } else {
        $secureLogsDir
    }
} else {
    Join-Path $projectRoot 'logs'
}
$script:executionLocation = Get-PathLocationInfo -Path $projectRoot
$systemLabel = if ($VersionMode -eq 'v2') { 'PDF2Excel VER2 - 生データ転記' } else { 'PDF2Excel VER1 - 標準変換' }
$completionSheetMessage = if ($VersionMode -eq 'v2') { 'Result / Review / Errors / Summary を確認してください。' } else { 'Result / Errors / Summary を確認してください。' }

function Assert-ExecutionLocationAllowed {
    if (-not $script:executionLocation.IsShared) {
        return
    }

    if ($isSecureDefault) {
        Write-Host ''
        Write-Host 'Secure モードでは共有パス上から実行できません。ローカルへ展開して再実行してください。'
        Write-Host ("実行元: {0}" -f $script:executionLocation.NormalizedPath)
        exit 2
    }

    Write-Host ''
    Write-Host '警告: 共有パス上から実行しています。ローカル実行を推奨します。'
    Write-Host ("実行元: {0}" -f $script:executionLocation.NormalizedPath)
    Pause
}

function Invoke-RunScript {
    param([string[]]$Arguments)

    $shellArgs = @('-NoProfile')
    if ($VersionMode -eq 'v2') {
        $shellArgs += @('-ExecutionPolicy', 'RemoteSigned')
    } else {
        $shellArgs += @('-ExecutionPolicy', 'Bypass')
    }
    $effectiveArguments = @()
    if ($PassCoreDefaults) {
        if (-not ($Arguments -contains '-VersionMode')) {
            $effectiveArguments += @('-VersionMode', $VersionMode)
        }
        if (-not [string]::IsNullOrWhiteSpace($DefaultSecurityMode) -and -not ($Arguments -contains '-SecurityMode')) {
            $effectiveArguments += @('-SecurityMode', $DefaultSecurityMode)
        }
        if (
            -not [string]::IsNullOrWhiteSpace($DefaultProfileName) -and
            -not ($Arguments -contains '-ProfileName') -and
            -not ($Arguments -contains '-ProfilePath')
        ) {
            $effectiveArguments += @('-ProfileName', $DefaultProfileName)
        }
    }

    $effectiveArguments += $Arguments
    $shellArgs += @('-File', $runScript)
    & powershell @shellArgs @effectiveArguments
    return $LASTEXITCODE
}

function Open-PathIfExists {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$MissingMessage
    )

    if (Test-Path -LiteralPath $Path) {
        Start-Process -FilePath $Path | Out-Null
        return
    }

    Write-Host ''
    Write-Host $MissingMessage
    Pause
}

function Invoke-ProfileScaffoldScript {
    $shellArgs = @('-NoProfile')
    if ($VersionMode -eq 'v2') {
        $shellArgs += @('-ExecutionPolicy', 'RemoteSigned')
    } else {
        $shellArgs += @('-ExecutionPolicy', 'Bypass')
    }

    $shellArgs += @('-File', $profileScaffoldScript, '-VersionMode', $VersionMode)
    & powershell @shellArgs
    return $LASTEXITCODE
}

$remainingArgs = @($ForwardArgs)
Assert-ExecutionLocationAllowed
if (-not $ForceMenu -and $remainingArgs.Count -gt 0) {
    exit (Invoke-RunScript -Arguments $remainingArgs)
}

while ($true) {
    Clear-Host
    Write-Host '=========================================='
    Write-Host (" $systemLabel")
    Write-Host '=========================================='
    Write-Host ''
    if ($VersionMode -eq 'v2') {
        Write-Host ' 生データ転記を行います。'
        Write-Host ' 曖昧な候補は Review へ分離し、確認しながら扱います。'
        Write-Host ' 保存先を指定しない場合は output フォルダに保存します。'
    } else {
        Write-Host ' PDF の表をまとめて Excel に変換します。'
        Write-Host ' 保存先を指定しない場合は output フォルダに保存します。'
    }
    Write-Host ''
    Write-Host ' [1] PDFファイルを選んで変換'
    Write-Host ' [2] PDFフォルダを選んで変換'
    Write-Host ' [3] 使い方マニュアルを開く'
    Write-Host ' [4] 出力フォルダを開く'
    Write-Host ' [5] プロファイルフォルダを開く'
    Write-Host ' [6] プロファイル雛形を作成'
    Write-Host ' [7] ログフォルダを開く'
    Write-Host ' [8] 終了'
    Write-Host ''

    $selection = Read-Host '番号を選んでください'
    switch ($selection) {
        '1' {
            $exitCode = Invoke-RunScript -Arguments @('-PromptForOutputFile')
            if ($exitCode -ne 0) {
                Write-Host ''
                Write-Host "PDF2Excel の処理に失敗しました。終了コード: $exitCode"
                Write-Host "詳細は '$logsDir' のログを確認してください。"
                Pause
            } else {
                Write-Host ''
                Write-Host '変換が完了しました。'
                Write-Host "出力先: $outputDir"
                Write-Host "ログ先: $logsDir"
                Write-Host $completionSheetMessage
                Write-Host '必要に応じて output と logs の内容を確認してください。'
                Pause
            }
        }
        '2' {
            $exitCode = Invoke-RunScript -Arguments @('-SelectInputFolder', '-PromptForOutputFile')
            if ($exitCode -ne 0) {
                Write-Host ''
                Write-Host "PDF2Excel の処理に失敗しました。終了コード: $exitCode"
                Write-Host "詳細は '$logsDir' のログを確認してください。"
                Pause
            } else {
                Write-Host ''
                Write-Host '変換が完了しました。'
                Write-Host "出力先: $outputDir"
                Write-Host "ログ先: $logsDir"
                Write-Host $completionSheetMessage
                Write-Host '必要に応じて output と logs の内容を確認してください。'
                Pause
            }
        }
        '3' {
            Open-PathIfExists -Path $manualPath -MissingMessage "使い方マニュアルが見つかりません: $manualPath"
        }
        '4' {
            Open-PathIfExists -Path $outputDir -MissingMessage "出力フォルダが見つかりません: $outputDir"
        }
        '5' {
            Open-PathIfExists -Path $profileDir -MissingMessage "プロファイルフォルダが見つかりません: $profileDir"
        }
        '6' {
            $exitCode = Invoke-ProfileScaffoldScript
            if ($exitCode -ne 0) {
                Write-Host ''
                Write-Host "プロファイル雛形の作成に失敗しました。終了コード: $exitCode"
                Pause
            } else {
                Write-Host ''
                Write-Host 'プロファイル雛形を作成しました。'
                Write-Host "保存先: $profileDir"
                Pause
            }
        }
        '7' {
            Open-PathIfExists -Path $logsDir -MissingMessage "ログフォルダが見つかりません: $logsDir"
        }
        '8' {
            exit 0
        }
        default {
            Write-Host ''
            Write-Host '1 から 8 の番号を入力してください。'
            Pause
        }
    }
}
