param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$ForwardArgs
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

[Console]::InputEncoding = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding = [Console]::OutputEncoding

$scriptRoot = $PSScriptRoot
$projectRoot = [System.IO.Path]::GetFullPath((Join-Path $scriptRoot '..'))
$runScript = Join-Path $scriptRoot 'run_pdf2excel.ps1'
$manualPath = Join-Path $projectRoot 'docs\user-manual.md'
$profileDir = Join-Path $projectRoot 'config\profiles'
$outputDir = Join-Path $projectRoot 'output'
$logsDir = Join-Path $projectRoot 'logs'

function Invoke-RunScript {
    param([string[]]$Arguments)

    & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript @Arguments
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

$remainingArgs = @($ForwardArgs)
if ($remainingArgs.Count -gt 0) {
    exit (Invoke-RunScript -Arguments $remainingArgs)
}

while ($true) {
    Clear-Host
    Write-Host '=========================================='
    Write-Host ' PDF2Excel - PDF表 一括変換ツール'
    Write-Host '=========================================='
    Write-Host ''
    Write-Host ' PDF の表をまとめて Excel に変換します。'
    Write-Host ' 保存先を指定しない場合は output フォルダに保存します。'
    Write-Host ''
    Write-Host ' [1] PDFファイルを選んで変換'
    Write-Host ' [2] PDFフォルダを選んで変換'
    Write-Host ' [3] 使い方マニュアルを開く'
    Write-Host ' [4] 出力フォルダを開く'
    Write-Host ' [5] プロファイルフォルダを開く'
    Write-Host ' [6] ログフォルダを開く'
    Write-Host ' [7] 終了'
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
            Open-PathIfExists -Path $logsDir -MissingMessage "ログフォルダが見つかりません: $logsDir"
        }
        '7' {
            exit 0
        }
        default {
            Write-Host ''
            Write-Host '1 から 7 の番号を入力してください。'
            Pause
        }
    }
}
