param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$ForwardArgs,
    [ValidateSet('v1', 'v2')]
    [string]$VersionMode = 'v2',
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
$secureRuntimeDir = Get-LocalAppDataPdf2ExcelPath -ChildPath 'runtime'
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
$systemLabel = if ($VersionMode -eq 'v2') { 'PDF2Excel VER2 - 生データ転記' } else { 'PDF2Excel VER1 - 開発用標準変換' }
$completionSheetMessage = if ($VersionMode -eq 'v2') { 'Result / Review / Errors / Summary を確認してください。' } else { 'Result / Errors / Summary を確認してください。' }

function Assert-SecureLocalStorageAvailable {
    if (-not $isSecureDefault) {
        return
    }

    if ([string]::IsNullOrWhiteSpace($secureRuntimeDir) -or [string]::IsNullOrWhiteSpace($secureLogsDir)) {
        Write-Host ''
        Write-Host 'VER2 Secure の正式運用では LOCALAPPDATA が必要です。'
        Write-Host 'この端末ではローカル runtime / logs を確保できないため、実行を停止します。'
        exit 2
    }
}

function Assert-ExecutionLocationAllowed {
    if (-not $script:executionLocation.IsShared) {
        return
    }

    if ($isSecureDefault) {
        Write-Host ''
        Write-Host 'VER2 Secure の正式運用では共有パス上から実行できません。ローカルへ展開して再実行してください。'
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

    $reportPath = Get-RunReportTempPath

    $shellArgs = @('-NoProfile', '-ExecutionPolicy', 'RemoteSigned')
    $effectiveArguments = @()
    if ($PassCoreDefaults) {
        if (-not ($Arguments -contains '-VersionMode')) {
            $effectiveArguments += @('-VersionMode', $VersionMode)
        }
        if (-not [string]::IsNullOrWhiteSpace($DefaultSecurityMode) -and -not ($Arguments -contains '-SecurityMode')) {
            $effectiveArguments += @('-SecurityMode', $DefaultSecurityMode)
        }
        if (
            -not [string]::IsNullOrWhiteSpace($script:currentProfileName) -and
            -not ($Arguments -contains '-ProfileName') -and
            -not ($Arguments -contains '-ProfilePath')
        ) {
            $effectiveArguments += @('-ProfileName', $script:currentProfileName)
        }
    }

    $effectiveArguments += @($Arguments + @('-RunReportPath', $reportPath))
    $shellArgs += @('-File', $runScript)
    & powershell @shellArgs @effectiveArguments | Out-Host
    $exitCode = $LASTEXITCODE
    $report = $null
    if (Test-Path -LiteralPath $reportPath) {
        try {
            $report = Get-Content -LiteralPath $reportPath -Raw -Encoding UTF8 | ConvertFrom-Json
        } catch {
        }
        Remove-Item -LiteralPath $reportPath -Force -ErrorAction SilentlyContinue
    }

    return [pscustomobject]@{
        ExitCode = $exitCode
        Report   = $report
    }
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
    $shellArgs = @('-NoProfile', '-ExecutionPolicy', 'RemoteSigned')
    $shellArgs += @('-File', $profileScaffoldScript, '-VersionMode', $VersionMode)
    if ($VersionMode -eq 'v2') {
        $shellArgs += '-Wizard'
    }
    & powershell @shellArgs
    return $LASTEXITCODE
}

function Get-MenuStorageRoot {
    if ($isSecureDefault) {
        return Join-Path $secureRuntimeDir 'menu'
    }

    return Join-Path $projectRoot 'logs\menu'
}

function Get-MenuStatePath {
    $stateRoot = Get-MenuStorageRoot
    Ensure-Directory -Path $stateRoot
    return Join-Path $stateRoot ("menu_state_{0}.json" -f $VersionMode)
}

function Get-RunReportTempPath {
    $stateRoot = Get-MenuStorageRoot
    Ensure-Directory -Path $stateRoot
    return Join-Path $stateRoot ("run_report_{0}.json" -f ([guid]::NewGuid().ToString('N')))
}

function Get-AvailableProfiles {
    if (-not (Test-Path -LiteralPath $profileDir)) {
        return @()
    }

    $profiles = @()
    foreach ($file in @(Get-ChildItem -LiteralPath $profileDir -Filter '*.json' -File | Sort-Object Name)) {
        try {
            $raw = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
            $profileName = if ([string]::IsNullOrWhiteSpace([string]$raw.name)) { [System.IO.Path]::GetFileNameWithoutExtension($file.Name) } else { [string]$raw.name }
            $displayName = if ([string]::IsNullOrWhiteSpace([string]$raw.displayName)) { $profileName } else { [string]$raw.displayName }
            $profiles += [pscustomobject]@{
                Name        = $profileName
                DisplayName = $displayName
                Path        = $file.FullName
                Description = [string]$raw.description
            }
        } catch {
        }
    }

    return @($profiles)
}

function Load-MenuState {
    $statePath = Get-MenuStatePath
    if (-not (Test-Path -LiteralPath $statePath)) {
        return $null
    }

    try {
        return Get-Content -LiteralPath $statePath -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        return $null
    }
}

function Save-MenuState {
    param([string]$SelectedProfileName)

    $statePath = Get-MenuStatePath
    $payload = [ordered]@{
        versionMode         = $VersionMode
        selectedProfileName = $SelectedProfileName
        updatedAt           = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    } | ConvertTo-Json
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($statePath, $payload, $utf8)
}

function Resolve-CurrentProfileName {
    param([object[]]$Profiles)

    $profileNames = @($Profiles | ForEach-Object { $_.Name })
    if (-not $profileNames) {
        return $null
    }

    $state = Load-MenuState
    if ($state -and $profileNames -contains [string]$state.selectedProfileName) {
        return [string]$state.selectedProfileName
    }

    if (-not [string]::IsNullOrWhiteSpace($DefaultProfileName) -and $profileNames -contains $DefaultProfileName) {
        return $DefaultProfileName
    }

    return $Profiles[0].Name
}

function Get-CurrentProfileDisplay {
    param([object[]]$Profiles)

    if ([string]::IsNullOrWhiteSpace($script:currentProfileName)) {
        return '未選択'
    }

    $profile = @($Profiles | Where-Object Name -eq $script:currentProfileName | Select-Object -First 1)
    if ($profile.Count -eq 0) {
        return $script:currentProfileName
    }

    return ("{0} ({1})" -f $profile[0].DisplayName, $profile[0].Name)
}

function Select-ProfileFromMenu {
    param([object[]]$Profiles)

    if ($Profiles.Count -eq 0) {
        Write-Host ''
        Write-Host '利用可能なプロファイルがありません。先に雛形を作成してください。'
        Pause
        return
    }

    Write-Host ''
    Write-Host '利用可能なプロファイル'
    for ($index = 0; $index -lt $Profiles.Count; $index += 1) {
        $profile = $Profiles[$index]
        $currentMark = if ($profile.Name -eq $script:currentProfileName) { ' (現在)' } else { '' }
        Write-Host (" [{0}] {1}{2}" -f ($index + 1), $profile.DisplayName, $currentMark)
        if (-not [string]::IsNullOrWhiteSpace($profile.Description)) {
            Write-Host ("      {0}" -f $profile.Description)
        }
    }
    Write-Host ''
    $selection = Read-Host '使うプロファイル番号を選んでください (Enter で戻る)'
    if ([string]::IsNullOrWhiteSpace($selection)) {
        return
    }

    $selectedIndex = 0
    if (-not [int]::TryParse($selection, [ref]$selectedIndex) -or $selectedIndex -lt 1 -or $selectedIndex -gt $Profiles.Count) {
        Write-Host ''
        Write-Host '有効な番号を入力してください。'
        Pause
        return
    }

    $script:currentProfileName = $Profiles[$selectedIndex - 1].Name
    Save-MenuState -SelectedProfileName $script:currentProfileName
    Write-Host ''
    Write-Host ("現在のプロファイルを '{0}' に変更しました。" -f $Profiles[$selectedIndex - 1].DisplayName)
    Pause
}

function Prompt-PostRunAction {
    param($Report)

    if ($null -eq $Report) {
        Pause
        return
    }

    Write-Host ''
    $selection = Read-Host 'O=出力ファイルを開く / F=保存先を開く / Enter=メニューへ戻る'
    switch ($selection.Trim().ToUpperInvariant()) {
        'O' {
            if (-not [string]::IsNullOrWhiteSpace($Report.OutputPath) -and (Test-Path -LiteralPath $Report.OutputPath)) {
                Start-Process -FilePath $Report.OutputPath | Out-Null
            }
        }
        'F' {
            if (-not [string]::IsNullOrWhiteSpace($Report.OutputParent) -and (Test-Path -LiteralPath $Report.OutputParent)) {
                Start-Process -FilePath $Report.OutputParent | Out-Null
            }
        }
    }
}

function Show-RunFailure {
    param($RunResult)

    Write-Host ''
    Write-Host ("PDF2Excel の処理に失敗しました。終了コード: {0}" -f $RunResult.ExitCode)
    if ($RunResult.Report) {
        if (-not [string]::IsNullOrWhiteSpace($RunResult.Report.ErrorCategory) -or -not [string]::IsNullOrWhiteSpace($RunResult.Report.ErrorCode)) {
            Write-Host ("分類: {0} / {1}" -f $RunResult.Report.ErrorCategory, $RunResult.Report.ErrorCode)
        }
        if (-not [string]::IsNullOrWhiteSpace($RunResult.Report.ActionHint)) {
            Write-Host ("対処: {0}" -f $RunResult.Report.ActionHint)
        }
        if (-not [string]::IsNullOrWhiteSpace($RunResult.Report.LogPath)) {
            Write-Host ("ログ: {0}" -f $RunResult.Report.LogPath)
        } else {
            Write-Host ("ログ: {0}" -f $logsDir)
        }
        if (-not [string]::IsNullOrWhiteSpace($RunResult.Report.EnvironmentReportPath)) {
            Write-Host ("環境チェックレポート: {0}" -f $RunResult.Report.EnvironmentReportPath)
        }
    } else {
        Write-Host ("詳細は '{0}' のログを確認してください。" -f $logsDir)
    }
    Pause
}

function Show-RunSuccess {
    param($RunResult)

    Write-Host ''
    Write-Host '変換が完了しました。'
    if ($RunResult.Report) {
        Write-Host ("出力ファイル: {0}" -f $RunResult.Report.OutputPath)
        if (-not [string]::IsNullOrWhiteSpace($RunResult.Report.LogPath)) {
            Write-Host ("ログ先: {0}" -f $RunResult.Report.LogPath)
        }
        if (-not [string]::IsNullOrWhiteSpace($RunResult.Report.RunHistoryPath)) {
            Write-Host ("実行履歴: {0}" -f $RunResult.Report.RunHistoryPath)
        }
        if (-not [string]::IsNullOrWhiteSpace($RunResult.Report.ProfileDisplayName)) {
            Write-Host ("使用プロファイル: {0}" -f $RunResult.Report.ProfileDisplayName)
        }
    } else {
        Write-Host ("出力先: {0}" -f $outputDir)
        Write-Host ("ログ先: {0}" -f $logsDir)
    }
    Write-Host $completionSheetMessage
    Prompt-PostRunAction -Report $RunResult.Report
}

function Invoke-MenuRun {
    param([string[]]$Arguments)

    $runResult = Invoke-RunScript -Arguments $Arguments
    if ($runResult.ExitCode -ne 0) {
        Show-RunFailure -RunResult $runResult
        return
    }

    Show-RunSuccess -RunResult $runResult
}

function Invoke-EnvironmentCheckFromMenu {
    $runResult = Invoke-RunScript -Arguments @('-CheckEnvironment')
    if ($runResult.ExitCode -ne 0) {
        Show-RunFailure -RunResult $runResult
        return
    }

    Write-Host ''
    Write-Host '環境チェックが完了しました。'
    if ($runResult.Report -and -not [string]::IsNullOrWhiteSpace($runResult.Report.EnvironmentReportPath)) {
        Write-Host ("レポート: {0}" -f $runResult.Report.EnvironmentReportPath)
    }
    Pause
}

$remainingArgs = @($ForwardArgs)
Assert-SecureLocalStorageAvailable
Assert-ExecutionLocationAllowed
$availableProfiles = @(Get-AvailableProfiles)
$script:currentProfileName = Resolve-CurrentProfileName -Profiles $availableProfiles
if (-not [string]::IsNullOrWhiteSpace($script:currentProfileName)) {
    Save-MenuState -SelectedProfileName $script:currentProfileName
}
if (-not $ForceMenu -and $remainingArgs.Count -gt 0) {
    exit ((Invoke-RunScript -Arguments $remainingArgs).ExitCode)
}

while ($true) {
    $availableProfiles = @(Get-AvailableProfiles)
    if (-not [string]::IsNullOrWhiteSpace($script:currentProfileName) -and -not (@($availableProfiles | ForEach-Object { $_.Name }) -contains $script:currentProfileName)) {
        $script:currentProfileName = Resolve-CurrentProfileName -Profiles $availableProfiles
        if (-not [string]::IsNullOrWhiteSpace($script:currentProfileName)) {
            Save-MenuState -SelectedProfileName $script:currentProfileName
        }
    }

    Clear-Host
    Write-Host '=========================================='
    Write-Host (" $systemLabel")
    Write-Host '=========================================='
    Write-Host ''
    if ($VersionMode -eq 'v2') {
        Write-Host ' 生データ転記を行います。'
        Write-Host ' 曖昧な候補は Review へ分離し、確認しながら扱います。'
        Write-Host ' 正式運用はローカル展開した VER2 Secure のみです。'
        Write-Host ' 保存先を指定しない場合は output フォルダに保存します。'
    } else {
        Write-Host ' 開発・検証用の標準変換です。'
        Write-Host ' 保存先を指定しない場合は output フォルダに保存します。'
    }
    Write-Host (" 現在のプロファイル: {0}" -f (Get-CurrentProfileDisplay -Profiles $availableProfiles))
    Write-Host ''
    Write-Host ' [1] PDFファイルを選んで変換'
    Write-Host ' [2] PDFフォルダを選んで変換'
    Write-Host ' [3] 使い方マニュアルを開く'
    Write-Host ' [4] 出力フォルダを開く'
    Write-Host ' [5] プロファイルフォルダを開く'
    Write-Host ' [6] プロファイル雛形を作成'
    Write-Host ' [7] ログフォルダを開く'
    Write-Host ' [8] 終了'
    Write-Host ' [9] プロファイルを選ぶ'
    Write-Host ' [0] 環境チェック'
    Write-Host ''

    $selection = Read-Host '番号を選んでください'
    switch ($selection) {
        '1' {
            Invoke-MenuRun -Arguments @('-PromptForOutputFile')
        }
        '2' {
            Invoke-MenuRun -Arguments @('-SelectInputFolder', '-PromptForOutputFile')
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
                $availableProfiles = @(Get-AvailableProfiles)
                Write-Host ''
                Write-Host 'プロファイル雛形を作成しました。'
                Write-Host "保存先: $profileDir"
                Write-Host '使うプロファイルを切り替える場合は [9] を選んでください。'
                Pause
            }
        }
        '7' {
            Open-PathIfExists -Path $logsDir -MissingMessage "ログフォルダが見つかりません: $logsDir"
        }
        '8' {
            exit 0
        }
        '9' {
            Select-ProfileFromMenu -Profiles $availableProfiles
        }
        '0' {
            Invoke-EnvironmentCheckFromMenu
        }
        default {
            Write-Host ''
            Write-Host '0 から 9 の番号を入力してください。'
            Pause
        }
    }
}
