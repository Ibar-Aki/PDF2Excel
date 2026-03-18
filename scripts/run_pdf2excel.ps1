param(
    [string]$InputFolder,
    [string[]]$InputFiles,
    [string]$OutputFile,
    [string]$ProfileName = 'default',
    [string]$ProfilePath,
    [ValidateSet('v1', 'v2')]
    [string]$VersionMode = 'v1',
    [ValidateSet('Standard', 'Secure')]
    [string]$SecurityMode = 'Standard',
    [ValidateSet('INFO', 'DEBUG')]
    [string]$LogLevel = 'INFO',
    [switch]$KeepInput,
    [switch]$RebuildTemplate,
    [switch]$OpenOutput,
    [switch]$OpenOutputFolder,
    [switch]$SelectInputFolder,
    [switch]$PromptForOutputFile,
    [switch]$NoConfirm,
    [switch]$SkipMain,
    [Alias('h')][switch]$Help
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'pdf2excel.common.ps1')

[Console]::InputEncoding = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding = [Console]::OutputEncoding

$script:runStartedAt = Get-Date
$baseDir = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$scriptDir = Join-Path $baseDir 'scripts'
$templateDir = Join-Path $baseDir 'template'
$configDir = Join-Path $baseDir 'config'
$profilesDir = Join-Path $configDir "profiles\$VersionMode"
$inputDir = Join-Path $baseDir 'input'
$outputDir = Join-Path $baseDir 'output'
$script:isSecureMode = ($VersionMode -eq 'v2' -and $SecurityMode -eq 'Secure')
$secureRuntimeRootDir = Get-LocalAppDataPdf2ExcelPath -ChildPath 'runtime'
$secureLogsDir = Get-LocalAppDataPdf2ExcelPath -ChildPath 'logs'
$script:runtimeUsesProjectFallback = ($script:isSecureMode -and [string]::IsNullOrWhiteSpace($secureRuntimeRootDir))
$script:logsUseProjectFallback = ($script:isSecureMode -and [string]::IsNullOrWhiteSpace($secureLogsDir))
$runtimeRootDir = if ($script:isSecureMode -and -not $script:runtimeUsesProjectFallback) {
    $secureRuntimeRootDir
} else {
    Join-Path $outputDir 'runtime'
}
$runtimeRunsDir = Join-Path $runtimeRootDir 'runs'
$logsDir = if ($script:isSecureMode -and -not $script:logsUseProjectFallback) {
    $secureLogsDir
} else {
    Join-Path $baseDir 'logs'
}
$script:executionLocation = Get-PathLocationInfo -Path $baseDir
$script:versionDisplayName = Get-VersionDisplayName -VersionMode $VersionMode
$templatePath = Join-Path $templateDir $(if ($VersionMode -eq 'v2') { 'PDF2Excel_V2_Converter.xlsm' } else { 'PDF2Excel_V1_Converter.xlsm' })
$buildTemplateScript = Join-Path $scriptDir 'build_excel_template.ps1'

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
$script:runInstanceId = "run_${timestamp}_$PID"
$script:runWorkspaceDir = Join-Path $runtimeRunsDir $script:runInstanceId
$script:runStagingDir = Join-Path $script:runWorkspaceDir 'staging'
$script:runRuntimeDir = Join-Path $script:runWorkspaceDir 'runtime'
$script:lockFilePath = Join-Path $runtimeRootDir 'run.lock'
$script:runMutex = $null
$script:logPath = Join-Path $logsDir "run_$timestamp.log"
$script:sensitivePathMap = [ordered]@{}
$script:cleanupFailureMessage = $null

# ============================================================
# Section: User-Facing Console Output
# ============================================================

function Show-Usage {
    @(
        "PDF2Excel $script:versionDisplayName 使い方",
        '',
        '1. かんたん操作:',
        '   正式運用では run_pdf2excel.bat をダブルクリックします。',
        '',
        '2. PowerShell から直接実行:',
        '',
        '   powershell -NoProfile -File .\scripts\run_pdf2excel.ps1',
        '   powershell -NoProfile -File .\scripts\run_pdf2excel.ps1 -InputFolder C:\Path\To\PdfFolder',
        '   powershell -NoProfile -File .\scripts\run_pdf2excel.ps1 -InputFolder C:\Path\To\PdfFolder -OutputFile C:\Path\To\result.xlsx',
        '   powershell -NoProfile -File .\scripts\run_pdf2excel.ps1 -SelectInputFolder -PromptForOutputFile',
        '',
        'オプション',
        '  -InputFolder         PDF が入っているフォルダを指定します。',
        '  -InputFiles          変換対象の PDF ファイルを個別指定します。',
        '  -OutputFile          出力する xlsx の保存先を指定します。',
        '  -ProfileName         使用する帳票プロファイル名を指定します。ラッパー経由では版ごとの既定値が使われます。',
        '  -ProfilePath         使用する帳票プロファイル JSON のフルパスを指定します。',
        '  -LogLevel           ログの詳細度を INFO または DEBUG で指定します。既定値は INFO です。',
        '  -KeepInput           input 内の過去PDFを保持します。実際の変換は今回分だけ別 staging で実行します。VER2 Secure では無効です。',
        '  -RebuildTemplate     xlsm テンプレートを再生成します。',
        '  -OpenOutput          完成した xlsx を自動で開きます。',
        '  -OpenOutputFolder    完成後に保存先フォルダを開きます。',
        '  -SelectInputFolder   フォルダ選択ダイアログを開いて入力フォルダを選びます。',
        '  -PromptForOutputFile 保存先の xlsx をダイアログで選びます。',
        '  -NoConfirm           実行前チェック画面を表示後、確認入力を求めずそのまま続行します。',
        '  -Help                このヘルプを表示します。'
    ) | Write-Host
}

if ($Help) {
    Show-Usage
    exit 0
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('DEBUG', 'INFO', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )

    if ($Level -eq 'DEBUG' -and $LogLevel -ne 'DEBUG') {
        return
    }

    $effectiveMessage = if ($script:isSecureMode -and $LogLevel -ne 'DEBUG') {
        Protect-MessagePaths -Message $Message -PathMap $script:sensitivePathMap
    } else {
        $Message
    }

    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $effectiveMessage
    Add-Content -LiteralPath $script:logPath -Value $line -Encoding UTF8
    Write-Host $line
}

function Register-SensitivePaths {
    param([string[]]$Paths)

    foreach ($path in @($Paths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        $normalizedPath = try {
            [System.IO.Path]::GetFullPath($path)
        } catch {
            $path
        }

        if (-not $script:sensitivePathMap.Contains($normalizedPath)) {
            $script:sensitivePathMap[$normalizedPath] = ConvertTo-MaskedPathText -Path $normalizedPath
        }
    }
}

function Get-ExcelProcessId {
    param($ExcelApplication)

    if ($null -eq $ExcelApplication) {
        return $null
    }

    try {
        $windowHandle = [int]$ExcelApplication.Hwnd
    } catch {
        return $null
    }

    return Get-WindowProcessId -WindowHandle $windowHandle
}

function Stop-ExcelProcessForCleanup {
    param(
        [Nullable[int]]$ProcessId,
        [string]$Reason = 'Secure クリーンアップのため'
    )

    if ($null -eq $ProcessId -or $ProcessId -le 0) {
        return $false
    }

    $excelProcess = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
    if ($null -eq $excelProcess -or $excelProcess.ProcessName -ne 'EXCEL') {
        return $false
    }

    try {
        Stop-Process -Id $ProcessId -Force -ErrorAction Stop
        Write-Log ("Excel プロセスを強制終了しました (PID={0}): {1}" -f $ProcessId, $Reason) 'WARN'
        Start-Sleep -Milliseconds 500
        return $true
    } catch {
        Write-Log ("Excel プロセスの強制終了に失敗しました (PID={0}): {1}" -f $ProcessId, $_.Exception.Message) 'WARN'
        return $false
    }
}

function Write-Banner {
    Write-Host ''
    Write-Host '==========================================' -ForegroundColor Cyan
    Write-Host (" PDF2Excel {0} - PDF表 一括変換ツール" -f $script:versionDisplayName) -ForegroundColor Cyan
    Write-Host '==========================================' -ForegroundColor Cyan
    Write-Host ''
}

function Show-RunSummary {
    param(
        [Parameter(Mandatory = $true)][string[]]$SourceFiles,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)]$Profile,
        [int]$ResultRows = 0,
        [int]$ErrorRows = 0,
        [int]$ReviewRows = 0,
        [int]$SuccessPdfCount = 0,
        [int]$FailedPdfCount = 0,
        [int]$ElapsedSeconds = 0
    )

    Write-Host ''
    Write-Host '実行結果' -ForegroundColor Green
    Write-Host ('  対象PDF数       : {0}' -f $SourceFiles.Count)
    Write-Host ('  成功PDF数       : {0}' -f $SuccessPdfCount)
    Write-Host ('  失敗PDF数       : {0}' -f $FailedPdfCount)
    Write-Host ('  取込データ行数   : {0}' -f $ResultRows)
    Write-Host ('  エラー件数      : {0}' -f $ErrorRows)
    if ($VersionMode -eq 'v2') {
        Write-Host ('  Review件数      : {0}' -f $ReviewRows)
    }
    Write-Host ('  処理時間(秒)    : {0}' -f $ElapsedSeconds)
    Write-Host ('  使用プロファイル : {0}' -f $Profile.DisplayName)
    Write-Host ('  出力ファイル     : {0}' -f $OutputPath)
    Write-Host ('  ログファイル     : {0}' -f $script:logPath)
    Write-Host ''
    Write-Host $(if ($VersionMode -eq 'v2') { '変換が完了しました。Result / Review / Errors / Summary を確認してください。' } else { '変換が完了しました。Result / Errors / Summary を確認してください。' }) -ForegroundColor Green
    Write-Host ''
}

function Show-ProcessingNotice {
    $message = 'PDF取り込みに時間がかかります。しばらくお待ちください....'
    Write-Host ''
    Write-Host $message -ForegroundColor Yellow
    Write-Host ''
    Write-Log $message
}

function Ensure-UiAssembliesLoaded {
    if ('System.Windows.Forms.Form' -as [type]) {
        return
    }

    Add-Type -AssemblyName System.Windows.Forms
}

function Remove-PathWithRetry {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [int]$MaxAttempts = 3
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt += 1) {
        if (-not (Test-Path -LiteralPath $Path)) {
            return $true
        }

        try {
            $item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
            if ($item.PSIsContainer) {
                Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
            } else {
                Remove-Item -LiteralPath $Path -Force -ErrorAction Stop
            }
            return $true
        } catch {
            if ($attempt -eq $MaxAttempts) {
                return $false
            }
            Start-Sleep -Milliseconds 500
        }
    }

    return $false
}

# ============================================================
# Section: Workspace and Run Lifecycle
# ============================================================

function Ensure-Workspace {
    foreach ($path in @(
        $inputDir,
        $outputDir,
        $runtimeRootDir,
        $runtimeRunsDir,
        $logsDir,
        $templateDir,
        (Join-Path $templateDir 'vba'),
        $configDir,
        $profilesDir
    )) {
        Ensure-Directory -Path $path
    }
}

function Write-EnvironmentWarnings {
    if ($script:isSecureMode -and $script:runtimeUsesProjectFallback) {
        throw 'VER2 Secure の正式運用では LOCALAPPDATA が必要です。runtime をローカルへ作成できないため停止します。'
    }

    if ($script:isSecureMode -and $script:logsUseProjectFallback) {
        throw 'VER2 Secure の正式運用では LOCALAPPDATA が必要です。ログ保存先をローカルへ作成できないため停止します。'
    }
}

function Assert-ExecutionLocationAllowed {
    if (-not $script:executionLocation.IsShared) {
        return
    }

    if ($script:isSecureMode) {
        throw 'VER2 Secure の正式運用では共有パス上から実行できません。ローカルへ展開して再実行してください。'
    }

    Write-Log ("共有パス上から実行しています。ローカル実行を推奨します: {0}" -f $script:executionLocation.NormalizedPath) 'WARN'
}

function Rotate-LogFiles {
    param(
        [int]$RetentionDays = 30,
        [int]$MaxFiles = 200
    )

    if (-not (Test-Path -LiteralPath $logsDir)) {
        return [pscustomobject]@{ RemovedByAge = 0; RemovedByCount = 0 }
    }

    $removedByAge = 0
    $removedByCount = 0
    $cutoff = (Get-Date).AddDays(-1 * $RetentionDays)
    $logFiles = @(Get-ChildItem -LiteralPath $logsDir -Filter 'run_*.log' -File -ErrorAction SilentlyContinue)

    foreach ($oldFile in @($logFiles | Where-Object LastWriteTime -lt $cutoff)) {
        Remove-Item -LiteralPath $oldFile.FullName -Force -ErrorAction SilentlyContinue
        $removedByAge += 1
    }

    $remaining = @(Get-ChildItem -LiteralPath $logsDir -Filter 'run_*.log' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)
    if ($remaining.Count -gt $MaxFiles) {
        foreach ($extraFile in @($remaining | Select-Object -Skip $MaxFiles)) {
            Remove-Item -LiteralPath $extraFile.FullName -Force -ErrorAction SilentlyContinue
            $removedByCount += 1
        }
    }

    return [pscustomobject]@{
        RemovedByAge   = $removedByAge
        RemovedByCount = $removedByCount
    }
}

function Compact-RuntimeArtifacts {
    if (-not (Test-Path -LiteralPath $runtimeRootDir)) {
        return
    }

    Get-ChildItem -LiteralPath $runtimeRunsDir -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        $deleted = Remove-PathWithRetry -Path $_.FullName
        if (-not $deleted) {
            Write-Log "実行時の一時ファイルを削除できませんでした: $($_.FullName)" 'WARN'
        }
    }

    Get-ChildItem -LiteralPath $runtimeRootDir -File -ErrorAction SilentlyContinue | Where-Object {
        $_.Name -notin @('.gitkeep', 'run.lock')
    } | ForEach-Object {
        $deleted = Remove-PathWithRetry -Path $_.FullName
        if (-not $deleted) {
            Write-Log "実行時の一時ファイルを削除できませんでした: $($_.FullName)" 'WARN'
        }
    }
}

function Acquire-RunLock {
    Ensure-Directory -Path $runtimeRootDir

    $mutexName = 'Global\PDF2Excel_RunMutex'
    $createdNew = $false
    $script:runMutex = New-Object System.Threading.Mutex($false, $mutexName, [ref]$createdNew)

    if (-not $script:runMutex.WaitOne(0, $false)) {
        $lockSummary = ''
        if (Test-Path -LiteralPath $script:lockFilePath) {
            try {
                $lockInfo = Get-Content -LiteralPath $script:lockFilePath -Raw -Encoding UTF8 | ConvertFrom-Json
                $lockSummary = " 実行中情報: 開始=$($lockInfo.startedAt), PID=$($lockInfo.pid)"
            } catch {
                $lockSummary = ' 実行中情報: run.lock は存在しますが内容を読めませんでした。'
            }
        }

        throw "別の PDF2Excel 実行が進行中です。完了後に再実行してください。$lockSummary"
    }

    $lockPayload = [ordered]@{
        runInstanceId = $script:runInstanceId
        pid           = $PID
        startedAt     = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    } | ConvertTo-Json

    Set-Content -LiteralPath $script:lockFilePath -Value $lockPayload -Encoding UTF8
}

function Release-RunLock {
    if ($script:runMutex) {
        try {
            $script:runMutex.ReleaseMutex() | Out-Null
        } catch {
        }

        try {
            $script:runMutex.Dispose()
        } catch {
        }

        $script:runMutex = $null
    }

    if (Test-Path -LiteralPath $script:lockFilePath) {
        Remove-PathWithRetry -Path $script:lockFilePath | Out-Null
    }
}

function Select-InputFolderDialog {
    Ensure-UiAssembliesLoaded
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = 'PDF が入っているフォルダを選択してください。'
    $dialog.ShowNewFolderButton = $false

    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        throw '入力フォルダが選択されませんでした。'
    }

    return $dialog.SelectedPath
}

# ============================================================
# Section: Input Resolution and Profile Loading
# ============================================================

function Select-PdfFiles {
    Ensure-UiAssembliesLoaded
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = '変換したい PDF を選択してください'
    $dialog.Filter = 'PDF files (*.pdf)|*.pdf'
    $dialog.Multiselect = $true

    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        throw 'PDF ファイルが選択されませんでした。'
    }

    return @($dialog.FileNames)
}

function Select-OutputFileDialog {
    param([Parameter(Mandatory = $true)][string]$DefaultOutputPath)

    Ensure-UiAssembliesLoaded
    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Title = '出力する Excel ファイルの保存先を選択してください'
    $dialog.Filter = 'Excel workbook (*.xlsx)|*.xlsx'
    $dialog.FileName = [System.IO.Path]::GetFileName($DefaultOutputPath)
    $dialog.InitialDirectory = Split-Path -Path $DefaultOutputPath -Parent
    $dialog.OverwritePrompt = $true

    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        throw '出力ファイルの保存先が選択されませんでした。'
    }

    return $dialog.FileName
}

function Resolve-InputPdfFiles {
    param(
        [string]$SourceFolder,
        [string[]]$SourceFiles
    )

    if ($SourceFiles) {
        $normalizedSourceFiles = @()
        foreach ($sourceFileEntry in $SourceFiles) {
            foreach ($candidate in ($sourceFileEntry -split '[,;\r\n]+')) {
                $trimmedCandidate = $candidate.Trim()
                if (-not [string]::IsNullOrWhiteSpace($trimmedCandidate)) {
                    $normalizedSourceFiles += $trimmedCandidate
                }
            }
        }

        $resolvedFiles = @()
        foreach ($sourceFile in $normalizedSourceFiles) {
            if (-not (Test-Path -LiteralPath $sourceFile)) {
                throw "入力ファイルが存在しません: $sourceFile"
            }

            $resolvedPath = (Resolve-Path -LiteralPath $sourceFile).Path
            if ([System.IO.Path]::GetExtension($resolvedPath).ToLowerInvariant() -ne '.pdf') {
                throw "PDF 以外のファイルは指定できません: $resolvedPath"
            }

            $resolvedFiles += $resolvedPath
        }

        if ($resolvedFiles.Count -eq 0) {
            throw '入力ファイルが指定されていません。'
        }

        return @($resolvedFiles)
    }

    if ([string]::IsNullOrWhiteSpace($SourceFolder)) {
        if ($SelectInputFolder) {
            $SourceFolder = Select-InputFolderDialog
        }
    }

    if ([string]::IsNullOrWhiteSpace($SourceFolder)) {
        return Select-PdfFiles
    }

    if (-not (Test-Path -LiteralPath $SourceFolder)) {
        throw "入力フォルダが存在しません: $SourceFolder"
    }

    $folderPath = (Resolve-Path -LiteralPath $SourceFolder).Path
    $files = Get-ChildItem -LiteralPath $folderPath -Filter '*.pdf' -File | Sort-Object Name | Select-Object -ExpandProperty FullName
    if (-not $files) {
        throw "指定したフォルダに PDF が見つかりません: $folderPath"
    }

    return @($files)
}

function Normalize-StringArray {
    param([object]$Value)

    if ($null -eq $Value) {
        return @()
    }

    if ($Value -is [System.Array]) {
        return @($Value | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }

    return @([string]$Value)
}

function Get-OptionalIntValue {
    param([object]$Value)

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return $null
    }

    return [int]$Value
}

function Get-OptionalProfileValue {
    param(
        [Parameter(Mandatory = $true)]$ProfileObject,
        [Parameter(Mandatory = $true)][string]$PropertyName
    )

    $property = $ProfileObject.PSObject.Properties[$PropertyName]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

function Get-ProfileConfiguration {
    param(
        [string]$RequestedProfileName,
        [string]$RequestedProfilePath
    )

    $resolvedProfilePath = $null
    if (-not [string]::IsNullOrWhiteSpace($RequestedProfilePath)) {
        if (-not (Test-Path -LiteralPath $RequestedProfilePath)) {
            throw "指定したプロファイルが存在しません: $RequestedProfilePath"
        }
        $resolvedProfilePath = (Resolve-Path -LiteralPath $RequestedProfilePath).Path
    } else {
        $profileFileName = if ([string]::IsNullOrWhiteSpace($RequestedProfileName)) { 'default.json' } else { "$RequestedProfileName.json" }
        $resolvedProfilePath = Join-Path $profilesDir $profileFileName
        if (-not (Test-Path -LiteralPath $resolvedProfilePath)) {
            $legacyProfilePath = Join-Path (Join-Path $configDir 'profiles') $profileFileName
            if (Test-Path -LiteralPath $legacyProfilePath) {
                $resolvedProfilePath = $legacyProfilePath
            } else {
                throw "プロファイルが見つかりません: $resolvedProfilePath"
            }
        }
        $resolvedProfilePath = (Resolve-Path -LiteralPath $resolvedProfilePath).Path
    }

    $rawProfile = Get-Content -LiteralPath $resolvedProfilePath -Raw -Encoding UTF8 | ConvertFrom-Json

    $expectedColumns = [int]$rawProfile.expectedColumns
    $headerRowsToSkip = [int]$rawProfile.headerRowsToSkip
    $targetRowCount = [int]$rawProfile.targetRowCount
    if ($expectedColumns -lt 1) {
        throw "expectedColumns は 1 以上で指定してください: $resolvedProfilePath"
    }
    if ($headerRowsToSkip -lt 0) {
        throw "headerRowsToSkip は 0 以上で指定してください: $resolvedProfilePath"
    }
    if ($targetRowCount -lt 1) {
        throw "targetRowCount は 1 以上で指定してください: $resolvedProfilePath"
    }

    $sourceFileColumnName = [string]$rawProfile.sourceFileColumnName
    if ([string]::IsNullOrWhiteSpace($sourceFileColumnName)) {
        $sourceFileColumnName = 'SourceFile'
    }

    $dataColumnPrefix = [string]$rawProfile.dataColumnPrefix
    if ([string]::IsNullOrWhiteSpace($dataColumnPrefix)) {
        $dataColumnPrefix = 'Column'
    }

    $reviewMappings = [pscustomobject]@{
        PersonColumn = Get-OptionalIntValue -Value (Get-OptionalProfileValue -ProfileObject $rawProfile -PropertyName 'reviewPersonColumn')
        SiteColumn   = Get-OptionalIntValue -Value (Get-OptionalProfileValue -ProfileObject $rawProfile -PropertyName 'reviewSiteColumn')
        InTimeColumn = Get-OptionalIntValue -Value (Get-OptionalProfileValue -ProfileObject $rawProfile -PropertyName 'reviewInTimeColumn')
        OutTimeColumn = Get-OptionalIntValue -Value (Get-OptionalProfileValue -ProfileObject $rawProfile -PropertyName 'reviewOutTimeColumn')
    }

    $normalizedTimeColumns = @()
    $rawNormalizedTimeColumns = Get-OptionalProfileValue -ProfileObject $rawProfile -PropertyName 'normalizedTimeColumns'
    if ($null -ne $rawNormalizedTimeColumns) {
        foreach ($entry in @($rawNormalizedTimeColumns)) {
            if ($null -eq $entry) {
                continue
            }

            $sourceColumn = Get-OptionalIntValue -Value (Get-OptionalProfileValue -ProfileObject $entry -PropertyName 'sourceColumn')
            $displayName = [string](Get-OptionalProfileValue -ProfileObject $entry -PropertyName 'displayName')
            if ($null -eq $sourceColumn -or [string]::IsNullOrWhiteSpace($displayName)) {
                continue
            }

            $minutesColumnName = [string](Get-OptionalProfileValue -ProfileObject $entry -PropertyName 'minutesColumnName')
            if ([string]::IsNullOrWhiteSpace($minutesColumnName)) {
                $minutesColumnName = '{0}_分' -f $displayName
            }

            $normalizedTimeColumns += [pscustomobject]@{
                SourceColumn      = $sourceColumn
                DisplayName       = $displayName
                MinutesColumnName = $minutesColumnName
            }
        }
    }

    return [pscustomobject]@{
        Name                       = if ([string]::IsNullOrWhiteSpace($rawProfile.name)) { [System.IO.Path]::GetFileNameWithoutExtension($resolvedProfilePath) } else { [string]$rawProfile.name }
        DisplayName                = if ([string]::IsNullOrWhiteSpace($rawProfile.displayName)) { [System.IO.Path]::GetFileNameWithoutExtension($resolvedProfilePath) } else { [string]$rawProfile.displayName }
        Description                = if ([string]::IsNullOrWhiteSpace($rawProfile.description)) { '説明は未設定です。' } else { [string]$rawProfile.description }
        ExpectedColumns            = $expectedColumns
        HeaderRowsToSkip           = $headerRowsToSkip
        TargetRowCount             = $targetRowCount
        AllowMoreColumns           = [bool]$rawProfile.allowMoreColumns
        PreferredTableKinds        = @(Normalize-StringArray -Value $rawProfile.preferredTableKinds)
        PreferredTableNameContains = @(Normalize-StringArray -Value $rawProfile.preferredTableNameContains)
        PreferredTableIdContains   = @(Normalize-StringArray -Value $rawProfile.preferredTableIdContains)
        SourceFileColumnName       = $sourceFileColumnName
        DataColumnPrefix           = $dataColumnPrefix
        ReviewMappings             = $reviewMappings
        NormalizedTimeColumns      = @($normalizedTimeColumns)
        MultiPageMergeMode         = if ([string]::IsNullOrWhiteSpace([string](Get-OptionalProfileValue -ProfileObject $rawProfile -PropertyName 'multiPageMergeMode'))) { 'single' } else { [string](Get-OptionalProfileValue -ProfileObject $rawProfile -PropertyName 'multiPageMergeMode') }
        ProfilePath                = $resolvedProfilePath
    }
}

function Stage-PdfFiles {
    param(
        [Parameter(Mandatory = $true)][string[]]$Files,
        [Parameter(Mandatory = $true)][string]$StagingDirectory
    )

    Ensure-Directory -Path $StagingDirectory

    Get-ChildItem -LiteralPath $StagingDirectory -Filter '*.pdf' -File -ErrorAction SilentlyContinue | ForEach-Object {
        Remove-Item -LiteralPath $_.FullName -Force
    }

    $selectedTargets = @{}
    $normalizedPairs = @()
    foreach ($file in $Files) {
        $sourcePath = [System.IO.Path]::GetFullPath((Resolve-Path -LiteralPath $file).Path)
        $targetPath = Join-Path $StagingDirectory ([System.IO.Path]::GetFileName($sourcePath))
        $targetKey = $targetPath.ToLowerInvariant()

        if ($selectedTargets.ContainsKey($targetKey) -and $selectedTargets[$targetKey] -ne $sourcePath) {
            throw "同名の PDF は同時に処理できません: $([System.IO.Path]::GetFileName($sourcePath))"
        }

        $selectedTargets[$targetKey] = $sourcePath
        $normalizedPairs += [pscustomobject]@{
            SourcePath = $sourcePath
            TargetPath = $targetPath
            TargetKey  = $targetKey
        }
    }

    $staged = @()
    foreach ($pair in $normalizedPairs) {
        if ($pair.SourcePath -ne $pair.TargetPath) {
            Copy-Item -LiteralPath $pair.SourcePath -Destination $pair.TargetPath -Force
        }
        $staged += $pair.TargetPath
    }

    if (-not $staged) {
        throw 'PDF の準備に失敗しました。'
    }

    return $staged
}

function Sync-InputStorage {
    param(
        [Parameter(Mandatory = $true)][string[]]$Files,
        [switch]$KeepExisting
    )

    Ensure-Directory -Path $inputDir

    $selectedByName = @{}
    foreach ($file in $Files) {
        $sourcePath = [System.IO.Path]::GetFullPath((Resolve-Path -LiteralPath $file).Path)
        $fileName = [System.IO.Path]::GetFileName($sourcePath)
        $fileKey = $fileName.ToLowerInvariant()

        if ($selectedByName.ContainsKey($fileKey) -and $selectedByName[$fileKey] -ne $sourcePath) {
            throw "同名の PDF は同時に処理できません: $fileName"
        }

        $selectedByName[$fileKey] = $sourcePath
    }

    if (-not $KeepExisting) {
        Get-ChildItem -LiteralPath $inputDir -Filter '*.pdf' -File -ErrorAction SilentlyContinue | ForEach-Object {
            if (-not $selectedByName.ContainsKey($_.Name.ToLowerInvariant())) {
                Remove-Item -LiteralPath $_.FullName -Force
            }
        }
    }

    $stored = @()
    foreach ($fileKey in $selectedByName.Keys) {
        $sourcePath = $selectedByName[$fileKey]
        $targetPath = Join-Path $inputDir ([System.IO.Path]::GetFileName($sourcePath))
        if ($sourcePath -ne [System.IO.Path]::GetFullPath($targetPath)) {
            Copy-Item -LiteralPath $sourcePath -Destination $targetPath -Force
        }
        $stored += $targetPath
    }

    return @($stored | Sort-Object)
}

function Ensure-Template {
    param([switch]$ForceRebuild)

    if ($ForceRebuild -or -not (Test-Path -LiteralPath $templatePath)) {
        Write-Log "Excel テンプレートを生成しています: $templatePath"
        & $buildTemplateScript -TemplatePath $templatePath -TemplateVariant $VersionMode
    }

    if (-not (Test-Path -LiteralPath $templatePath)) {
        throw "テンプレートの作成に失敗しました: $templatePath"
    }
}

function Copy-TemplateToRuntime {
    Ensure-Directory -Path $script:runRuntimeDir
    $runtimePath = Join-Path $script:runRuntimeDir "PDF2Excel_runtime_$timestamp.xlsm"
    Copy-Item -LiteralPath $templatePath -Destination $runtimePath -Force
    Register-SensitivePaths -Paths @($runtimePath)
    return $runtimePath
}

# ============================================================
# Section: Power Query Formula Builders
# ============================================================

function Get-ReviewQueryFormulaV2 {
    param([Parameter(Mandatory = $true)]$Profile)

    $reviewColumns = @('FileName', 'Page', 'PersonRaw', 'SiteRaw')
    $reviewColumns += @(Get-NormalizedTimeColumnDefinitions -Profile $Profile -VersionMode 'v2' | ForEach-Object { $_.ReviewRawColumnName })
    $reviewColumns += 'Reason'
    $reviewColumns += 'ReasonCategory'
    $reviewColumnsLiteral = ConvertTo-MTextListLiteral -Values $reviewColumns

@"
let
    ReviewColumns = $reviewColumnsLiteral,
    Source = PDF2Excel_Staging,
    ReviewRows = Table.SelectRows(Source, each [Review] <> null),
    Expanded =
        if Table.RowCount(ReviewRows) = 0 then
            #table(ReviewColumns, {})
        else
            Table.ExpandTableColumn(ReviewRows, "Review", ReviewColumns, ReviewColumns)
in
    Expanded
"@
}

function Get-FileSummaryQueryFormulaV2 {
@"
let
    Source = PDF2Excel_Staging,
    Selected =
        Table.SelectColumns(
            Source,
            {"Name", "IsError", "OutputRowCount", "PageCount", "ReviewCount", "ErrorCategory", "ErrorCode", "UserMessage", "CandidateColumns", "CandidateRows", "SelectedTableKind", "SelectedTableName"},
            MissingField.UseNull
        ),
    WithStatus = Table.AddColumn(Selected, "Status", each if [IsError] = true then "失敗" else "成功", type text),
    Reordered =
        Table.ReorderColumns(
            WithStatus,
            {"Name", "Status", "PageCount", "OutputRowCount", "ReviewCount", "ErrorCategory", "ErrorCode", "UserMessage", "CandidateColumns", "CandidateRows", "SelectedTableKind", "SelectedTableName"},
            MissingField.UseNull
        ),
    Renamed =
        Table.RenameColumns(
            Reordered,
            {
                {"Name", "FileName"},
                {"PageCount", "MergedPages"},
                {"OutputRowCount", "ImportedRows"},
                {"ReviewCount", "ReviewRows"},
                {"SelectedTableKind", "TableKind"},
                {"SelectedTableName", "TableName"}
            }
        )
in
    Renamed
"@
}

function Get-StagingQueryFormulaV2 {
    param(
        [Parameter(Mandatory = $true)][string]$InputPath,
        [Parameter(Mandatory = $true)]$Profile
    )

    $escapedPath = Escape-MString -Value $InputPath
    $outputColumnsLiteral = ConvertTo-MTextListLiteral -Values (Get-ProfileOutputColumnNames -Profile $Profile)
    $dataColumnsLiteral = ConvertTo-MTextListLiteral -Values @((1..$Profile.ExpectedColumns | ForEach-Object { '{0}{1}' -f $Profile.DataColumnPrefix, $_ }))
    $preferredKindsLiteral = ConvertTo-MTextListLiteral -Values @($Profile.PreferredTableKinds | ForEach-Object { $_.ToUpperInvariant() })
    $preferredNamesLiteral = ConvertTo-MTextListLiteral -Values @($Profile.PreferredTableNameContains | ForEach-Object { $_.ToUpperInvariant() })
    $preferredIdsLiteral = ConvertTo-MTextListLiteral -Values @($Profile.PreferredTableIdContains | ForEach-Object { $_.ToUpperInvariant() })
    $sourceFileColumnNameLiteral = Escape-MString -Value $Profile.SourceFileColumnName
    $dataColumnPrefixLiteral = Escape-MString -Value $Profile.DataColumnPrefix
    $allowMoreColumnsLiteral = ConvertTo-MLogicalLiteral -Value $Profile.AllowMoreColumns
    $reviewPersonColumnLiteral = if ($null -eq $Profile.ReviewMappings.PersonColumn) { 'null' } else { [string]$Profile.ReviewMappings.PersonColumn }
    $reviewSiteColumnLiteral = if ($null -eq $Profile.ReviewMappings.SiteColumn) { 'null' } else { [string]$Profile.ReviewMappings.SiteColumn }
    $reviewInTimeColumnLiteral = if ($null -eq $Profile.ReviewMappings.InTimeColumn) { 'null' } else { [string]$Profile.ReviewMappings.InTimeColumn }
    $reviewOutTimeColumnLiteral = if ($null -eq $Profile.ReviewMappings.OutTimeColumn) { 'null' } else { [string]$Profile.ReviewMappings.OutTimeColumn }
    $normalizedTimeDefinitions = @(
        Get-NormalizedTimeColumnDefinitions -Profile $Profile -VersionMode 'v2' |
            ForEach-Object {
                [pscustomobject]@{
                    SourceColumnName    = $_.SourceColumnName
                    DisplayName         = $_.DisplayName
                    ReviewRawColumnName = $_.ReviewRawColumnName
                }
            }
    )
    $normalizedTimeDefinitionsLiteral = ConvertTo-MRecordListLiteral -Records $normalizedTimeDefinitions
    $reviewColumnsLiteral = ConvertTo-MTextListLiteral -Values @('FileName', 'Page', 'PersonRaw', 'SiteRaw') + @($normalizedTimeDefinitions | ForEach-Object { $_.ReviewRawColumnName }) + @('Reason')

@"
let
    ExpectedColumns = $($Profile.ExpectedColumns),
    HeaderRowsToSkip = $($Profile.HeaderRowsToSkip),
    TargetRowCount = $($Profile.TargetRowCount),
    AllowMoreColumns = $allowMoreColumnsLiteral,
    SourceFileColumnName = "$sourceFileColumnNameLiteral",
    DataColumnPrefix = "$dataColumnPrefixLiteral",
    OutputColumns = $outputColumnsLiteral,
    DataColumns = $dataColumnsLiteral,
    PreferredKinds = $preferredKindsLiteral,
    PreferredNames = $preferredNamesLiteral,
    PreferredIds = $preferredIdsLiteral,
    ReviewPersonColumn = $reviewPersonColumnLiteral,
    ReviewSiteColumn = $reviewSiteColumnLiteral,
    ReviewInTimeColumn = $reviewInTimeColumnLiteral,
    ReviewOutTimeColumn = $reviewOutTimeColumnLiteral,
    NormalizeText = (value as any) as text =>
        let
            textValue = if value = null then "" else Text.From(value),
            replacedBreaks = Text.Replace(Text.Replace(Text.Replace(textValue, "#(cr,lf)", " "), "#(lf)", " "), "#(cr)", " "),
            replacedWideSpace = Text.Replace(replacedBreaks, "　", " "),
            tokens = List.Select(Text.Split(replacedWideSpace, " "), each _ <> ""),
            normalized = Text.Trim(Text.Combine(tokens, " "))
        in
            normalized,
    GetHeaderSignature = (tableValue as table) as text =>
        let
            headerTable = Table.FirstN(tableValue, HeaderRowsToSkip),
            rows = Table.ToRows(headerTable),
            flattened = List.Combine(List.Transform(rows, each List.Transform(_, each NormalizeText(_)))),
            signature = Text.Combine(flattened, "|")
        in
            signature,
    GetPageNumber = (tableId as nullable text, tableName as nullable text) as nullable number =>
        let
            combined = Text.Upper(Text.Combine(List.RemoveNulls({tableId, tableName}), " ")),
            tokens = List.Select(Text.SplitAny(combined, " _-:/\[]()"), each _ <> ""),
            pageTokens = List.Select(tokens, each Text.StartsWith(_, "PAGE") and Text.Length(Text.Select(_, {"0".."9"})) > 0),
            firstPageToken = if List.Count(pageTokens) > 0 then List.First(pageTokens) else null,
            numericTokens = List.Select(tokens, each Text.Length(_) > 0 and Text.Length(Text.Select(_, {"0".."9"})) = Text.Length(_)),
            fallbackToken = if List.Count(numericTokens) > 0 then List.First(numericTokens) else null,
            digits = if firstPageToken <> null then Text.Select(firstPageToken, {"0".."9"}) else if fallbackToken <> null then fallbackToken else ""
        in
            if digits = "" then null else Number.FromText(digits),
    NormalizeDataTable = (tableValue as table, pageNumber as nullable number) as table =>
        let
            dataOnly = Table.Skip(tableValue, HeaderRowsToSkip),
            originalColumns = Table.ColumnNames(dataOnly),
            renamed = Table.RenameColumns(dataOnly, List.Transform(List.Positions(originalColumns), each {originalColumns{_}, DataColumnPrefix & Text.From(_ + 1)}), MissingField.Ignore),
            renamedCount = Table.ColumnCount(renamed),
            missingColumns = if renamedCount >= ExpectedColumns then {} else List.Transform({renamedCount + 1 .. ExpectedColumns}, each DataColumnPrefix & Text.From(_)),
            padded = List.Accumulate(missingColumns, renamed, (state, columnName) => Table.AddColumn(state, columnName, each null, type text)),
            selected = Table.SelectColumns(padded, DataColumns, MissingField.UseNull),
            normalized = Table.TransformColumns(selected, List.Transform(Table.ColumnNames(selected), each {_, NormalizeText, type text})),
            withPage = Table.AddColumn(normalized, "__PageNumber", each pageNumber, type nullable number)
        in
            withPage,
    CountTimeLikeTokens = (value as any) as number =>
        let
            normalized = NormalizeText(value),
            tokens = List.Select(Text.Split(normalized, " "), each _ <> ""),
            count = List.Count(List.Select(tokens, each Text.Contains(_, ":") or Text.Contains(_, "：") or Text.Contains(_, "時")))
        in
            count,
    BuildReviewTable = (dataTable as nullable table, fileName as text, noteText as nullable text, noteCategory as nullable text) as table =>
        let
            empty = #table({"FileName", "Page", "PersonRaw", "SiteRaw", "InRaw", "OutRaw", "Reason"}, {}),
            personColumnName = if ReviewPersonColumn = null then null else DataColumnPrefix & Text.From(ReviewPersonColumn),
            siteColumnName = if ReviewSiteColumn = null then null else DataColumnPrefix & Text.From(ReviewSiteColumn),
            inColumnName = if ReviewInTimeColumn = null then null else DataColumnPrefix & Text.From(ReviewInTimeColumn),
            outColumnName = if ReviewOutTimeColumn = null then null else DataColumnPrefix & Text.From(ReviewOutTimeColumn),
            rowReviewRecords =
                if dataTable = null then
                    {}
                else
                    List.RemoveNulls(
                        List.Transform(
                            Table.ToRecords(dataTable),
                            each
                                let
                                    personRaw = if personColumnName = null then "" else Record.FieldOrDefault(_, personColumnName, ""),
                                    siteRaw = if siteColumnName = null then "" else Record.FieldOrDefault(_, siteColumnName, ""),
                                    inRaw = if inColumnName = null then "" else Record.FieldOrDefault(_, inColumnName, ""),
                                    outRaw = if outColumnName = null then "" else Record.FieldOrDefault(_, outColumnName, ""),
                                    reasons =
                                        List.RemoveNulls(
                                            {
                                                if Text.Length(inRaw) = 0 or Text.Length(outRaw) = 0 then "入退場時刻の片方が空です" else null,
                                                if CountTimeLikeTokens(inRaw) > 1 or CountTimeLikeTokens(outRaw) > 1 then "同一セルに複数の時刻らしき文字列があります" else null,
                                                if Text.Contains(Text.From(personRaw), "#(lf)") or Text.Contains(Text.From(siteRaw), "#(lf)") or Text.Contains(Text.From(personRaw), "　") or Text.Contains(Text.From(siteRaw), "　") then "改行または全角空白を含むため確認してください" else null
                                            }
                                        ),
                                    reason = Text.Combine(reasons, " / ")
                                in
                                    if Text.Length(reason) = 0 then
                                        null
                                    else
                                        [
                                            FileName = fileName,
                                            Page = Record.FieldOrDefault(_, "__PageNumber", null),
                                            PersonRaw = personRaw,
                                            SiteRaw = siteRaw,
                                            InRaw = inRaw,
                                            OutRaw = outRaw,
                                            Reason = reason
                                        ]
                        )
                    ),
            rowReviews =
                if List.Count(rowReviewRecords) = 0 then
                    empty
                else
                    Table.FromRecords(
                        rowReviewRecords,
                        type table [FileName = text, Page = nullable number, PersonRaw = text, SiteRaw = text, InRaw = text, OutRaw = text, Reason = text]
                    ),
            pageNote =
                if noteText = null or Text.Length(noteText) = 0 then
                    empty
                else
                    #table({"FileName", "Page", "PersonRaw", "SiteRaw", "InRaw", "OutRaw", "Reason"}, {{fileName, null, "", "", "", "", noteText}}),
            combined =
                if Table.RowCount(rowReviews) = 0 then
                    pageNote
                else if Table.RowCount(pageNote) = 0 then
                    rowReviews
                else
                    Table.Combine({rowReviews, pageNote})
        in
            combined,
    Source = Folder.Files("$escapedPath"),
    PdfFiles = Table.SelectRows(Source, each Text.Lower([Extension]) = ".pdf"),
    KeepColumns = Table.SelectColumns(PdfFiles, {"Name", "Extension", "Folder Path", "Content"}),
    WithProcessed =
        Table.AddColumn(
            KeepColumns,
            "Processed",
            each
                let
                    fileName = [Name],
                    processingTry =
                        try
                            let
                                pdfTry = try Pdf.Tables([Content], [Implementation = "1.3", MultiPageTables = false]),
                                pdfTables = if pdfTry[HasError] then null else pdfTry[Value],
                                pdfError = if pdfTry[HasError] then try Error.Message(pdfTry[Error]) otherwise "Pdf.Tables の読み取りに失敗しました。" else null,
                                candidateRecords = if pdfTables = null then {} else Table.ToRecords(pdfTables),
                                indexedCandidates = List.Transform(List.Positions(candidateRecords), each Record.AddField(candidateRecords{_}, "CandidateIndex", _ + 1)),
                                scoredCandidates =
                                    List.Transform(
                                        indexedCandidates,
                                        each
                                            let
                                                dataTry = try Record.Field(_, "Data"),
                                                dataValue = if dataTry[HasError] then null else dataTry[Value],
                                                columnCount = if dataValue = null then null else Table.ColumnCount(dataValue),
                                                rowCount = if dataValue = null then null else Table.RowCount(dataValue),
                                                tableId = try Text.From(Record.Field(_, "Id")) otherwise "",
                                                tableKind = try Text.From(Record.Field(_, "Kind")) otherwise "",
                                                tableName = try Text.From(Record.Field(_, "Name")) otherwise "",
                                                normalizedKind = Text.Upper(tableKind),
                                                normalizedName = Text.Upper(tableName),
                                                normalizedId = Text.Upper(tableId),
                                                headerSignature = if dataValue = null or rowCount = null or rowCount <= HeaderRowsToSkip then "" else GetHeaderSignature(dataValue),
                                                pageNumber = GetPageNumber(tableId, tableName),
                                                dataRowCount = if rowCount = null then 0 else Number.Max(rowCount - HeaderRowsToSkip, 0),
                                                kindBonus = if List.Contains(PreferredKinds, normalizedKind) then -250 else 0,
                                                nameBonus = if List.Count(PreferredNames) = 0 then 0 else if List.AnyTrue(List.Transform(PreferredNames, each Text.Contains(normalizedName, _))) then -120 else 0,
                                                idBonus = if List.Count(PreferredIds) = 0 then 0 else if List.AnyTrue(List.Transform(PreferredIds, each Text.Contains(normalizedId, _))) then -120 else 0,
                                                score =
                                                    if dataValue = null or rowCount = null or columnCount = null then
                                                        999999
                                                    else
                                                        Number.Abs(columnCount - ExpectedColumns) * 1000 +
                                                        Number.Abs(dataRowCount - TargetRowCount) * 10 +
                                                        kindBonus + nameBonus + idBonus
                                            in
                                                [
                                                    Data = dataValue,
                                                    ColumnCount = columnCount,
                                                    RowCount = rowCount,
                                                    DataRowCount = dataRowCount,
                                                    Score = score,
                                                    TableId = tableId,
                                                    TableKind = tableKind,
                                                    TableName = tableName,
                                                    HeaderSignature = headerSignature,
                                                    PageNumber = pageNumber,
                                                    CandidateIndex = Record.Field(_, "CandidateIndex")
                                                ]
                                    ),
                                viableCandidates =
                                    List.Select(
                                        scoredCandidates,
                                        each
                                            [Data] <> null and [RowCount] <> null and [RowCount] > HeaderRowsToSkip and (
                                                (AllowMoreColumns = true and [ColumnCount] >= ExpectedColumns) or
                                                (AllowMoreColumns = false and [ColumnCount] = ExpectedColumns)
                                            )
                                    ),
                                sortedCandidates =
                                    List.Sort(
                                        viableCandidates,
                                        (left, right) =>
                                            if left[Score] < right[Score] then
                                                -1
                                            else if left[Score] > right[Score] then
                                                1
                                            else if left[DataRowCount] > right[DataRowCount] then
                                                -1
                                            else if left[DataRowCount] < right[DataRowCount] then
                                                1
                                            else
                                                0
                                    ),
                                primaryCandidate = if List.Count(sortedCandidates) = 0 then null else List.First(sortedCandidates),
                                chosenGroupTables =
                                    if primaryCandidate = null then
                                        {}
                                    else
                                        List.Select(
                                            viableCandidates,
                                            each [HeaderSignature] = primaryCandidate[HeaderSignature] and [ColumnCount] = primaryCandidate[ColumnCount]
                                        ),
                                sortedChosenTables =
                                    List.Sort(
                                        chosenGroupTables,
                                        (left, right) =>
                                            if left[PageNumber] = null and right[PageNumber] = null then
                                                if left[CandidateIndex] < right[CandidateIndex] then -1 else if left[CandidateIndex] > right[CandidateIndex] then 1 else 0
                                            else if left[PageNumber] = null then
                                                1
                                            else if right[PageNumber] = null then
                                                -1
                                            else if left[PageNumber] < right[PageNumber] then
                                                -1
                                            else if left[PageNumber] > right[PageNumber] then
                                                1
                                            else if left[CandidateIndex] < right[CandidateIndex] then
                                                -1
                                            else if left[CandidateIndex] > right[CandidateIndex] then
                                                1
                                            else
                                                0
                                    ),
                                chosenColumns = if primaryCandidate = null then null else primaryCandidate[ColumnCount],
                                chosenRows = if List.Count(sortedChosenTables) = 0 then null else List.Sum(List.Transform(sortedChosenTables, each [DataRowCount])),
                                chosenTable = if List.Count(sortedChosenTables) = 0 then null else List.First(sortedChosenTables),
                                pageMergeNote =
                                    if List.Count(viableCandidates) > List.Count(sortedChosenTables) and List.Count(sortedChosenTables) > 0 then
                                        "同一 PDF 内にヘッダー不一致または列ずれの候補表がありました。"
                                    else
                                        null,
                                failureRecord =
                                    if pdfTry[HasError] then
                                        [
                                            ErrorCode = "PDF_READ_FAILURE",
                                            ErrorCategory = "PDF読込エラー",
                                            UserMessage = "PDF を読み取れませんでした。壊れているか、Excel の PDF 解析で扱えない可能性があります。",
                                            TechnicalDetail = pdfError
                                        ]
                                    else if primaryCandidate = null then
                                        [
                                            ErrorCode = "TABLE_NOT_FOUND",
                                            ErrorCategory = "表検出エラー",
                                            UserMessage = "想定に近い表を検出できませんでした。",
                                            TechnicalDetail = "Pdf.Tables で抽出候補が見つからないか、ヘッダー行のみでした。"
                                        ]
                                    else
                                        null,
                                normalizedDataTables =
                                    if failureRecord <> null then
                                        {}
                                    else
                                        List.Transform(sortedChosenTables, each NormalizeDataTable([Data], [PageNumber])),
                                mergedDataWithPage =
                                    if List.Count(normalizedDataTables) = 0 then
                                        null
                                    else
                                        Table.Combine(normalizedDataTables),
                                resultDataWithoutPage =
                                    if mergedDataWithPage = null then
                                        null
                                    else
                                        Table.RemoveColumns(mergedDataWithPage, {"__PageNumber"}),
                                withFileName =
                                    if resultDataWithoutPage = null then
                                        null
                                    else
                                        Table.AddColumn(resultDataWithoutPage, SourceFileColumnName, each fileName, type text),
                                reordered =
                                    if withFileName = null then
                                        null
                                    else
                                        Table.ReorderColumns(withFileName, OutputColumns, MissingField.UseNull),
                                reviewData = BuildReviewTable(mergedDataWithPage, fileName, pageMergeNote),
                                outputRowCount = if reordered = null then 0 else Table.RowCount(reordered),
                                reviewCount = if reviewData = null then 0 else Table.RowCount(reviewData)
                            in
                                [
                                    IsError = failureRecord <> null,
                                    ErrorCode = if failureRecord = null then null else failureRecord[ErrorCode],
                                    ErrorCategory = if failureRecord = null then null else failureRecord[ErrorCategory],
                                    UserMessage = if failureRecord = null then null else failureRecord[UserMessage],
                                    TechnicalDetail = if failureRecord = null then null else failureRecord[TechnicalDetail],
                                    CandidateColumns = chosenColumns,
                                    CandidateRows = chosenRows,
                                    OutputRowCount = outputRowCount,
                                    ReviewCount = reviewCount,
                                    PageCount = List.Count(sortedChosenTables),
                                    SelectedTableId = if chosenTable = null then null else chosenTable[TableId],
                                    SelectedTableKind = if chosenTable = null then null else chosenTable[TableKind],
                                    SelectedTableName = if chosenTable = null then null else chosenTable[TableName],
                                    Data = reordered,
                                    Review = reviewData
                                ]
                in
                    if processingTry[HasError] then
                        [
                            IsError = true,
                            ErrorCode = "UNEXPECTED_PROCESSING_ERROR",
                            ErrorCategory = "システムエラー",
                            UserMessage = "この PDF の処理中に予期しないエラーが発生しました。",
                            TechnicalDetail = try Error.Message(processingTry[Error]) otherwise "unknown error",
                            CandidateColumns = null,
                            CandidateRows = null,
                            OutputRowCount = 0,
                            ReviewCount = 0,
                            PageCount = 0,
                            SelectedTableId = null,
                            SelectedTableKind = null,
                            SelectedTableName = null,
                            Data = null,
                            Review = null
                        ]
                    else
                        processingTry[Value],
            type record
        ),
    Expanded =
        Table.ExpandRecordColumn(
            WithProcessed,
            "Processed",
            {"IsError", "ErrorCode", "ErrorCategory", "UserMessage", "TechnicalDetail", "CandidateColumns", "CandidateRows", "OutputRowCount", "ReviewCount", "PageCount", "SelectedTableId", "SelectedTableKind", "SelectedTableName", "Data", "Review"},
            {"IsError", "ErrorCode", "ErrorCategory", "UserMessage", "TechnicalDetail", "CandidateColumns", "CandidateRows", "OutputRowCount", "ReviewCount", "PageCount", "SelectedTableId", "SelectedTableKind", "SelectedTableName", "Data", "Review"}
        )
in
    Expanded
"@
}

function Get-StagingQueryFormulaV2Simple {
    param(
        [Parameter(Mandatory = $true)][string]$InputPath,
        [Parameter(Mandatory = $true)]$Profile
    )

    $escapedPath = Escape-MString -Value $InputPath
    $outputColumnsLiteral = ConvertTo-MTextListLiteral -Values (Get-ProfileOutputColumnNames -Profile $Profile)
    $dataColumnsLiteral = ConvertTo-MTextListLiteral -Values @((1..$Profile.ExpectedColumns | ForEach-Object { '{0}{1}' -f $Profile.DataColumnPrefix, $_ }))
    $preferredKindsLiteral = ConvertTo-MTextListLiteral -Values @($Profile.PreferredTableKinds | ForEach-Object { $_.ToUpperInvariant() })
    $preferredNamesLiteral = ConvertTo-MTextListLiteral -Values @($Profile.PreferredTableNameContains | ForEach-Object { $_.ToUpperInvariant() })
    $preferredIdsLiteral = ConvertTo-MTextListLiteral -Values @($Profile.PreferredTableIdContains | ForEach-Object { $_.ToUpperInvariant() })
    $sourceFileColumnNameLiteral = Escape-MString -Value $Profile.SourceFileColumnName
    $dataColumnPrefixLiteral = Escape-MString -Value $Profile.DataColumnPrefix
    $allowMoreColumnsLiteral = ConvertTo-MLogicalLiteral -Value $Profile.AllowMoreColumns
    $reviewPersonColumnLiteral = if ($null -eq $Profile.ReviewMappings.PersonColumn) { 'null' } else { [string]$Profile.ReviewMappings.PersonColumn }
    $reviewSiteColumnLiteral = if ($null -eq $Profile.ReviewMappings.SiteColumn) { 'null' } else { [string]$Profile.ReviewMappings.SiteColumn }
    $reviewInTimeColumnLiteral = if ($null -eq $Profile.ReviewMappings.InTimeColumn) { 'null' } else { [string]$Profile.ReviewMappings.InTimeColumn }
    $reviewOutTimeColumnLiteral = if ($null -eq $Profile.ReviewMappings.OutTimeColumn) { 'null' } else { [string]$Profile.ReviewMappings.OutTimeColumn }
    $normalizedTimeDefinitionsLiteral = ConvertTo-MRecordListLiteral -Records (Get-NormalizedTimeColumnDefinitions -Profile $Profile -VersionMode 'v2')
    $reviewColumnsLiteral = ConvertTo-MTextListLiteral -Values @(
        'FileName'
        'Page'
        'PersonRaw'
        'SiteRaw'
        @((Get-NormalizedTimeColumnDefinitions -Profile $Profile -VersionMode 'v2' | ForEach-Object { $_.ReviewRawColumnName }))
        'Reason'
        'ReasonCategory'
    )

@"
let
    ExpectedColumns = $($Profile.ExpectedColumns),
    HeaderRowsToSkip = $($Profile.HeaderRowsToSkip),
    TargetRowCount = $($Profile.TargetRowCount),
    AllowMoreColumns = $allowMoreColumnsLiteral,
    SourceFileColumnName = "$sourceFileColumnNameLiteral",
    DataColumnPrefix = "$dataColumnPrefixLiteral",
    OutputColumns = $outputColumnsLiteral,
    DataColumns = $dataColumnsLiteral,
    PreferredKinds = $preferredKindsLiteral,
    PreferredNames = $preferredNamesLiteral,
    PreferredIds = $preferredIdsLiteral,
    ReviewPersonColumn = $reviewPersonColumnLiteral,
    ReviewSiteColumn = $reviewSiteColumnLiteral,
    ReviewInTimeColumn = $reviewInTimeColumnLiteral,
    ReviewOutTimeColumn = $reviewOutTimeColumnLiteral,
    NormalizedTimeDefinitions = $normalizedTimeDefinitionsLiteral,
    ReviewColumns = $reviewColumnsLiteral,
    NormalizeText = (value as any) as text =>
        let
            source = if value = null then "" else Text.From(value),
            lineNormalized = Text.Replace(Text.Replace(Text.Replace(source, "#(cr,lf)", " "), "#(lf)", " "), "#(cr)", " "),
            wideSpaceNormalized = Text.Replace(lineNormalized, "　", " "),
            tokens = List.Select(Text.Split(wideSpaceNormalized, " "), each _ <> ""),
            normalized = Text.Trim(Text.Combine(tokens, " "))
        in normalized,
    NormalizeHeaderCell = (value as any) as text =>
        let
            normalized = Text.Upper(NormalizeText(value)),
            tokens = List.Select(Text.Split(normalized, " "), each _ <> ""),
            filteredTokens =
                List.Select(
                    tokens,
                    each
                        not Text.Contains(_, "ページ") and
                        not Text.Contains(_, "対象月") and
                        not Text.Contains(_, "管理者") and
                        not (
                            Text.Contains(_, "年") and
                            Text.Contains(_, "月") and
                            Text.Length(Text.Select(_, {"0".."9", "年", "月"})) = Text.Length(_)
                        )
                ),
            compact = Text.Trim(Text.Combine(filteredTokens, " "))
        in
            compact,
    CountTimeLikeTokens = (value as any) as number =>
        let
            normalized = NormalizeText(value),
            tokens = List.Select(Text.Split(normalized, " "), each _ <> ""),
            count = List.Count(List.Select(tokens, each Text.Contains(_, ":") or Text.Contains(_, "：") or Text.Contains(_, "時")))
        in count,
    GetHeaderSignature = (tableValue as table) as text =>
        let
            headerTable = Table.FirstN(tableValue, HeaderRowsToSkip),
            rows = Table.ToRows(headerTable),
            flattened = List.Combine(List.Transform(rows, each List.Transform(_, each NormalizeText(_)))),
            signature = Text.Combine(flattened, "|")
        in signature,
    GetCanonicalHeaderSignature = (tableValue as table) as text =>
        let
            requestedHeaderRows = if HeaderRowsToSkip <= 0 then 1 else HeaderRowsToSkip,
            rowLimit = if Table.RowCount(tableValue) < requestedHeaderRows then Table.RowCount(tableValue) else requestedHeaderRows,
            headerTable = Table.FirstN(tableValue, rowLimit),
            rows = Table.ToRows(headerTable),
            flattened = List.Combine(List.Transform(rows, each List.Transform(_, each NormalizeHeaderCell(_)))),
            filtered = List.Select(flattened, each _ <> ""),
            signature = Text.Combine(filtered, "|")
        in signature,
    GetPageNumber = (tableId as nullable text, tableName as nullable text) as nullable number =>
        let
            combined = Text.Upper(Text.Combine(List.RemoveNulls({tableId, tableName}), " ")),
            tokens = List.Select(Text.SplitAny(combined, " _-:/\[]()"), each _ <> ""),
            pageTokens = List.Select(tokens, each Text.StartsWith(_, "PAGE") and Text.Length(Text.Select(_, {"0".."9"})) > 0),
            firstPageToken = if List.Count(pageTokens) > 0 then List.First(pageTokens) else null,
            numericTokens = List.Select(tokens, each Text.Length(_) > 0 and Text.Length(Text.Select(_, {"0".."9"})) = Text.Length(_)),
            fallbackToken = if List.Count(numericTokens) > 0 then List.First(numericTokens) else null,
            digits = if firstPageToken <> null then Text.Select(firstPageToken, {"0".."9"}) else if fallbackToken <> null then fallbackToken else ""
        in
            if digits = "" then null else Number.FromText(digits),
    BuildResultTable = (tableValue as nullable table, pageNumber as nullable number, columnStartNumber as number, padToExpected as logical) as nullable table =>
        if tableValue = null then null else
        let
            dataOnly = Table.Skip(tableValue, HeaderRowsToSkip),
            originalColumns = Table.ColumnNames(dataOnly),
            renamed = Table.RenameColumns(dataOnly, List.Transform(List.Positions(originalColumns), each {originalColumns{_}, DataColumnPrefix & Text.From(columnStartNumber + _)}), MissingField.Ignore),
            renamedCount = Table.ColumnCount(renamed),
            currentColumns = Table.ColumnNames(renamed),
            padded =
                if padToExpected then
                    let
                        missingColumns = if renamedCount >= ExpectedColumns then {} else List.Transform({renamedCount + 1 .. ExpectedColumns}, each DataColumnPrefix & Text.From(_)),
                        paddedTable = List.Accumulate(missingColumns, renamed, (state, columnName) => Table.AddColumn(state, columnName, each null, type text))
                    in
                        paddedTable
                else
                    renamed,
            selected = if padToExpected then Table.SelectColumns(padded, DataColumns, MissingField.UseNull) else Table.SelectColumns(padded, currentColumns, MissingField.UseNull),
            normalized = Table.TransformColumns(selected, List.Transform(Table.ColumnNames(selected), each {_, NormalizeText, type text})),
            withPage = if padToExpected then Table.AddColumn(normalized, "__PageNumber", each pageNumber, type nullable number) else normalized
        in withPage,
    BuildReviewRecord = (rowRecord as record, fileName as text, reasonText as nullable text, reasonCategories as nullable list) as record =>
        let
            personColumnName = if ReviewPersonColumn = null then null else DataColumnPrefix & Text.From(ReviewPersonColumn),
            siteColumnName = if ReviewSiteColumn = null then null else DataColumnPrefix & Text.From(ReviewSiteColumn),
            baseRecord = [
                FileName = fileName,
                Page = Record.FieldOrDefault(rowRecord, "__PageNumber", null),
                PersonRaw = if personColumnName = null then "" else Record.FieldOrDefault(rowRecord, personColumnName, ""),
                SiteRaw = if siteColumnName = null then "" else Record.FieldOrDefault(rowRecord, siteColumnName, ""),
                Reason = if reasonText = null then "" else reasonText,
                ReasonCategory = if reasonCategories = null then "" else Text.Combine(List.Distinct(List.RemoveNulls(reasonCategories)), ",")
            ],
            withTimeFields =
                List.Accumulate(
                    NormalizedTimeDefinitions,
                    baseRecord,
                    (state, current) => Record.AddField(state, current[ReviewRawColumnName], Record.FieldOrDefault(rowRecord, current[SourceColumnName], ""))
                )
        in
            withTimeFields,
    BuildReviewTable = (dataTable as nullable table, fileName as text, noteText as nullable text, noteCategory as nullable text) as table =>
        let
            empty = #table(ReviewColumns, {}),
            rowRecords =
                if dataTable = null then
                    {}
                else
                    List.Transform(
                        Table.ToRecords(dataTable),
                        each
                            let
                                timeReasonRecords =
                                    List.RemoveNulls(
                                        List.Transform(
                                            NormalizedTimeDefinitions,
                                            (definition) =>
                                                let
                                                    rawText = Text.From(Record.FieldOrDefault(_, definition[SourceColumnName], ""))
                                                in
                                                    if CountTimeLikeTokens(rawText) > 1 then [Text = definition[DisplayName] & ": 同一セルに複数の時刻らしき文字列があります", Category = "TIME_MULTI"] else null
                                        )
                                    ),
                                pairReasonRecords =
                                    List.RemoveNulls(
                                        List.Transform(
                                            {0..Number.IntegerDivide(List.Count(NormalizedTimeDefinitions) - 1, 2)},
                                            (pairIndex) =>
                                                let
                                                    leftIndex = pairIndex * 2,
                                                    rightIndex = leftIndex + 1,
                                                    leftDefinition = if leftIndex < List.Count(NormalizedTimeDefinitions) then NormalizedTimeDefinitions{leftIndex} else null,
                                                    rightDefinition = if rightIndex < List.Count(NormalizedTimeDefinitions) then NormalizedTimeDefinitions{rightIndex} else null,
                                                    leftRaw = if leftDefinition = null then "" else Text.From(Record.FieldOrDefault(_, leftDefinition[SourceColumnName], "")),
                                                    rightRaw = if rightDefinition = null then "" else Text.From(Record.FieldOrDefault(_, rightDefinition[SourceColumnName], "")),
                                                    leftHasValue = Text.Length(leftRaw) > 0,
                                                    rightHasValue = Text.Length(rightRaw) > 0
                                                in
                                                    if leftDefinition = null or rightDefinition = null then null else if leftHasValue <> rightHasValue then [Text = leftDefinition[DisplayName] & "/" & rightDefinition[DisplayName] & ": 片側の時刻だけ埋まっています。", Category = "TIME_MISSING"] else null
                                        )
                                    ),
                                reasonRecords = List.Combine({timeReasonRecords, pairReasonRecords}),
                                reasonText = Text.Combine(List.Transform(reasonRecords, each _[Text]), " / "),
                                reasonCategories = List.Transform(reasonRecords, each _[Category])
                            in
                                BuildReviewRecord(_, fileName, reasonText, reasonCategories)
                    ),
            rowReviews = if List.Count(rowRecords) = 0 then empty else Table.SelectColumns(Table.FromRecords(rowRecords), ReviewColumns, MissingField.UseNull),
            pageNoteRecord =
                if noteText = null or Text.Length(noteText) = 0 then
                    null
                else
                    List.Accumulate(
                        NormalizedTimeDefinitions,
                        [FileName = fileName, Page = null, PersonRaw = "", SiteRaw = "", Reason = noteText, ReasonCategory = if noteCategory = null then "" else noteCategory],
                        (state, current) => Record.AddField(state, current[ReviewRawColumnName], "")
                    ),
            pageNote = if pageNoteRecord = null then empty else Table.SelectColumns(Table.FromRecords({pageNoteRecord}), ReviewColumns, MissingField.UseNull)
        in
            if Table.RowCount(rowReviews) = 0 then pageNote else if Table.RowCount(pageNote) = 0 then rowReviews else Table.Combine({rowReviews, pageNote}),
    CombineHorizontalTables = (candidateGroup as list) as nullable table =>
        let
            prepared =
                List.Transform(
                    List.Positions(candidateGroup),
                    each
                        let
                            priorColumns = if _ = 0 then 0 else List.Sum(List.Transform(List.FirstN(candidateGroup, _), each [ColumnCount])),
                            segment = BuildResultTable(candidateGroup{_}[Data], null, priorColumns + 1, false)
                        in
                            if segment = null then null else Table.AddIndexColumn(segment, "__JoinIndex", 0, 1, Int64.Type)
                ),
            validPrepared = List.RemoveNulls(prepared),
            rowCounts = List.Transform(validPrepared, each Table.RowCount(_)),
            canMerge = List.Count(validPrepared) > 1 and List.Count(List.Distinct(rowCounts)) = 1,
            merged =
                if canMerge then
                    List.Accumulate(
                        List.Skip(validPrepared, 1),
                        List.First(validPrepared),
                        (state, current) => Table.Join(state, "__JoinIndex", current, "__JoinIndex", JoinKind.Inner)
                    )
                else
                    null,
            withoutJoin = if merged = null then null else Table.RemoveColumns(merged, {"__JoinIndex"}, MissingField.Ignore),
            selected = if withoutJoin = null then null else Table.SelectColumns(withoutJoin, DataColumns, MissingField.UseNull),
            withPage = if selected = null then null else Table.AddColumn(selected, "__PageNumber", each null, type nullable number)
        in
            withPage,
    SortCandidatesByPageAndIndex = (candidates as list) as list =>
        if List.Count(candidates) = 0 then
            {}
        else
            List.Sort(
                candidates,
                (left, right) =>
                    if left[PageNumber] = null and right[PageNumber] = null then
                        if left[CandidateIndex] < right[CandidateIndex] then -1 else if left[CandidateIndex] > right[CandidateIndex] then 1 else 0
                    else if left[PageNumber] = null then
                        1
                    else if right[PageNumber] = null then
                        -1
                    else if left[PageNumber] < right[PageNumber] then
                        -1
                    else if left[PageNumber] > right[PageNumber] then
                        1
                    else if left[CandidateIndex] < right[CandidateIndex] then
                        -1
                    else if left[CandidateIndex] > right[CandidateIndex] then
                        1
                    else
                        0
            ),
    GetCandidateGroupKey = (candidate as record) as text => candidate[CanonicalHeaderSignature] & "|" & Text.Upper(candidate[TableKind]),
    IsContinuousCandidateSequence = (candidates as list) as logical =>
        let
            sorted = SortCandidatesByPageAndIndex(candidates),
            indices = List.Transform(sorted, each [CandidateIndex]),
            pageNumbers = List.Transform(sorted, each [PageNumber]),
            nonNullPageNumbers = List.RemoveNulls(pageNumbers),
            indexContinuous =
                if List.Count(indices) <= 1 then
                    true
                else
                    List.AllTrue(List.Transform({1..List.Count(indices) - 1}, each indices{_} = indices{_ - 1} + 1)),
            pageContinuous =
                if List.Count(nonNullPageNumbers) = 0 then
                    indexContinuous
                else if List.Count(nonNullPageNumbers) <> List.Count(pageNumbers) then
                    false
                else
                    List.AllTrue(List.Transform({1..List.Count(nonNullPageNumbers) - 1}, each nonNullPageNumbers{_} = nonNullPageNumbers{_ - 1} + 1))
        in
            pageContinuous,
    FindHorizontalMergeSequencesInGroup = (candidateGroup as list) as list =>
        let
            sortedGroup = SortCandidatesByPageAndIndex(candidateGroup),
            candidateCount = List.Count(sortedGroup),
            startPositions = if candidateCount = 0 then {} else {0..candidateCount - 1},
            sequences =
                List.RemoveNulls(
                    List.Transform(
                        startPositions,
                        (startIndex) =>
                            let
                                tail = List.Skip(sortedGroup, startIndex),
                                accumulator =
                                    List.Accumulate(
                                        tail,
                                        [Items = {}, TotalColumns = 0, Signatures = {}, Valid = true],
                                        (state, current) =>
                                            if state[Valid] = false or state[TotalColumns] >= ExpectedColumns then
                                                state
                                            else
                                                let
                                                    currentSignature = current[CanonicalHeaderSignature],
                                                    nextTotalColumns = state[TotalColumns] + current[ColumnCount],
                                                    isCompatible =
                                                        currentSignature <> "" and
                                                        not List.Contains(state[Signatures], currentSignature) and
                                                        nextTotalColumns <= ExpectedColumns,
                                                    nextState =
                                                        if isCompatible then
                                                            [
                                                                Items = state[Items] & {current},
                                                                TotalColumns = nextTotalColumns,
                                                                Signatures = state[Signatures] & {currentSignature},
                                                                Valid = true
                                                            ]
                                                        else
                                                            [Items = state[Items], TotalColumns = state[TotalColumns], Signatures = state[Signatures], Valid = false]
                                                in
                                                    nextState
                                    )
                            in
                                if accumulator[TotalColumns] = ExpectedColumns and List.Count(accumulator[Items]) > 1 then accumulator[Items] else null
                    )
                )
        in
            sequences,
    GetHorizontalGroupKey = (candidate as record) as text =>
        Text.Upper(candidate[TableKind]) & "|" & Text.From(candidate[DataRowCount]),
    FindHorizontalMergeSequences = (candidates as list) as list =>
        let
            sortedCandidates = SortCandidatesByPageAndIndex(candidates),
            groupedCandidates =
                if List.Count(sortedCandidates) = 0 then
                    #table({"GroupKey", "Candidates"}, {})
                else
                    Table.Group(
                        Table.FromRecords(List.Transform(sortedCandidates, each [GroupKey = GetHorizontalGroupKey(_), Candidate = _])),
                        {"GroupKey"},
                        {{"Candidates", each [Candidate], type list}}
                    ),
            sequences =
                if Table.RowCount(groupedCandidates) = 0 then
                    {}
                else
                    List.Combine(
                        List.Transform(
                            Table.ToRecords(groupedCandidates),
                            each FindHorizontalMergeSequencesInGroup([Candidates])
                        )
                    )
        in
            sequences,
    ScoreCandidates = (tableValue as nullable table, preferMoreRows as logical, allowPartialColumns as logical) as list =>
        let
            records = if tableValue = null then {} else Table.ToRecords(tableValue),
            indexedRecords = List.Transform(List.Positions(records), each Record.AddField(records{_}, "CandidateIndex", _ + 1)),
            scored = List.Transform(indexedRecords, each
                let
                    dataTry = try Record.Field(_, "Data"),
                    dataValue = if dataTry[HasError] then null else dataTry[Value],
                    columnCount = if dataValue = null then null else Table.ColumnCount(dataValue),
                    rowCount = if dataValue = null then null else Table.RowCount(dataValue),
                    tableId = try Text.From(Record.Field(_, "Id")) otherwise "",
                    tableKind = try Text.From(Record.Field(_, "Kind")) otherwise "",
                    tableName = try Text.From(Record.Field(_, "Name")) otherwise "",
                    normalizedKind = Text.Upper(tableKind),
                    normalizedName = Text.Upper(tableName),
                    normalizedId = Text.Upper(tableId),
                    headerSignature = if dataValue = null or rowCount = null or rowCount <= HeaderRowsToSkip then "" else GetCanonicalHeaderSignature(dataValue),
                    pageNumber = GetPageNumber(tableId, tableName),
                    dataRowCount = if rowCount = null then 0 else rowCount - HeaderRowsToSkip,
                    kindBonus = if List.Contains(PreferredKinds, normalizedKind) then -250 else 0,
                    nameBonus = if List.Count(PreferredNames) = 0 then 0 else if List.AnyTrue(List.Transform(PreferredNames, each Text.Contains(normalizedName, _))) then -120 else 0,
                    idBonus = if List.Count(PreferredIds) = 0 then 0 else if List.AnyTrue(List.Transform(PreferredIds, each Text.Contains(normalizedId, _))) then -120 else 0,
                    score = if dataValue = null or rowCount = null or columnCount = null then 999999 else if preferMoreRows then Number.Abs(columnCount - ExpectedColumns) * 100000 - dataRowCount * 10 + kindBonus + nameBonus + idBonus else Number.Abs(columnCount - ExpectedColumns) * 1000 + Number.Abs(dataRowCount - TargetRowCount) * 10 + kindBonus + nameBonus + idBonus
                in [Data = dataValue, ColumnCount = columnCount, RowCount = rowCount, DataRowCount = dataRowCount, Score = score, TableId = tableId, TableKind = tableKind, TableName = tableName, CanonicalHeaderSignature = headerSignature, PageNumber = pageNumber, CandidateIndex = Record.Field(_, "CandidateIndex")]
            )
        in
            List.Select(
                scored,
                each
                    [Data] <> null and
                    [RowCount] <> null and
                    [RowCount] > HeaderRowsToSkip and
                    (
                        if allowPartialColumns then
                            [ColumnCount] > 0 and [ColumnCount] <= ExpectedColumns
                        else if AllowMoreColumns = true then
                            [ColumnCount] >= ExpectedColumns
                        else
                            [ColumnCount] = ExpectedColumns
                    )
            ),
    Source = Folder.Files("$escapedPath"),
    PdfFiles = Table.SelectRows(Source, each Text.Lower([Extension]) = ".pdf"),
    KeepColumns = Table.SelectColumns(PdfFiles, {"Name", "Extension", "Folder Path", "Content"}),
    WithProcessed = Table.AddColumn(KeepColumns, "Processed", each
        let
            fileName = [Name],
            processTry = try
                let
                    mergedTry = try Pdf.Tables([Content], [Implementation = "1.3", MultiPageTables = true]),
                    pageTry = try Pdf.Tables([Content], [Implementation = "1.3", MultiPageTables = false]),
                    mergedCandidatesRaw = ScoreCandidates(if mergedTry[HasError] then null else mergedTry[Value], true, false),
                    pageCandidatesRaw = ScoreCandidates(if pageTry[HasError] then null else pageTry[Value], false, true),
                    mergedCandidates =
                        if List.Count(PreferredKinds) = 0 then
                            mergedCandidatesRaw
                        else
                            List.Select(mergedCandidatesRaw, each List.Contains(PreferredKinds, Text.Upper([TableKind]))),
                    pageCandidates =
                        if List.Count(PreferredKinds) = 0 then
                            pageCandidatesRaw
                        else
                            List.Select(pageCandidatesRaw, each List.Contains(PreferredKinds, Text.Upper([TableKind]))),
                    sortedPageCandidates = SortCandidatesByPageAndIndex(pageCandidates),
                    exactPageCandidates = List.Select(sortedPageCandidates, each [ColumnCount] = ExpectedColumns),
                    bestExactCandidate = if List.Count(exactPageCandidates) = 0 then null else List.First(List.Sort(exactPageCandidates, (left, right) => if left[Score] < right[Score] then -1 else if left[Score] > right[Score] then 1 else 0)),
                    groupedVerticalCandidates = if bestExactCandidate = null then {} else List.Select(exactPageCandidates, each GetCandidateGroupKey(_) = GetCandidateGroupKey(bestExactCandidate)),
                    hasExactCandidatesOutsidePrimaryGroup = bestExactCandidate <> null and List.Count(groupedVerticalCandidates) <> List.Count(exactPageCandidates),
                    verticalSequenceValid = if List.Count(groupedVerticalCandidates) <= 1 then true else IsContinuousCandidateSequence(groupedVerticalCandidates),
                    partialPageCandidates = List.Select(sortedPageCandidates, each [ColumnCount] < ExpectedColumns),
                    horizontalMergeSequences = FindHorizontalMergeSequences(partialPageCandidates),
                    hasHorizontalAmbiguity = List.Count(horizontalMergeSequences) > 1,
                    horizontalMergeCandidates = if hasHorizontalAmbiguity or List.Count(horizontalMergeSequences) = 0 then {} else horizontalMergeSequences{0},
                    chosenMergedCandidate = if List.Count(mergedCandidates) = 0 then null else List.First(List.Sort(mergedCandidates, (left, right) => if left[Score] < right[Score] then -1 else if left[Score] > right[Score] then 1 else 0)),
                    verticalMergedData = if List.Count(groupedVerticalCandidates) <= 1 or verticalSequenceValid = false then null else Table.Combine(List.Transform(SortCandidatesByPageAndIndex(groupedVerticalCandidates), each BuildResultTable([Data], [PageNumber], 1, true))),
                    horizontalMergedData = if List.Count(horizontalMergeCandidates) <= 1 then null else CombineHorizontalTables(horizontalMergeCandidates),
                    chosenSinglePage =
                        if List.Count(exactPageCandidates) = 1 then
                            BuildResultTable(exactPageCandidates{0}[Data], exactPageCandidates{0}[PageNumber], 1, true)
                        else if List.Count(exactPageCandidates) = 0 and horizontalMergedData = null and chosenMergedCandidate <> null then
                            BuildResultTable(chosenMergedCandidate[Data], null, 1, true)
                        else
                            null,
                    failureRecord =
                        if mergedTry[HasError] then
                            [ErrorCode = "PDF_READ_FAILURE", ErrorCategory = "PDF読込エラー", UserMessage = "PDF を読み取れませんでした。壊れているか、Excel の PDF 解析で扱えない可能性があります。", TechnicalDetail = try Error.Message(mergedTry[Error]) otherwise "Pdf.Tables の読み取りに失敗しました。"]
                        else if hasExactCandidatesOutsidePrimaryGroup or (List.Count(groupedVerticalCandidates) > 1 and verticalSequenceValid = false) then
                            [ErrorCode = "TABLE_GROUP_AMBIGUOUS", ErrorCategory = "表分離エラー", UserMessage = "同一 PDF 内にヘッダー不一致または連続しない候補表があり、安全に結合できませんでした。", TechnicalDetail = "sameHeader の厳格判定で複数の full-table 候補が競合しました。"]
                        else if hasHorizontalAmbiguity then
                            [ErrorCode = "PARTIAL_TABLE_AMBIGUOUS", ErrorCategory = "表分離エラー", UserMessage = "同一 PDF 内に横分割候補が複数あり、安全に結合できませんでした。", TechnicalDetail = "partial-table の組み合わせが一意に定まりませんでした。"]
                        else if chosenMergedCandidate = null and horizontalMergedData = null and verticalMergedData = null and List.Count(exactPageCandidates) = 0 then
                            [ErrorCode = "TABLE_NOT_FOUND", ErrorCategory = "表検出エラー", UserMessage = "想定に近い表を検出できませんでした。", TechnicalDetail = "Pdf.Tables で抽出候補が見つからないか、ヘッダー行のみでした。"]
                        else
                            null,
                    resultData = if failureRecord <> null then null else if horizontalMergedData <> null then horizontalMergedData else if verticalMergedData <> null then verticalMergedData else chosenSinglePage,
                    withFileName = if resultData = null then null else Table.AddColumn(resultData, SourceFileColumnName, each fileName, type text),
                    reordered = if withFileName = null then null else Table.ReorderColumns(Table.RemoveColumns(withFileName, {"__PageNumber"}, MissingField.Ignore), OutputColumns, MissingField.UseNull),
                    reviewData =
                        BuildReviewTable(
                            resultData,
                            fileName,
                            if failureRecord = null then null else failureRecord[UserMessage],
                            if failureRecord = null then null else if List.Contains({"TABLE_GROUP_AMBIGUOUS", "PARTIAL_TABLE_AMBIGUOUS"}, failureRecord[ErrorCode]) then "HEADER_MISMATCH" else null
                        )
                in [
                    IsError = failureRecord <> null,
                    ErrorCode = if failureRecord = null then null else failureRecord[ErrorCode],
                    ErrorCategory = if failureRecord = null then null else failureRecord[ErrorCategory],
                    UserMessage = if failureRecord = null then null else failureRecord[UserMessage],
                    TechnicalDetail = if failureRecord = null then null else failureRecord[TechnicalDetail],
                    CandidateColumns = if bestExactCandidate <> null then bestExactCandidate[ColumnCount] else if chosenMergedCandidate <> null then chosenMergedCandidate[ColumnCount] else null,
                    CandidateRows = if bestExactCandidate <> null then bestExactCandidate[DataRowCount] else if chosenMergedCandidate <> null then chosenMergedCandidate[DataRowCount] else null,
                    OutputRowCount = if reordered = null then 0 else Table.RowCount(reordered),
                    ReviewCount = if reviewData = null then 0 else Table.RowCount(reviewData),
                    PageCount = if horizontalMergedData <> null then List.Count(horizontalMergeCandidates) else if verticalMergedData <> null then List.Count(groupedVerticalCandidates) else if chosenSinglePage <> null then 1 else 0,
                    SelectedTableId = if bestExactCandidate <> null then bestExactCandidate[TableId] else if chosenMergedCandidate <> null then chosenMergedCandidate[TableId] else null,
                    SelectedTableKind = if bestExactCandidate <> null then bestExactCandidate[TableKind] else if chosenMergedCandidate <> null then chosenMergedCandidate[TableKind] else null,
                    SelectedTableName = if bestExactCandidate <> null then bestExactCandidate[TableName] else if chosenMergedCandidate <> null then chosenMergedCandidate[TableName] else null,
                    Data = reordered,
                    Review = reviewData
                ]
        in if processTry[HasError] then [IsError = true, ErrorCode = "UNEXPECTED_PROCESSING_ERROR", ErrorCategory = "システムエラー", UserMessage = "この PDF の処理中に予期しないエラーが発生しました。", TechnicalDetail = try Error.Message(processTry[Error]) otherwise "unknown error", CandidateColumns = null, CandidateRows = null, OutputRowCount = 0, ReviewCount = 0, PageCount = 0, SelectedTableId = null, SelectedTableKind = null, SelectedTableName = null, Data = null, Review = null] else processTry[Value], type record),
    Expanded = Table.ExpandRecordColumn(WithProcessed, "Processed", {"IsError", "ErrorCode", "ErrorCategory", "UserMessage", "TechnicalDetail", "CandidateColumns", "CandidateRows", "OutputRowCount", "ReviewCount", "PageCount", "SelectedTableId", "SelectedTableKind", "SelectedTableName", "Data", "Review"}, {"IsError", "ErrorCode", "ErrorCategory", "UserMessage", "TechnicalDetail", "CandidateColumns", "CandidateRows", "OutputRowCount", "ReviewCount", "PageCount", "SelectedTableId", "SelectedTableKind", "SelectedTableName", "Data", "Review"})
in
    Expanded
"@
}

function Get-StagingQueryFormula {
    param(
        [Parameter(Mandatory = $true)][string]$InputPath,
        [Parameter(Mandatory = $true)]$Profile
    )

    if ($VersionMode -eq 'v2' -or $Profile.MultiPageMergeMode -eq 'sameHeader') {
        return Get-StagingQueryFormulaV2Simple -InputPath $InputPath -Profile $Profile
    }

    $escapedPath = Escape-MString -Value $InputPath
    $outputColumnsLiteral = ConvertTo-MTextListLiteral -Values (Get-ProfileOutputColumnNames -Profile $Profile)
    $dataColumnsLiteral = ConvertTo-MTextListLiteral -Values @((1..$Profile.ExpectedColumns | ForEach-Object { '{0}{1}' -f $Profile.DataColumnPrefix, $_ }))
    $preferredKindsLiteral = ConvertTo-MTextListLiteral -Values @($Profile.PreferredTableKinds | ForEach-Object { $_.ToUpperInvariant() })
    $preferredNamesLiteral = ConvertTo-MTextListLiteral -Values @($Profile.PreferredTableNameContains | ForEach-Object { $_.ToUpperInvariant() })
    $preferredIdsLiteral = ConvertTo-MTextListLiteral -Values @($Profile.PreferredTableIdContains | ForEach-Object { $_.ToUpperInvariant() })
    $sourceFileColumnNameLiteral = Escape-MString -Value $Profile.SourceFileColumnName
    $dataColumnPrefixLiteral = Escape-MString -Value $Profile.DataColumnPrefix
    $allowMoreColumnsLiteral = ConvertTo-MLogicalLiteral -Value $Profile.AllowMoreColumns

@"
let
    ExpectedColumns = $($Profile.ExpectedColumns),
    HeaderRowsToSkip = $($Profile.HeaderRowsToSkip),
    TargetRowCount = $($Profile.TargetRowCount),
    AllowMoreColumns = $allowMoreColumnsLiteral,
    SourceFileColumnName = "$sourceFileColumnNameLiteral",
    DataColumnPrefix = "$dataColumnPrefixLiteral",
    OutputColumns = $outputColumnsLiteral,
    DataColumns = $dataColumnsLiteral,
    PreferredKinds = $preferredKindsLiteral,
    PreferredNames = $preferredNamesLiteral,
    PreferredIds = $preferredIdsLiteral,
    Source = Folder.Files("$escapedPath"),
    PdfFiles = Table.SelectRows(Source, each Text.Lower([Extension]) = ".pdf"),
    KeepColumns = Table.SelectColumns(PdfFiles, {"Name", "Extension", "Folder Path", "Content"}),
    WithProcessed =
        Table.AddColumn(
            KeepColumns,
            "Processed",
            each
                let
                    fileName = [Name],
                    processingTry =
                        try
                            let
                                pdfTry = try Pdf.Tables([Content]),
                                pdfTables = if pdfTry[HasError] then null else pdfTry[Value],
                                pdfError = if pdfTry[HasError] then try Error.Message(pdfTry[Error]) otherwise "Pdf.Tables の読み取りに失敗しました。" else null,
                                candidateRecords = if pdfTables = null then {} else Table.ToRecords(pdfTables),
                                scoredCandidates =
                                    List.Transform(
                                        candidateRecords,
                                        each
                                            let
                                                dataTry = try Record.Field(_, "Data"),
                                                dataValue = if dataTry[HasError] then null else dataTry[Value],
                                                columnCount = if dataValue = null then null else Table.ColumnCount(dataValue),
                                                rowCount = if dataValue = null then null else Table.RowCount(dataValue),
                                                tableId = try Text.From(Record.Field(_, "Id")) otherwise "",
                                                tableKind = try Text.From(Record.Field(_, "Kind")) otherwise "",
                                                tableName = try Text.From(Record.Field(_, "Name")) otherwise "",
                                                normalizedKind = Text.Upper(tableKind),
                                                normalizedName = Text.Upper(tableName),
                                                normalizedId = Text.Upper(tableId),
                                                kindBonus = if List.Contains(PreferredKinds, normalizedKind) then -250 else 0,
                                                nameBonus =
                                                    if List.Count(PreferredNames) = 0 then
                                                        0
                                                    else if List.AnyTrue(List.Transform(PreferredNames, each Text.Contains(normalizedName, _))) then
                                                        -120
                                                    else
                                                        0,
                                                idBonus =
                                                    if List.Count(PreferredIds) = 0 then
                                                        0
                                                    else if List.AnyTrue(List.Transform(PreferredIds, each Text.Contains(normalizedId, _))) then
                                                        -120
                                                    else
                                                        0,
                                                score =
                                                    if dataValue = null or rowCount = null or columnCount = null then
                                                        999999
                                                    else
                                                        Number.Abs(columnCount - ExpectedColumns) * 1000 +
                                                        Number.Abs(rowCount - TargetRowCount) * 10 +
                                                        kindBonus + nameBonus + idBonus
                                            in
                                                [
                                                    Data = dataValue,
                                                    ColumnCount = columnCount,
                                                    RowCount = rowCount,
                                                    Score = score,
                                                    TableId = tableId,
                                                    TableKind = tableKind,
                                                    TableName = tableName
                                                ]
                                    ),
                                viableCandidates = List.Select(scoredCandidates, each [Data] <> null and [RowCount] <> null and [RowCount] > HeaderRowsToSkip),
                                sortedCandidates =
                                    List.Sort(
                                        viableCandidates,
                                        (left, right) =>
                                            if left[Score] < right[Score] then
                                                -1
                                            else if left[Score] > right[Score] then
                                                1
                                            else
                                                0
                                    ),
                                chosen = if List.Count(sortedCandidates) = 0 then null else List.First(sortedCandidates),
                                chosenColumns = if chosen = null then null else chosen[ColumnCount],
                                chosenRows = if chosen = null then null else chosen[RowCount],
                                chosenTableId = if chosen = null then null else chosen[TableId],
                                chosenTableKind = if chosen = null then null else chosen[TableKind],
                                chosenTableName = if chosen = null then null else chosen[TableName],
                                failureRecord =
                                    if pdfTry[HasError] then
                                        [
                                            ErrorCode = "PDF_READ_FAILURE",
                                            ErrorCategory = "PDF読込エラー",
                                            UserMessage = "PDF を読み取れませんでした。壊れているか、Excel の PDF 解析で扱えない可能性があります。",
                                            TechnicalDetail = pdfError
                                        ]
                                    else if chosen = null then
                                        [
                                            ErrorCode = "TABLE_NOT_FOUND",
                                            ErrorCategory = "表検出エラー",
                                            UserMessage = "想定に近い表を検出できませんでした。",
                                            TechnicalDetail = "Pdf.Tables で抽出候補が見つからないか、ヘッダー行のみでした。"
                                        ]
                                    else if AllowMoreColumns = false and chosenColumns <> null and chosenColumns > ExpectedColumns then
                                        [
                                            ErrorCode = "COLUMN_OVERFLOW",
                                            ErrorCategory = "列数不一致",
                                            UserMessage = "検出された表の列数がプロファイルの想定列数を超えています。",
                                            TechnicalDetail = "検出列数=" & Text.From(chosenColumns) & ", 想定列数=" & Text.From(ExpectedColumns)
                                        ]
                                    else
                                        null,
                                rawData =
                                    if failureRecord <> null or chosen = null then
                                        null
                                    else
                                        Table.Skip(chosen[Data], HeaderRowsToSkip),
                                originalColumns = if rawData = null then {} else Table.ColumnNames(rawData),
                                renamed =
                                    if rawData = null then
                                        null
                                    else
                                        Table.RenameColumns(
                                            rawData,
                                            List.Transform(List.Positions(originalColumns), each {originalColumns{_}, DataColumnPrefix & Text.From(_ + 1)}),
                                            MissingField.Ignore
                                        ),
                                renamedCount = if renamed = null then 0 else Table.ColumnCount(renamed),
                                missingColumns =
                                    if renamed = null or renamedCount >= ExpectedColumns then
                                        {}
                                    else
                                        List.Transform({renamedCount + 1 .. ExpectedColumns}, each DataColumnPrefix & Text.From(_)),
                                padded =
                                    if renamed = null then
                                        null
                                    else
                                        List.Accumulate(
                                            missingColumns,
                                            renamed,
                                            (state, columnName) => Table.AddColumn(state, columnName, each null, type text)
                                        ),
                                selected =
                                    if padded = null then
                                        null
                                    else
                                        Table.SelectColumns(padded, DataColumns, MissingField.UseNull),
                                textified =
                                    if selected = null then
                                        null
                                    else
                                        Table.TransformColumns(
                                            selected,
                                            List.Transform(
                                                Table.ColumnNames(selected),
                                                each {_, (value) => if value = null then "" else Text.From(value), type text}
                                            )
                                        ),
                                withFileName =
                                    if textified = null then
                                        null
                                    else
                                        Table.AddColumn(textified, SourceFileColumnName, each fileName, type text),
                                reordered =
                                    if withFileName = null then
                                        null
                                    else
                                        Table.ReorderColumns(withFileName, OutputColumns, MissingField.UseNull),
                                outputRowCount = if reordered = null then 0 else Table.RowCount(reordered)
                            in
                                [
                                    IsError = failureRecord <> null,
                                    ErrorCode = if failureRecord = null then null else failureRecord[ErrorCode],
                                    ErrorCategory = if failureRecord = null then null else failureRecord[ErrorCategory],
                                    UserMessage = if failureRecord = null then null else failureRecord[UserMessage],
                                    TechnicalDetail = if failureRecord = null then null else failureRecord[TechnicalDetail],
                                    CandidateColumns = chosenColumns,
                                    CandidateRows = chosenRows,
                                    OutputRowCount = outputRowCount,
                                    SelectedTableId = chosenTableId,
                                    SelectedTableKind = chosenTableKind,
                                    SelectedTableName = chosenTableName,
                                    Data = reordered
                                ]
                in
                    if processingTry[HasError] then
                        [
                            IsError = true,
                            ErrorCode = "UNEXPECTED_PROCESSING_ERROR",
                            ErrorCategory = "システムエラー",
                            UserMessage = "この PDF の処理中に予期しないエラーが発生しました。",
                            TechnicalDetail = try Error.Message(processingTry[Error]) otherwise "unknown error",
                            CandidateColumns = null,
                            CandidateRows = null,
                            OutputRowCount = 0,
                            SelectedTableId = null,
                            SelectedTableKind = null,
                            SelectedTableName = null,
                            Data = null
                        ]
                    else
                        processingTry[Value],
            type record
        ),
    Expanded =
        Table.ExpandRecordColumn(
            WithProcessed,
            "Processed",
            {"IsError", "ErrorCode", "ErrorCategory", "UserMessage", "TechnicalDetail", "CandidateColumns", "CandidateRows", "OutputRowCount", "SelectedTableId", "SelectedTableKind", "SelectedTableName", "Data"},
            {"IsError", "ErrorCode", "ErrorCategory", "UserMessage", "TechnicalDetail", "CandidateColumns", "CandidateRows", "OutputRowCount", "SelectedTableId", "SelectedTableKind", "SelectedTableName", "Data"}
        )
in
    Expanded
"@
}

function Get-ResultQueryFormula {
    param(
        [Parameter(Mandatory = $true)][string]$InputPath,
        [Parameter(Mandatory = $true)]$Profile
    )

    $outputColumnsLiteral = ConvertTo-MTextListLiteral -Values (Get-ProfileOutputColumnNames -Profile $Profile)

@"
let
    OutputColumns = $outputColumnsLiteral,
    Source = PDF2Excel_Staging,
    SuccessRows = Table.SelectRows(Source, each [IsError] <> true and [Data] <> null),
    Expanded = if Table.RowCount(SuccessRows) = 0 then #table(OutputColumns, {}) else Table.ExpandTableColumn(SuccessRows, "Data", OutputColumns, OutputColumns),
    Reordered = Table.SelectColumns(Expanded, OutputColumns, MissingField.UseNull)
in
    Reordered
"@
}

function Get-ErrorsQueryFormula {
@"
let
    Source = PDF2Excel_Staging,
    ErrorRows = Table.SelectRows(Source, each [IsError] = true),
    Selected =
        Table.SelectColumns(
            ErrorRows,
            {"Name", "ErrorCode", "ErrorCategory", "UserMessage", "TechnicalDetail", "CandidateColumns", "CandidateRows", "SelectedTableId", "SelectedTableKind", "SelectedTableName"},
            MissingField.UseNull
        ),
    Renamed =
        Table.RenameColumns(
            Selected,
            {
                {"Name", "FileName"},
                {"SelectedTableId", "TableId"},
                {"SelectedTableKind", "TableKind"},
                {"SelectedTableName", "TableName"}
            }
        )
in
    Renamed
"@
}

function Get-FileSummaryQueryFormula {
    if ($VersionMode -eq 'v2') {
        return Get-FileSummaryQueryFormulaV2
    }

@"
let
    Source = PDF2Excel_Staging,
    Selected =
        Table.SelectColumns(
            Source,
            {"Name", "IsError", "OutputRowCount", "ErrorCategory", "ErrorCode", "UserMessage", "CandidateColumns", "CandidateRows", "SelectedTableKind", "SelectedTableName"},
            MissingField.UseNull
        ),
    WithStatus = Table.AddColumn(Selected, "Status", each if [IsError] = true then "失敗" else "成功", type text),
    Reordered =
        Table.ReorderColumns(
            WithStatus,
            {"Name", "Status", "OutputRowCount", "ErrorCategory", "ErrorCode", "UserMessage", "CandidateColumns", "CandidateRows", "SelectedTableKind", "SelectedTableName"},
            MissingField.UseNull
        ),
    Renamed =
        Table.RenameColumns(
            Reordered,
            {
                {"Name", "FileName"},
                {"OutputRowCount", "ImportedRows"},
                {"SelectedTableKind", "TableKind"},
                {"SelectedTableName", "TableName"}
            }
        )
in
    Renamed
"@
}

function Get-ReviewQueryFormula {
    param([Parameter(Mandatory = $true)]$Profile)

    if ($VersionMode -eq 'v2') {
        return Get-ReviewQueryFormulaV2 -Profile $Profile
    }

@"
let
    Source = #table({"FileName", "Page", "PersonRaw", "SiteRaw", "InRaw", "OutRaw", "Reason"}, {})
in
    Source
"@
}

function Get-ErrorSummaryQueryFormula {
@"
let
    Source = PDF2Excel_Staging,
    ErrorRows = Table.SelectRows(Source, each [IsError] = true),
    Grouped =
        Table.Group(
            ErrorRows,
            {"ErrorCategory", "ErrorCode"},
            {{"Count", each Table.RowCount(_), Int64.Type}}
        ),
    Sorted = Table.Sort(Grouped, {{"Count", Order.Descending}, {"ErrorCategory", Order.Ascending}, {"ErrorCode", Order.Ascending}})
in
    Sorted
"@
}

# ============================================================
# Section: Excel Workbook Operations
# ============================================================

function Remove-WorkbookQuery {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)][string]$QueryName
    )

    try {
        $query = $Workbook.Queries.Item($QueryName)
        $query.Delete()
        $query | Release-ComObject
    } catch {
    }
}

function Add-OrReplaceWorkbookQuery {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)][string]$QueryName,
        [Parameter(Mandatory = $true)][string]$Formula
    )

    Remove-WorkbookQuery -Workbook $Workbook -QueryName $QueryName
    $null = $Workbook.Queries.Add($QueryName, $Formula)
}

function Get-OrCreateWorksheet {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)][string]$WorksheetName
    )

    try {
        return $Workbook.Worksheets.Item($WorksheetName)
    } catch {
        $worksheet = $Workbook.Worksheets.Add()
        $worksheet.Name = $WorksheetName
        return $worksheet
    }
}

function Load-WorkbookQueryToWorksheet {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)][string]$WorksheetName,
        [Parameter(Mandatory = $true)][string]$QueryName,
        [Parameter(Mandatory = $true)][string]$TableName,
        [string]$DestinationAddress = 'A1',
        [switch]$ClearSheet = $true
    )

    $worksheet = $null
    $listObject = $null
    $queryTable = $null

    try {
        $worksheet = Get-OrCreateWorksheet -Workbook $Workbook -WorksheetName $WorksheetName
        if ($ClearSheet) {
            while ($worksheet.ListObjects.Count -gt 0) {
                $listObject = $worksheet.ListObjects.Item(1)
                $listObject.Delete()
                $listObject | Release-ComObject
                $listObject = $null
            }
            $worksheet.Cells.Clear() | Out-Null
        } else {
            for ($index = $worksheet.ListObjects.Count; $index -ge 1; $index -= 1) {
                $existingListObject = $null
                try {
                    $existingListObject = $worksheet.ListObjects.Item($index)
                    if ([string]$existingListObject.Name -eq $TableName) {
                        $existingListObject.Delete()
                    }
                } finally {
                    $existingListObject | Release-ComObject
                }
            }
        }

        $source = @('OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=' + $QueryName + ';Extended Properties=""')
        try {
            $listObject = $worksheet.ListObjects.Add(0, $source, $null, 1, $worksheet.Range($DestinationAddress))
        } catch {
            throw "ワークシートへのテーブル追加に失敗しました: $($_.Exception.Message)"
        }

        try {
            $listObject.Name = $TableName
            $queryTable = $listObject.QueryTable
        } catch {
            throw "QueryTable の取得に失敗しました: $($_.Exception.Message)"
        }

        try {
            $queryTable.CommandType = 2
            $queryTable.CommandText = @("SELECT * FROM [$QueryName]")
            $queryTable.BackgroundQuery = $false
        } catch {
            throw "QueryTable の設定に失敗しました: $($_.Exception.Message)"
        }

        try {
            $queryTable.Refresh($false) | Out-Null
            $worksheet.Columns.AutoFit() | Out-Null
        } catch {
            throw "QueryTable の更新に失敗しました: $($_.Exception.Message)"
        }
    } catch {
        throw "シート ${WorksheetName} へのクエリ $QueryName 読み込みに失敗しました: $($_.Exception.Message)"
    } finally {
        $queryTable | Release-ComObject
        $listObject | Release-ComObject
        $worksheet | Release-ComObject
    }
}

function Set-ControlValues {
    param(
        [Parameter(Mandatory = $true)]$Worksheet,
        [Parameter(Mandatory = $true)][string]$StagingInputFolder,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)][string]$LogPath,
        [Parameter(Mandatory = $true)]$Profile,
        [ValidateSet('v1', 'v2')]
        [string]$VersionMode = 'v1'
    )

    Set-ControlSheetStaticCells -Worksheet $Worksheet -VersionMode $VersionMode

    $Worksheet.Range('B2').Value2 = $StagingInputFolder
    $Worksheet.Range('B3').Value2 = $OutputPath
    $Worksheet.Range('B4').Value2 = $LogPath
    $Worksheet.Range('B5').Value2 = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    $Worksheet.Range('B6').Value2 = '準備完了'
    $Worksheet.Range('B7:B12').Value2 = ''
    $Worksheet.Range('B13').Value2 = $Profile.DisplayName
    $Worksheet.Range('B14').Value2 = $Profile.Description
    $Worksheet.Range('B15').Value2 = Get-VersionDisplayName -VersionMode $VersionMode
}

function Clear-ControlPathsForSecureOutput {
    param(
        [Parameter(Mandatory = $true)]$Worksheet,
        [ValidateSet('v1', 'v2')]
        [string]$VersionMode = 'v1'
    )

    if ($VersionMode -ne 'v2' -or -not $script:isSecureMode) {
        return
    }

    $Worksheet.Range('B2:B4').Value2 = ''
}

function Set-ControlMetrics {
    param(
        [Parameter(Mandatory = $true)]$Worksheet,
        [int]$SourcePdfCount,
        [int]$ResultRowCount,
        [int]$ErrorRowCount,
        [int]$SuccessPdfCount,
        [int]$FailedPdfCount,
        [int]$ElapsedSeconds
    )

    $Worksheet.Range('B7').Value2 = $SourcePdfCount
    $Worksheet.Range('B8').Value2 = $ResultRowCount
    $Worksheet.Range('B9').Value2 = $ErrorRowCount
    $Worksheet.Range('B10').Value2 = $SuccessPdfCount
    $Worksheet.Range('B11').Value2 = $FailedPdfCount
    $Worksheet.Range('B12').Value2 = $ElapsedSeconds
}

function Initialize-SummarySheet {
    param(
        [Parameter(Mandatory = $true)]$Worksheet,
        [ValidateSet('v1', 'v2')]
        [string]$VersionMode = 'v1'
    )

    while ($Worksheet.ListObjects.Count -gt 0) {
        $listObject = $null
        try {
            $listObject = $Worksheet.ListObjects.Item(1)
            $listObject.Delete()
        } finally {
            $listObject | Release-ComObject
        }
    }

    $Worksheet.Cells.Clear() | Out-Null
    Set-SummarySheetStaticCells -Worksheet $Worksheet -VersionMode $VersionMode
    $Worksheet.Range('A1').Font.Bold = $true
    $Worksheet.Range('A1').Font.Size = 14
    $Worksheet.Range('A4:B4').Font.Bold = $true
    $Worksheet.Range('A4:B4').Interior.Color = 15773696
    $Worksheet.Range('A13').Font.Bold = $true
    if ($VersionMode -eq 'v2') {
        $Worksheet.Range('S13').Value2 = 'エラー分類別件数'
        $Worksheet.Range('S13').Font.Bold = $true
    } else {
        $Worksheet.Range('M13').Font.Bold = $true
    }
    $Worksheet.Columns.Item('A').ColumnWidth = 24
    $Worksheet.Columns.Item('B').ColumnWidth = 28
    $Worksheet.Columns.Item('C').ColumnWidth = 22
    $Worksheet.Columns.Item('D').ColumnWidth = 18
    $Worksheet.Columns.Item('E').ColumnWidth = 18
    $Worksheet.Columns.Item('F').ColumnWidth = 20
    $Worksheet.Columns.Item('G').ColumnWidth = 16
    $Worksheet.Columns.Item('H').ColumnWidth = 16
    $Worksheet.Columns.Item('I').ColumnWidth = 18
    $Worksheet.Columns.Item('J').ColumnWidth = 18
    $Worksheet.Columns.Item('K').ColumnWidth = 18
    $Worksheet.Columns.Item('L').ColumnWidth = 16
    $Worksheet.Columns.Item('M').ColumnWidth = 18
    $Worksheet.Columns.Item('N').ColumnWidth = 18
    $Worksheet.Columns.Item('O').ColumnWidth = 16
}

function Set-SummaryMetrics {
    param(
        [Parameter(Mandatory = $true)]$Worksheet,
        [int]$SourcePdfCount,
        [int]$SuccessPdfCount,
        [int]$FailedPdfCount,
        [int]$ResultRowCount,
        [int]$ErrorRowCount,
        [int]$ReviewRowCount,
        [int]$ElapsedSeconds,
        [Parameter(Mandatory = $true)]$Profile,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [ValidateSet('v1', 'v2')]
        [string]$VersionMode = 'v1'
    )

    $Worksheet.Range('A5').Value2 = '対象PDF数'
    $Worksheet.Range('B5').Value2 = $SourcePdfCount
    $Worksheet.Range('A6').Value2 = '成功PDF数'
    $Worksheet.Range('B6').Value2 = $SuccessPdfCount
    $Worksheet.Range('A7').Value2 = '失敗PDF数'
    $Worksheet.Range('B7').Value2 = $FailedPdfCount
    $Worksheet.Range('A8').Value2 = '取込データ行数'
    $Worksheet.Range('B8').Value2 = $ResultRowCount
    $Worksheet.Range('A9').Value2 = 'エラー件数'
    $Worksheet.Range('B9').Value2 = $ErrorRowCount
    if ($VersionMode -eq 'v2') {
        $Worksheet.Range('A10').Value2 = 'Review件数'
        $Worksheet.Range('B10').Value2 = $ReviewRowCount
        $Worksheet.Range('A11').Value2 = '処理時間(秒)'
        $Worksheet.Range('B11').Value2 = $ElapsedSeconds
        $Worksheet.Range('A12').Value2 = '使用プロファイル'
        $Worksheet.Range('B12').Value2 = $Profile.DisplayName
        $Worksheet.Range('D5').Value2 = '出力ファイル'
        $Worksheet.Range('E5').Value2 = $OutputPath
    } else {
        $Worksheet.Range('A10').Value2 = '処理時間(秒)'
        $Worksheet.Range('B10').Value2 = $ElapsedSeconds
        $Worksheet.Range('A11').Value2 = '使用プロファイル'
        $Worksheet.Range('B11').Value2 = $Profile.DisplayName
        $Worksheet.Range('A12').Value2 = '出力ファイル'
        $Worksheet.Range('B12').Value2 = $OutputPath
    }
}

function Get-DataRowCount {
    param(
        [Parameter(Mandatory = $true)]$Worksheet,
        [string]$ExpectedTableName
    )

    $listObject = $null
    try {
        if ($Worksheet.ListObjects.Count -gt 0) {
            if ([string]::IsNullOrWhiteSpace($ExpectedTableName)) {
                $listObject = $Worksheet.ListObjects.Item(1)
            } else {
                $listObject = $Worksheet.ListObjects.Item($ExpectedTableName)
            }

            if ($null -eq $listObject.DataBodyRange) {
                return 0
            }

            return [int]$listObject.DataBodyRange.Rows.Count
        }

        $usedRows = [int]$Worksheet.UsedRange.Rows.Count
        return [Math]::Max($usedRows - 1, 0)
    } finally {
        $listObject | Release-ComObject
    }
}

function Try-RunMacro {
    param(
        [Parameter(Mandatory = $true)]$Excel,
        [Parameter(Mandatory = $true)][string]$WorkbookPath,
        [Parameter(Mandatory = $true)][string]$MacroName
    )

    $workbookName = [System.IO.Path]::GetFileName($WorkbookPath)
    $macroTarget = "'$workbookName'!$MacroName"
    $Excel.Run($macroTarget) | Out-Null
}

function Export-WorkbookDirectly {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)][string]$OutputPath
    )

    if (Test-Path -LiteralPath $OutputPath) {
        Remove-Item -LiteralPath $OutputPath -Force
    }

    Clear-ControlPathsForSecureOutput -Worksheet $Workbook.Worksheets.Item('Control') -VersionMode $VersionMode
    $Workbook.SaveAs($OutputPath, 51)
}

function Get-PreflightState {
    param(
        [Parameter(Mandatory = $true)][string[]]$SourceFiles,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)]$Profile
    )

    return [pscustomobject]@{
        SourceFileCount = $SourceFiles.Count
        OutputPath = $OutputPath
        OutputExists = Test-Path -LiteralPath $OutputPath
        Profile = $Profile
        SampleFiles = @($SourceFiles | Select-Object -First 5 | ForEach-Object { [System.IO.Path]::GetFileName($_) })
    }
}

function Show-PreflightCheck {
    param([Parameter(Mandatory = $true)]$PreflightState)

    Write-Host ''
    Write-Host '実行前チェック' -ForegroundColor Yellow
    Write-Host ('  対象PDF数       : {0}' -f $PreflightState.SourceFileCount)
    Write-Host ('  出力ファイル     : {0}' -f $PreflightState.OutputPath)
    Write-Host ('  既存ファイル上書き: {0}' -f $(if ($PreflightState.OutputExists) { 'あり' } else { 'なし' }))
    Write-Host ('  使用プロファイル : {0}' -f $PreflightState.Profile.DisplayName)
    Write-Host ('  想定列数        : {0}' -f $PreflightState.Profile.ExpectedColumns)
    Write-Host ('  ヘッダー除外行数 : {0}' -f $PreflightState.Profile.HeaderRowsToSkip)
    Write-Host ('  想定行数        : {0}' -f $PreflightState.Profile.TargetRowCount)
    Write-Host ('  プロファイル説明 : {0}' -f $PreflightState.Profile.Description)
    Write-Host '  代表ファイル     :'
    foreach ($fileName in $PreflightState.SampleFiles) {
        Write-Host ('    - {0}' -f $fileName)
    }
    if ($PreflightState.SourceFileCount -gt $PreflightState.SampleFiles.Count) {
        Write-Host ('    ... 他 {0} 件' -f ($PreflightState.SourceFileCount - $PreflightState.SampleFiles.Count))
    }
    Write-Host ''
}

function Confirm-Preflight {
    param([Parameter(Mandatory = $true)]$PreflightState)

    Show-PreflightCheck -PreflightState $PreflightState
    if ($NoConfirm) {
        Write-Log 'NoConfirm が指定されたため、実行前チェック後にそのまま続行します。'
        return
    }

    $confirmation = Read-Host 'この内容で実行しますか? [Y/N]'
    if ($confirmation.Trim().ToUpperInvariant() -notin @('Y', 'YES')) {
        throw '実行前チェックでキャンセルされました。'
    }
}

function Get-EffectiveInputFileCandidates {
    $inputFileCandidates = @($InputFiles)
    if ($inputFileCandidates.Count -gt 0) {
        $inputFileCandidates += @($args)
    } elseif ($args.Count -gt 0) {
        throw "不明な引数があります: $($args -join ', ')"
    }

    return @($inputFileCandidates)
}

function Resolve-ExecutionPlan {
    $inputFileCandidates = @(Get-EffectiveInputFileCandidates)
    $sourceFiles = @(Resolve-InputPdfFiles -SourceFolder $InputFolder -SourceFiles $inputFileCandidates)
    Write-Log ("対象 PDF 数: {0}" -f $sourceFiles.Count)

    $profile = Get-ProfileConfiguration -RequestedProfileName $ProfileName -RequestedProfilePath $ProfilePath
    Register-SensitivePaths -Paths @($profile.ProfilePath)
    Write-Log ("使用プロファイル: {0} ({1})" -f $profile.DisplayName, $profile.ProfilePath)

    $defaultOutputPath = Join-Path $outputDir "PDF2Excel_$timestamp.xlsx"
    $effectiveOutputPath = $OutputFile
    if ([string]::IsNullOrWhiteSpace($effectiveOutputPath)) {
        $effectiveOutputPath = if ($PromptForOutputFile) { Select-OutputFileDialog -DefaultOutputPath $defaultOutputPath } else { $defaultOutputPath }
    }

    if (-not [string]::IsNullOrWhiteSpace($effectiveOutputPath)) {
        $effectiveOutputPath = [System.IO.Path]::GetFullPath($effectiveOutputPath)
    }

    $outputParent = Split-Path -Path $effectiveOutputPath -Parent
    Ensure-Directory -Path $outputParent
    Register-SensitivePaths -Paths @($sourceFiles + @($effectiveOutputPath, $outputParent))

    $preflightState = Get-PreflightState -SourceFiles $sourceFiles -OutputPath $effectiveOutputPath -Profile $profile
    Confirm-Preflight -PreflightState $preflightState

    return [pscustomobject]@{
        SourceFiles    = $sourceFiles
        Profile        = $profile
        OutputPath     = $effectiveOutputPath
        OutputParent   = $outputParent
        PreflightState = $preflightState
    }
}

function Initialize-RunWorkspace {
    Ensure-Workspace
    Register-SensitivePaths -Paths @(
        $baseDir,
        $scriptDir,
        $templateDir,
        $configDir,
        $profilesDir,
        $inputDir,
        $outputDir,
        $runtimeRootDir,
        $runtimeRunsDir,
        $logsDir,
        $script:runWorkspaceDir,
        $script:runStagingDir,
        $script:runRuntimeDir,
        $script:lockFilePath,
        $script:logPath,
        $templatePath,
        $buildTemplateScript
    )
    $rotation = Rotate-LogFiles
    Set-Content -LiteralPath $script:logPath -Value "PDF2Excel run started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Encoding UTF8
    if ($rotation.RemovedByAge -gt 0 -or $rotation.RemovedByCount -gt 0) {
        Write-Log ("ログ整理を実行しました: 期限切れ削除={0}, 件数調整削除={1}" -f $rotation.RemovedByAge, $rotation.RemovedByCount)
    }
    Acquire-RunLock
    Compact-RuntimeArtifacts
    Ensure-Directory -Path $script:runWorkspaceDir
    Ensure-Directory -Path $script:runStagingDir
    Ensure-Directory -Path $script:runRuntimeDir
}

function Prepare-RunInputs {
    param(
        [Parameter(Mandatory = $true)][string[]]$SourceFiles
    )

    $stagedFiles = Stage-PdfFiles -Files $SourceFiles -StagingDirectory $script:runStagingDir
    $storedFiles = @()
    if ($script:isSecureMode) {
        if ($KeepInput) {
            Write-Log 'VER2 Secure では -KeepInput は無効です。input フォルダ同期を行いません。' 'WARN'
        } else {
            Write-Log 'VER2 Secure のため input フォルダ同期を行いません。'
        }
    } else {
        $storedFiles = Sync-InputStorage -Files $SourceFiles -KeepExisting:$KeepInput
    }
    Register-SensitivePaths -Paths @($stagedFiles + $storedFiles)
    Write-Log ("今回実行分の staging が完了しました: {0}" -f ($stagedFiles -join ', '))
    if ($script:isSecureMode) {
        Write-Log 'VER2 Secure のため input フォルダへの PDF 複製は作成していません。'
    } else {
        Write-Log ("input フォルダ同期が完了しました: {0}" -f ($storedFiles -join ', '))
    }
    # Give Excel's PDF connector a brief moment to observe freshly staged files on disk.
    Start-Sleep -Seconds 1

    return [pscustomobject]@{
        StagedFiles = $stagedFiles
        StoredFiles = $storedFiles
    }
}

function Open-ExcelRuntimeContext {
    param(
        [Parameter(Mandatory = $true)][string]$RuntimeWorkbookPath,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)]$Profile
    )

    Write-Log 'Excel を起動しています。'
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 1
    $excelProcessId = Get-ExcelProcessId -ExcelApplication $excel

    $workbook = $excel.Workbooks.Open($RuntimeWorkbookPath)
    $controlSheet = Get-OrCreateWorksheet -Workbook $workbook -WorksheetName 'Control'
    $summarySheet = Get-OrCreateWorksheet -Workbook $workbook -WorksheetName 'Summary'
    $resultSheetRef = Get-OrCreateWorksheet -Workbook $workbook -WorksheetName 'Result'
    $errorsSheetRef = Get-OrCreateWorksheet -Workbook $workbook -WorksheetName 'Errors'
    if ($VersionMode -eq 'v2') {
        $reviewSheetRef = Get-OrCreateWorksheet -Workbook $workbook -WorksheetName 'Review'
        $reviewSheetRef | Release-ComObject
    }
    $resultSheetRef | Release-ComObject
    $errorsSheetRef | Release-ComObject

    Set-ControlValues -Worksheet $controlSheet -StagingInputFolder $script:runStagingDir -OutputPath $OutputPath -LogPath $script:logPath -Profile $Profile -VersionMode $VersionMode
    Initialize-SummarySheet -Worksheet $summarySheet -VersionMode $VersionMode

    return [pscustomobject]@{
        Excel          = $excel
        ExcelProcessId = $excelProcessId
        Workbook       = $workbook
        ControlSheet   = $controlSheet
        SummarySheet   = $summarySheet
    }
}

function Configure-WorkbookQueries {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)]$Profile
    )

    Write-Log 'Power Query を設定しています。'
    Add-OrReplaceWorkbookQuery -Workbook $Workbook -QueryName 'PDF2Excel_Staging' -Formula (Get-StagingQueryFormula -InputPath $script:runStagingDir -Profile $Profile)
    Add-OrReplaceWorkbookQuery -Workbook $Workbook -QueryName 'PDF2Excel_Result' -Formula (Get-ResultQueryFormula -InputPath $script:runStagingDir -Profile $Profile)
    Add-OrReplaceWorkbookQuery -Workbook $Workbook -QueryName 'PDF2Excel_Errors' -Formula (Get-ErrorsQueryFormula)
    Add-OrReplaceWorkbookQuery -Workbook $Workbook -QueryName 'PDF2Excel_FileSummary' -Formula (Get-FileSummaryQueryFormula)
    Add-OrReplaceWorkbookQuery -Workbook $Workbook -QueryName 'PDF2Excel_ErrorSummary' -Formula (Get-ErrorSummaryQueryFormula)
    if ($VersionMode -eq 'v2') {
        Add-OrReplaceWorkbookQuery -Workbook $Workbook -QueryName 'PDF2Excel_Review' -Formula (Get-ReviewQueryFormula -Profile $Profile)
    } else {
        Remove-WorkbookQuery -Workbook $Workbook -QueryName 'PDF2Excel_Review'
    }
}

function Load-WorkbookOutputSheets {
    param(
        [Parameter(Mandatory = $true)]$Workbook
    )

    Write-Log $(if ($VersionMode -eq 'v2') { 'Result / Review / Errors / Summary シートへ読み込んでいます。' } else { 'Result / Errors / Summary シートへ読み込んでいます。' })
    Load-WorkbookQueryToWorksheet -Workbook $Workbook -WorksheetName 'Result' -QueryName 'PDF2Excel_Result' -TableName 'tblResult'
    Load-WorkbookQueryToWorksheet -Workbook $Workbook -WorksheetName 'Errors' -QueryName 'PDF2Excel_Errors' -TableName 'tblErrors'
    if ($VersionMode -eq 'v2') {
        Load-WorkbookQueryToWorksheet -Workbook $Workbook -WorksheetName 'Review' -QueryName 'PDF2Excel_Review' -TableName 'tblReview'
    }
    Load-WorkbookQueryToWorksheet -Workbook $Workbook -WorksheetName 'Summary' -QueryName 'PDF2Excel_FileSummary' -TableName 'tblFileSummary' -DestinationAddress 'A14' -ClearSheet:$false
    Load-WorkbookQueryToWorksheet -Workbook $Workbook -WorksheetName 'Summary' -QueryName 'PDF2Excel_ErrorSummary' -TableName 'tblErrorSummary' -DestinationAddress $(if ($VersionMode -eq 'v2') { 'S14' } else { 'M14' }) -ClearSheet:$false
}

function Add-V2ResultNormalizedColumns {
    param(
        [Parameter(Mandatory = $true)]$Worksheet,
        [Parameter(Mandatory = $true)]$Profile
    )

    $definitions = @(Get-NormalizedTimeColumnDefinitions -Profile $Profile -VersionMode 'v2')
    if ($definitions.Count -eq 0) {
        return
    }

    $usedRange = $Worksheet.UsedRange
    $rowCount = [int]$usedRange.Rows.Count
    $columnCount = [int]$usedRange.Columns.Count
    $headerMap = @{}

    for ($column = 1; $column -le $columnCount; $column += 1) {
        $header = [string]$Worksheet.Cells.Item(1, $column).Value2
        if (-not [string]::IsNullOrWhiteSpace($header)) {
            $headerMap[$header] = $column
        }
    }

    foreach ($definition in $definitions) {
        foreach ($headerName in @($definition.DisplayName, $definition.MinutesColumnName)) {
            if (-not $headerMap.ContainsKey($headerName)) {
                $columnCount += 1
                $Worksheet.Cells.Item(1, $columnCount).Value2 = $headerName
                $headerMap[$headerName] = $columnCount
            }
        }
    }

    foreach ($headerName in @('ReasonCategory', '時刻正規化状態', '時刻確認メモ')) {
        if (-not $headerMap.ContainsKey($headerName)) {
            $columnCount += 1
            $Worksheet.Cells.Item(1, $columnCount).Value2 = $headerName
            $headerMap[$headerName] = $columnCount
        }
    }

    foreach ($definition in $definitions) {
        $Worksheet.Cells.Item(1, $headerMap[$definition.DisplayName]).EntireColumn.NumberFormat = '@'
        $Worksheet.Cells.Item(1, $headerMap[$definition.MinutesColumnName]).EntireColumn.NumberFormat = '0'
    }

    for ($row = 2; $row -le $rowCount; $row += 1) {
        $rawValuesByDisplayName = @{}
        foreach ($definition in $definitions) {
            $rawValuesByDisplayName[$definition.DisplayName] =
                if ($headerMap.ContainsKey($definition.SourceColumnName)) {
                    [string]$Worksheet.Cells.Item($row, $headerMap[$definition.SourceColumnName]).Text
                } else {
                    ''
                }
        }

        $audit = Get-TimeNormalizationAudit -Definitions $definitions -RawValuesByDisplayName $rawValuesByDisplayName

        foreach ($definition in $definitions) {
            try {
                $normalized = $audit.Results[$definition.DisplayName]
                $displayCell = $Worksheet.Cells.Item($row, $headerMap[$definition.DisplayName])
                $minutesCell = $Worksheet.Cells.Item($row, $headerMap[$definition.MinutesColumnName])

                $displayCell.Value2 = $normalized.NormalizedText
                if ($null -eq $normalized.MinutesFromMidnight) {
                    $minutesCell.ClearContents()
                } else {
                    $minutesCell.Value2 = [double]$normalized.MinutesFromMidnight
                }
            } catch {
                throw "Result 正規化列 '$($definition.DisplayName)' の書き込みに失敗しました (row=$row): $($_.Exception.Message)"
            }
        }

        $Worksheet.Cells.Item($row, $headerMap['時刻正規化状態']).Value2 = $audit.Status
        $Worksheet.Cells.Item($row, $headerMap['時刻確認メモ']).Value2 = $audit.Note
    }

    foreach ($definition in $definitions) {
        try {
            $Worksheet.Cells.Item(1, $headerMap[$definition.DisplayName]).EntireColumn.NumberFormat = '@'
            $Worksheet.Cells.Item(1, $headerMap[$definition.MinutesColumnName]).EntireColumn.NumberFormat = '0'
        } catch {
            throw "Result 正規化列 '$($definition.DisplayName)' の表示形式設定に失敗しました: $($_.Exception.Message)"
        }
    }
    $Worksheet.Range('A1').EntireRow.Font.Bold = $true
    $Worksheet.Columns.AutoFit() | Out-Null
}

function Add-V2ReviewNormalizedColumns {
    param(
        [Parameter(Mandatory = $true)]$Worksheet,
        [Parameter(Mandatory = $true)]$Profile
    )

    $usedRange = $Worksheet.UsedRange
    $rowCount = [int]$usedRange.Rows.Count
    $columnCount = [int]$usedRange.Columns.Count
    $headerMap = @{}
    $definitions = @(Get-NormalizedTimeColumnDefinitions -Profile $Profile -VersionMode 'v2')
    if ($definitions.Count -eq 0) {
        return
    }

    for ($column = 1; $column -le $columnCount; $column += 1) {
        $header = [string]$Worksheet.Cells.Item(1, $column).Value2
        if (-not [string]::IsNullOrWhiteSpace($header)) {
            $headerMap[$header] = $column
        }
    }

    foreach ($definition in $definitions) {
        foreach ($headerName in @($definition.DisplayName, $definition.MinutesColumnName)) {
            if (-not $headerMap.ContainsKey($headerName)) {
                $columnCount += 1
                $Worksheet.Cells.Item(1, $columnCount).Value2 = $headerName
                $headerMap[$headerName] = $columnCount
            }
        }
    }

    foreach ($headerName in @('時刻正規化状態', '時刻確認メモ')) {
        if (-not $headerMap.ContainsKey($headerName)) {
            $columnCount += 1
            $Worksheet.Cells.Item(1, $columnCount).Value2 = $headerName
            $headerMap[$headerName] = $columnCount
        }
    }

    foreach ($definition in $definitions) {
        $Worksheet.Cells.Item(1, $headerMap[$definition.DisplayName]).EntireColumn.NumberFormat = '@'
        $Worksheet.Cells.Item(1, $headerMap[$definition.MinutesColumnName]).EntireColumn.NumberFormat = '0'
    }

    for ($row = 2; $row -le $rowCount; $row += 1) {
        $rawValuesByDisplayName = @{}
        foreach ($definition in $definitions) {
            $rawValuesByDisplayName[$definition.DisplayName] =
                if ($headerMap.ContainsKey($definition.ReviewRawColumnName)) {
                    [string]$Worksheet.Cells.Item($row, $headerMap[$definition.ReviewRawColumnName]).Text
                } else {
                    ''
                }
        }

        $existingReason = if ($headerMap.ContainsKey('Reason')) { [string]$Worksheet.Cells.Item($row, $headerMap['Reason']).Text } else { '' }
        $existingReasonCategory = if ($headerMap.ContainsKey('ReasonCategory')) { [string]$Worksheet.Cells.Item($row, $headerMap['ReasonCategory']).Text } else { '' }
        $audit = Get-TimeNormalizationAudit -Definitions $definitions -RawValuesByDisplayName $rawValuesByDisplayName -ExistingReason $existingReason
        $reasonCategory = Get-ReviewReasonCategories -Definitions $definitions -RawValuesByDisplayName $rawValuesByDisplayName -ExistingReason $existingReason -ExistingCategoryCsv $existingReasonCategory -Audit $audit

        try {
            foreach ($definition in $definitions) {
                $normalized = $audit.Results[$definition.DisplayName]
                $displayCell = $Worksheet.Cells.Item($row, $headerMap[$definition.DisplayName])
                $minutesCell = $Worksheet.Cells.Item($row, $headerMap[$definition.MinutesColumnName])

                $displayCell.Value2 = $normalized.NormalizedText
                if ($null -eq $normalized.MinutesFromMidnight) {
                    $minutesCell.ClearContents()
                } else {
                    $minutesCell.Value2 = [double]$normalized.MinutesFromMidnight
                }
            }
        } catch {
            throw "Review 正規化列の書き込みに失敗しました (row=$row): $($_.Exception.Message)"
        }

        $Worksheet.Cells.Item($row, $headerMap['ReasonCategory']).Value2 = $reasonCategory
        $Worksheet.Cells.Item($row, $headerMap['時刻正規化状態']).Value2 = $audit.Status
        $Worksheet.Cells.Item($row, $headerMap['時刻確認メモ']).Value2 = $audit.Note
    }

    foreach ($definition in $definitions) {
        try {
            $Worksheet.Cells.Item(1, $headerMap[$definition.DisplayName]).EntireColumn.NumberFormat = '@'
            $Worksheet.Cells.Item(1, $headerMap[$definition.MinutesColumnName]).EntireColumn.NumberFormat = '0'
        } catch {
            throw "Review 正規化列 '$($definition.DisplayName)' の表示形式設定に失敗しました: $($_.Exception.Message)"
        }
    }
    $Worksheet.Columns.AutoFit() | Out-Null
}

function Apply-VersionSpecificWorkbookEnrichments {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)]$Profile
    )

    if ($VersionMode -ne 'v2') {
        return
    }

    $resultSheet = $null
    $reviewSheet = $null
    try {
        $resultSheet = $Workbook.Worksheets.Item('Result')
        Add-V2ResultNormalizedColumns -Worksheet $resultSheet -Profile $Profile
        $reviewSheet = $Workbook.Worksheets.Item('Review')
        Add-V2ReviewNormalizedColumns -Worksheet $reviewSheet -Profile $Profile
    } finally {
        $reviewSheet | Release-ComObject
        $resultSheet | Release-ComObject
    }
}

function Measure-WorkbookOutcome {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)][int]$SourcePdfCount
    )

    $resultSummarySheet = $null
    $errorsSummarySheet = $null
    $reviewSummarySheet = $null

    try {
        $resultSummarySheet = $Workbook.Worksheets.Item('Result')
        $errorsSummarySheet = $Workbook.Worksheets.Item('Errors')
        $resultRowCount = Get-DataRowCount -Worksheet $resultSummarySheet -ExpectedTableName 'tblResult'
        $errorRowCount = Get-DataRowCount -Worksheet $errorsSummarySheet -ExpectedTableName 'tblErrors'
        if ($VersionMode -eq 'v2') {
            $reviewSummarySheet = $Workbook.Worksheets.Item('Review')
            $reviewRowCount = Get-DataRowCount -Worksheet $reviewSummarySheet -ExpectedTableName 'tblReview'
        } else {
            $reviewRowCount = 0
        }
        $failedPdfCount = $errorRowCount
        $successPdfCount = [Math]::Max($SourcePdfCount - $failedPdfCount, 0)
    } finally {
        $reviewSummarySheet | Release-ComObject
        $errorsSummarySheet | Release-ComObject
        $resultSummarySheet | Release-ComObject
    }

    return [pscustomobject]@{
        ResultRowCount  = $resultRowCount
        ErrorRowCount   = $errorRowCount
        ReviewRowCount  = $reviewRowCount
        SuccessPdfCount = $successPdfCount
        FailedPdfCount  = $failedPdfCount
        ElapsedSeconds  = [int][Math]::Ceiling(((Get-Date) - $script:runStartedAt).TotalSeconds)
    }
}

function Update-WorkbookSummaryState {
    param(
        [Parameter(Mandatory = $true)]$ControlSheet,
        [Parameter(Mandatory = $true)]$SummarySheet,
        [Parameter(Mandatory = $true)]$RunPlan,
        [Parameter(Mandatory = $true)]$Outcome
    )

    Set-ControlMetrics -Worksheet $ControlSheet -SourcePdfCount $RunPlan.SourceFiles.Count -ResultRowCount $Outcome.ResultRowCount -ErrorRowCount $Outcome.ErrorRowCount -SuccessPdfCount $Outcome.SuccessPdfCount -FailedPdfCount $Outcome.FailedPdfCount -ElapsedSeconds $Outcome.ElapsedSeconds
    Set-SummaryMetrics -Worksheet $SummarySheet -SourcePdfCount $RunPlan.SourceFiles.Count -SuccessPdfCount $Outcome.SuccessPdfCount -FailedPdfCount $Outcome.FailedPdfCount -ResultRowCount $Outcome.ResultRowCount -ErrorRowCount $Outcome.ErrorRowCount -ReviewRowCount $Outcome.ReviewRowCount -ElapsedSeconds $Outcome.ElapsedSeconds -Profile $RunPlan.Profile -OutputPath $RunPlan.OutputPath -VersionMode $VersionMode
    $ControlSheet.Range('B6').Value2 = '出力準備完了'
}

function Publish-WorkbookOutput {
    param(
        [Parameter(Mandatory = $true)]$Excel,
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)][string]$RuntimeWorkbookPath,
        [Parameter(Mandatory = $true)][string]$OutputPath
    )

    if ($script:isSecureMode) {
        Write-Log 'VER2 Secure のため、PowerShell 側で xlsx を保存しています。'
        Export-WorkbookDirectly -Workbook $Workbook -OutputPath $OutputPath
        return
    }

    $usedMacro = $false
    try {
        Write-Log 'Excel 出力処理を実行しています。'
        Try-RunMacro -Excel $Excel -WorkbookPath $RuntimeWorkbookPath -MacroName 'ExportResultAsXlsx'
        $usedMacro = $true
    } catch {
        Write-Log "VBA マクロが使えなかったため、PowerShell 側で xlsx を保存します。理由: $($_.Exception.Message)" 'WARN'
    }

    if (-not $usedMacro) {
        Write-Log 'PowerShell 側で xlsx を保存しています。'
        Export-WorkbookDirectly -Workbook $Workbook -OutputPath $OutputPath
    }
}

function Finalize-RunWorkspace {
    param([Nullable[int]]$ExcelProcessId)

    if (Test-Path -LiteralPath $script:runWorkspaceDir) {
        $deleted = Remove-PathWithRetry -Path $script:runWorkspaceDir
        if (-not $deleted -and $script:isSecureMode) {
            $stopped = Stop-ExcelProcessForCleanup -ProcessId $ExcelProcessId -Reason 'Secure finally で一時領域を削除するため'
            if ($stopped) {
                $deleted = Remove-PathWithRetry -Path $script:runWorkspaceDir -MaxAttempts 5
            }
        }
        if (-not $deleted) {
            $script:cleanupFailureMessage = '一時領域の削除に失敗しました。端末再起動または管理者確認が必要です。'
            Write-Log ("{0} 対象: {1}" -f $script:cleanupFailureMessage, $script:runWorkspaceDir) 'ERROR'
            Write-Host ''
            Write-Host $script:cleanupFailureMessage -ForegroundColor Red
        }
    }
}

function Close-ExcelRuntimeContext {
    param(
        $RuntimeContext,
        [switch]$ForceStopProcess
    )

    if ($null -eq $RuntimeContext) {
        return
    }

    if ($RuntimeContext.Workbook) {
        try {
            $RuntimeContext.Workbook.Close($false)
        } catch {
        }
    }
    if ($RuntimeContext.Excel) {
        try {
            $RuntimeContext.Excel.Quit()
        } catch {
        }
    }

    foreach ($comObject in @($RuntimeContext.SummarySheet, $RuntimeContext.ControlSheet, $RuntimeContext.Workbook, $RuntimeContext.Excel)) {
        try {
            $comObject | Release-ComObject
        } catch {
        }
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    if ($ForceStopProcess) {
        Stop-ExcelProcessForCleanup -ProcessId $RuntimeContext.ExcelProcessId -Reason 'Excel COM 終了後もプロセスが残存したため' | Out-Null
    }
}

if ($SkipMain) {
    return
}

$runtimeContext = $null
$runtimeWorkbookPath = $null
$runPlan = $null
$resultRowCount = 0
$errorRowCount = 0
$reviewRowCount = 0
$successPdfCount = 0
$failedPdfCount = 0
$elapsedSeconds = 0
$runFailed = $false

try {
    Initialize-RunWorkspace
    Write-EnvironmentWarnings
    Write-Banner
    Assert-ExecutionLocationAllowed
    if ($script:isSecureMode) {
        Write-Log ("VER2 Secure モードで実行します。runtime はローカル領域を使用します: {0}" -f $runtimeRootDir)
        Write-Log ("VER2 Secure のログ出力先: {0}" -f $logsDir)
        Write-Log ("DEBUG ログは保守者向けにのみ有効です。現在のログレベル: {0}" -f $LogLevel) 'DEBUG'
    }
    Write-Log '入力 PDF を確認しています。'
    $runPlan = Resolve-ExecutionPlan
    Show-ProcessingNotice
    Prepare-RunInputs -SourceFiles $runPlan.SourceFiles | Out-Null

    Ensure-Template -ForceRebuild:$RebuildTemplate
    $runtimeWorkbookPath = Copy-TemplateToRuntime
    $runtimeContext = Open-ExcelRuntimeContext -RuntimeWorkbookPath $runtimeWorkbookPath -OutputPath $runPlan.OutputPath -Profile $runPlan.Profile
    Configure-WorkbookQueries -Workbook $runtimeContext.Workbook -Profile $runPlan.Profile
    Load-WorkbookOutputSheets -Workbook $runtimeContext.Workbook
    Apply-VersionSpecificWorkbookEnrichments -Workbook $runtimeContext.Workbook -Profile $runPlan.Profile
    $runtimeContext.SummarySheet = $runtimeContext.Workbook.Worksheets.Item('Summary')

    $outcome = Measure-WorkbookOutcome -Workbook $runtimeContext.Workbook -SourcePdfCount $runPlan.SourceFiles.Count
    $resultRowCount = $outcome.ResultRowCount
    $errorRowCount = $outcome.ErrorRowCount
    $reviewRowCount = $outcome.ReviewRowCount
    $successPdfCount = $outcome.SuccessPdfCount
    $failedPdfCount = $outcome.FailedPdfCount
    $elapsedSeconds = $outcome.ElapsedSeconds

    Update-WorkbookSummaryState -ControlSheet $runtimeContext.ControlSheet -SummarySheet $runtimeContext.SummarySheet -RunPlan $runPlan -Outcome $outcome
    $runtimeContext.Workbook.Save()
    Publish-WorkbookOutput -Excel $runtimeContext.Excel -Workbook $runtimeContext.Workbook -RuntimeWorkbookPath $runtimeWorkbookPath -OutputPath $runPlan.OutputPath

    Write-Log "処理が完了しました: $($runPlan.OutputPath)"
    Show-RunSummary -SourceFiles $runPlan.SourceFiles -OutputPath $runPlan.OutputPath -Profile $runPlan.Profile -ResultRows $resultRowCount -ErrorRows $errorRowCount -ReviewRows $reviewRowCount -SuccessPdfCount $successPdfCount -FailedPdfCount $failedPdfCount -ElapsedSeconds $elapsedSeconds

    if ($OpenOutput) {
        Invoke-Item -LiteralPath $runPlan.OutputPath
    }

    if ($OpenOutputFolder) {
        Invoke-Item -LiteralPath $runPlan.OutputParent
    }
} catch {
    $runFailed = $true
    $runError = Resolve-RunErrorInfo -Message $_.Exception.Message
    Write-Log ("処理に失敗しました: [{0}] [{1}] {2}" -f $runError.ErrorCategory, $runError.ErrorCode, $_.Exception.Message) 'ERROR'
    if ($runtimeContext -and $runtimeContext.ControlSheet) {
        try {
            $runtimeContext.ControlSheet.Range('B6').Value2 = '失敗'
            $runtimeContext.Workbook.Save()
        } catch {
        }
    }
    throw
} finally {
    Close-ExcelRuntimeContext -RuntimeContext $runtimeContext -ForceStopProcess:$script:isSecureMode
    if ($script:isSecureMode) {
        Write-Log $(if ($runFailed) { 'VER2 Secure のため、finally で失敗時の一時領域を即時削除します。' } else { 'VER2 Secure のため、finally で一時領域を即時削除します。' })
    }
    Finalize-RunWorkspace -ExcelProcessId $(if ($runtimeContext) { $runtimeContext.ExcelProcessId } else { $null })
    Release-RunLock
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($script:cleanupFailureMessage -and -not $runFailed) {
        throw $script:cleanupFailureMessage
    }
}
