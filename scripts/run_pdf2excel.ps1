param(
    [string]$InputFolder,
    [string[]]$InputFiles,
    [string]$OutputFile,
    [string]$ProfileName = 'default',
    [string]$ProfilePath,
    [switch]$KeepInput,
    [switch]$RebuildTemplate,
    [switch]$OpenOutput,
    [switch]$OpenOutputFolder,
    [switch]$SelectInputFolder,
    [switch]$PromptForOutputFile,
    [switch]$NoConfirm,
    [Alias('h')][switch]$Help
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms

$script:runStartedAt = Get-Date
$baseDir = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$scriptDir = Join-Path $baseDir 'scripts'
$templateDir = Join-Path $baseDir 'template'
$configDir = Join-Path $baseDir 'config'
$profilesDir = Join-Path $configDir 'profiles'
$inputDir = Join-Path $baseDir 'input'
$outputDir = Join-Path $baseDir 'output'
$runtimeRootDir = Join-Path $outputDir 'runtime'
$runtimeRunsDir = Join-Path $runtimeRootDir 'runs'
$logsDir = Join-Path $baseDir 'logs'
$templatePath = Join-Path $templateDir 'PDF2Excel_Converter.xlsm'
$buildTemplateScript = Join-Path $scriptDir 'build_excel_template.ps1'

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
$script:runInstanceId = "run_${timestamp}_$PID"
$script:runWorkspaceDir = Join-Path $runtimeRunsDir $script:runInstanceId
$script:runStagingDir = Join-Path $script:runWorkspaceDir 'staging'
$script:runRuntimeDir = Join-Path $script:runWorkspaceDir 'runtime'
$script:lockFilePath = Join-Path $runtimeRootDir 'run.lock'
$script:runMutex = $null
$script:logPath = Join-Path $logsDir "run_$timestamp.log"

function Show-Usage {
    @(
        'PDF2Excel 使い方',
        '',
        '1. かんたん操作:',
        '   run_pdf2excel.bat をダブルクリックします。',
        '',
        '2. PowerShell から直接実行:',
        '',
        '   powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\run_pdf2excel.ps1',
        '   powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\run_pdf2excel.ps1 -InputFolder C:\Path\To\PdfFolder',
        '   powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\run_pdf2excel.ps1 -InputFolder C:\Path\To\PdfFolder -OutputFile C:\Path\To\result.xlsx',
        '   powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\run_pdf2excel.ps1 -SelectInputFolder -PromptForOutputFile',
        '',
        'オプション',
        '  -InputFolder         PDF が入っているフォルダを指定します。',
        '  -InputFiles          変換対象の PDF ファイルを個別指定します。',
        '  -OutputFile          出力する xlsx の保存先を指定します。',
        '  -ProfileName         使用する帳票プロファイル名を指定します。既定値は default です。',
        '  -ProfilePath         使用する帳票プロファイル JSON のフルパスを指定します。',
        '  -KeepInput           input 内の過去PDFを保持します。実際の変換は今回分だけ別 staging で実行します。',
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
        [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )

    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Add-Content -LiteralPath $script:logPath -Value $line -Encoding UTF8
    Write-Host $line
}

function Write-Banner {
    Write-Host ''
    Write-Host '==========================================' -ForegroundColor Cyan
    Write-Host ' PDF2Excel - PDF表 一括変換ツール' -ForegroundColor Cyan
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
    Write-Host ('  処理時間(秒)    : {0}' -f $ElapsedSeconds)
    Write-Host ('  使用プロファイル : {0}' -f $Profile.DisplayName)
    Write-Host ('  出力ファイル     : {0}' -f $OutputPath)
    Write-Host ('  ログファイル     : {0}' -f $script:logPath)
    Write-Host ''
}

function Release-ComObject {
    param([Parameter(ValueFromPipeline = $true)]$InputObject)

    process {
        if ($null -ne $InputObject -and [System.Runtime.InteropServices.Marshal]::IsComObject($InputObject)) {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($InputObject)
        }
    }
}

function Ensure-Directory {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
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

function Compact-RuntimeArtifacts {
    if (-not (Test-Path -LiteralPath $runtimeRootDir)) {
        return
    }

    Get-ChildItem -LiteralPath $runtimeRunsDir -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        $deleted = Remove-PathWithRetry -Path $_.FullName
        if (-not $deleted) {
            Write-Log "Runtime artifact could not be removed: $($_.FullName)" 'WARN'
        }
    }

    Get-ChildItem -LiteralPath $runtimeRootDir -File -ErrorAction SilentlyContinue | Where-Object {
        $_.Name -notin @('.gitkeep', 'run.lock')
    } | ForEach-Object {
        $deleted = Remove-PathWithRetry -Path $_.FullName
        if (-not $deleted) {
            Write-Log "Runtime artifact could not be removed: $($_.FullName)" 'WARN'
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
                $lockSummary = " 実行中情報: 開始=$($lockInfo.startedAt), PID=$($lockInfo.pid), User=$($lockInfo.userName)"
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
        machineName   = $env:COMPUTERNAME
        userName      = $env:USERNAME
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
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = 'PDF が入っているフォルダを選択してください。'
    $dialog.ShowNewFolderButton = $false

    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        throw '入力フォルダが選択されませんでした。'
    }

    return $dialog.SelectedPath
}

function Select-PdfFiles {
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
            throw "プロファイルが見つかりません: $resolvedProfilePath"
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
        ProfilePath                = $resolvedProfilePath
    }
}

function Get-ProfileOutputColumnNames {
    param([Parameter(Mandatory = $true)]$Profile)

    $names = @($Profile.SourceFileColumnName)
    foreach ($index in 1..$Profile.ExpectedColumns) {
        $names += '{0}{1}' -f $Profile.DataColumnPrefix, $index
    }
    return $names
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
        Write-Log "Creating Excel template: $templatePath"
        & powershell -NoProfile -ExecutionPolicy Bypass -File $buildTemplateScript -TemplatePath $templatePath
        if ($LASTEXITCODE -ne 0) {
            throw "テンプレート生成スクリプトが失敗しました。終了コード: $LASTEXITCODE"
        }
    }

    if (-not (Test-Path -LiteralPath $templatePath)) {
        throw "テンプレートの作成に失敗しました: $templatePath"
    }
}

function Copy-TemplateToRuntime {
    Ensure-Directory -Path $script:runRuntimeDir
    $runtimePath = Join-Path $script:runRuntimeDir "PDF2Excel_runtime_$timestamp.xlsm"
    Copy-Item -LiteralPath $templatePath -Destination $runtimePath -Force
    return $runtimePath
}

function Escape-MString {
    param([Parameter(Mandatory = $true)][string]$Value)

    return $Value.Replace('"', '""')
}

function ConvertTo-MTextLiteral {
    param([Parameter(Mandatory = $true)][string]$Value)

    return '"' + (Escape-MString -Value $Value) + '"'
}

function ConvertTo-MTextListLiteral {
    param([string[]]$Values)

    if ($null -eq $Values -or $Values.Count -eq 0) {
        return '{}'
    }

    $items = @($Values | ForEach-Object { ConvertTo-MTextLiteral -Value $_ })
    return '{' + ($items -join ', ') + '}'
}

function ConvertTo-MLogicalLiteral {
    param([bool]$Value)

    if ($Value) {
        return 'true'
    }

    return 'false'
}

function Get-StagingQueryFormula {
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
    KeepColumns = Table.SelectColumns(PdfFiles, {"Name", "Content"}),
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
                                                    Score = score
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
                                rawData =
                                    if chosen = null then
                                        null
                                    else if AllowMoreColumns = false and chosenColumns <> null and chosenColumns > ExpectedColumns then
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
                                        Table.ReorderColumns(withFileName, OutputColumns, MissingField.UseNull)
                            in
                                [Data = reordered]
                in
                    if processingTry[HasError] then
                        [Data = null]
                    else
                        processingTry[Value],
            type record
        ),
    ExpandedProcessed = Table.ExpandRecordColumn(WithProcessed, "Processed", {"Data"}, {"Data"}),
    SuccessRows = Table.SelectRows(ExpandedProcessed, each [Data] <> null),
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
@"
let
    Source = PDF2Excel_Staging,
    Selected =
        Table.SelectColumns(
            Source,
            {"Name", "IsError", "OutputRowCount", "ErrorCategory", "ErrorCode", "UserMessage", "CandidateColumns", "CandidateRows", "SelectedTableKind", "SelectedTableName"},
            MissingField.UseNull
        ),
    WithStatus = Table.AddColumn(Selected, "Status", each if [IsError] = true then "Failed" else "Success", type text),
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
            throw "ListObjects.Add failed: $($_.Exception.Message)"
        }

        try {
            $listObject.Name = $TableName
            $queryTable = $listObject.QueryTable
        } catch {
            throw "QueryTable acquisition failed: $($_.Exception.Message)"
        }

        try {
            $queryTable.CommandType = 2
            $queryTable.CommandText = @("SELECT * FROM [$QueryName]")
            $queryTable.BackgroundQuery = $false
        } catch {
            throw "QueryTable configuration failed: $($_.Exception.Message)"
        }

        try {
            $queryTable.Refresh($false) | Out-Null
            $worksheet.Columns.AutoFit() | Out-Null
        } catch {
            throw "QueryTable refresh failed: $($_.Exception.Message)"
        }
    } catch {
        throw "Query load failed for $QueryName on sheet ${WorksheetName}: $($_.Exception.Message)"
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
        [Parameter(Mandatory = $true)]$Profile
    )

    $Worksheet.Range('A1').Value2 = '項目'
    $Worksheet.Range('B1').Value2 = '内容'
    $Worksheet.Range('A2').Value2 = '入力フォルダ'
    $Worksheet.Range('A3').Value2 = '出力ファイル'
    $Worksheet.Range('A4').Value2 = 'ログファイル'
    $Worksheet.Range('A5').Value2 = '最終実行日時'
    $Worksheet.Range('A6').Value2 = '状態'
    $Worksheet.Range('A7').Value2 = '対象PDF数'
    $Worksheet.Range('A8').Value2 = '取込データ行数'
    $Worksheet.Range('A9').Value2 = 'エラー件数'
    $Worksheet.Range('A10').Value2 = '成功PDF数'
    $Worksheet.Range('A11').Value2 = '失敗PDF数'
    $Worksheet.Range('A12').Value2 = '処理時間(秒)'
    $Worksheet.Range('A13').Value2 = '使用プロファイル'
    $Worksheet.Range('A14').Value2 = 'プロファイル説明'

    $Worksheet.Range('B2').Value2 = $StagingInputFolder
    $Worksheet.Range('B3').Value2 = $OutputPath
    $Worksheet.Range('B4').Value2 = $LogPath
    $Worksheet.Range('B5').Value2 = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    $Worksheet.Range('B6').Value2 = '準備完了'
    $Worksheet.Range('B7:B12').Value2 = ''
    $Worksheet.Range('B13').Value2 = $Profile.DisplayName
    $Worksheet.Range('B14').Value2 = $Profile.Description
    $Worksheet.Range('A16').Value2 = 'かんたんな使い方'
    $Worksheet.Range('B16').Value2 = '1. run_pdf2excel.bat を実行  2. PDF を選択  3. 実行前チェックを確認  4. Result / Summary / Errors を確認'
    $Worksheet.Range('A17').Value2 = '確認ポイント'
    $Worksheet.Range('B17').Value2 = 'Summary は件数の全体像、Result は変換成功データ、Errors は失敗した PDF と理由です。'
    $Worksheet.Range('A18').Value2 = '注意'
    $Worksheet.Range('B18').Value2 = '文字を選択できるテキスト PDF と、ほぼ同じレイアウトの帳票を想定しています。'
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
    param([Parameter(Mandatory = $true)]$Worksheet)

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
    $Worksheet.Range('A1').Value2 = 'Summary'
    $Worksheet.Range('A2').Value2 = '実行結果の集計と、PDFごとの内訳を表示します。'
    $Worksheet.Range('A4').Value2 = '項目'
    $Worksheet.Range('B4').Value2 = '内容'
    $Worksheet.Range('A13').Value2 = 'PDF別サマリー'
    $Worksheet.Range('M13').Value2 = 'エラー分類別件数'
    $Worksheet.Range('A1').Font.Bold = $true
    $Worksheet.Range('A1').Font.Size = 14
    $Worksheet.Range('A4:B4').Font.Bold = $true
    $Worksheet.Range('A4:B4').Interior.Color = 15773696
    $Worksheet.Range('A13').Font.Bold = $true
    $Worksheet.Range('M13').Font.Bold = $true
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
        [int]$ElapsedSeconds,
        [Parameter(Mandatory = $true)]$Profile,
        [Parameter(Mandatory = $true)][string]$OutputPath
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
    $Worksheet.Range('A10').Value2 = '処理時間(秒)'
    $Worksheet.Range('B10').Value2 = $ElapsedSeconds
    $Worksheet.Range('A11').Value2 = '使用プロファイル'
    $Worksheet.Range('B11').Value2 = $Profile.DisplayName
    $Worksheet.Range('A12').Value2 = '出力ファイル'
    $Worksheet.Range('B12').Value2 = $OutputPath
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

function Resolve-RunErrorInfo {
    param([Parameter(Mandatory = $true)][string]$Message)

    $normalizedMessage = $Message.ToLowerInvariant()

    if ($normalizedMessage.Contains('実行前チェックでキャンセル')) {
        return [pscustomobject]@{ ErrorCode = 'RUN_CANCELLED'; ErrorCategory = '実行キャンセル' }
    }
    if ($normalizedMessage.Contains('同名の pdf')) {
        return [pscustomobject]@{ ErrorCode = 'DUPLICATE_FILE_NAME'; ErrorCategory = '入力エラー' }
    }
    if ($normalizedMessage.Contains('別の pdf2excel 実行が進行中')) {
        return [pscustomobject]@{ ErrorCode = 'RUN_LOCKED'; ErrorCategory = '実行競合' }
    }
    if ($normalizedMessage.Contains('入力フォルダ') -or $normalizedMessage.Contains('入力ファイル')) {
        return [pscustomobject]@{ ErrorCode = 'INPUT_RESOLUTION_ERROR'; ErrorCategory = '入力エラー' }
    }
    if ($normalizedMessage.Contains('テンプレート')) {
        return [pscustomobject]@{ ErrorCode = 'TEMPLATE_ERROR'; ErrorCategory = 'テンプレートエラー' }
    }
    if ($normalizedMessage.Contains('excel')) {
        return [pscustomobject]@{ ErrorCode = 'EXCEL_RUNTIME_ERROR'; ErrorCategory = 'Excel実行エラー' }
    }

    return [pscustomobject]@{ ErrorCode = 'UNEXPECTED_RUN_ERROR'; ErrorCategory = 'システムエラー' }
}

$excel = $null
$workbook = $null
$runtimeWorkbookPath = $null
$controlSheet = $null
$summarySheet = $null
$resultRowCount = 0
$errorRowCount = 0
$successPdfCount = 0
$failedPdfCount = 0
$elapsedSeconds = 0

try {
    Ensure-Workspace
    Set-Content -LiteralPath $script:logPath -Value "PDF2Excel run started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Encoding UTF8
    Acquire-RunLock
    Compact-RuntimeArtifacts
    Ensure-Directory -Path $script:runWorkspaceDir
    Ensure-Directory -Path $script:runStagingDir
    Ensure-Directory -Path $script:runRuntimeDir

    Write-Banner
    Write-Log '入力 PDF を確認しています。'
    $inputFileCandidates = @($InputFiles)
    if ($inputFileCandidates.Count -gt 0) {
        $inputFileCandidates += @($args)
    } elseif ($args.Count -gt 0) {
        throw "不明な引数があります: $($args -join ', ')"
    }

    $sourceFiles = @(Resolve-InputPdfFiles -SourceFolder $InputFolder -SourceFiles $inputFileCandidates)
    Write-Log ("対象 PDF 数: {0}" -f $sourceFiles.Count)

    $profile = Get-ProfileConfiguration -RequestedProfileName $ProfileName -RequestedProfilePath $ProfilePath
    Write-Log ("使用プロファイル: {0} ({1})" -f $profile.DisplayName, $profile.ProfilePath)

    $defaultOutputPath = Join-Path $outputDir "PDF2Excel_$timestamp.xlsx"
    if ([string]::IsNullOrWhiteSpace($OutputFile)) {
        $OutputFile = if ($PromptForOutputFile) { Select-OutputFileDialog -DefaultOutputPath $defaultOutputPath } else { $defaultOutputPath }
    }

    if (-not [string]::IsNullOrWhiteSpace($OutputFile)) {
        $OutputFile = [System.IO.Path]::GetFullPath($OutputFile)
    }

    $outputParent = Split-Path -Path $OutputFile -Parent
    Ensure-Directory -Path $outputParent

    $preflightState = Get-PreflightState -SourceFiles $sourceFiles -OutputPath $OutputFile -Profile $profile
    Confirm-Preflight -PreflightState $preflightState

    $stagedFiles = Stage-PdfFiles -Files $sourceFiles -StagingDirectory $script:runStagingDir
    $storedFiles = Sync-InputStorage -Files $sourceFiles -KeepExisting:$KeepInput
    Write-Log ("今回実行分の staging が完了しました: {0}" -f ($stagedFiles -join ', '))
    Write-Log ("input フォルダ同期が完了しました: {0}" -f ($storedFiles -join ', '))
    Start-Sleep -Seconds 1

    Ensure-Template -ForceRebuild:$RebuildTemplate
    $runtimeWorkbookPath = Copy-TemplateToRuntime

    Write-Log 'Excel を起動しています。'
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 1

    $workbook = $excel.Workbooks.Open($runtimeWorkbookPath)
    $controlSheet = Get-OrCreateWorksheet -Workbook $workbook -WorksheetName 'Control'
    $summarySheet = Get-OrCreateWorksheet -Workbook $workbook -WorksheetName 'Summary'
    $resultSheetRef = Get-OrCreateWorksheet -Workbook $workbook -WorksheetName 'Result'
    $errorsSheetRef = Get-OrCreateWorksheet -Workbook $workbook -WorksheetName 'Errors'
    $resultSheetRef | Release-ComObject
    $errorsSheetRef | Release-ComObject

    Set-ControlValues -Worksheet $controlSheet -StagingInputFolder $script:runStagingDir -OutputPath $OutputFile -LogPath $script:logPath -Profile $profile
    Initialize-SummarySheet -Worksheet $summarySheet

    Write-Log 'Power Query を設定しています。'
    Add-OrReplaceWorkbookQuery -Workbook $workbook -QueryName 'PDF2Excel_Staging' -Formula (Get-StagingQueryFormula -InputPath $script:runStagingDir -Profile $profile)
    Add-OrReplaceWorkbookQuery -Workbook $workbook -QueryName 'PDF2Excel_Result' -Formula (Get-ResultQueryFormula -InputPath $script:runStagingDir -Profile $profile)
    Add-OrReplaceWorkbookQuery -Workbook $workbook -QueryName 'PDF2Excel_Errors' -Formula (Get-ErrorsQueryFormula)
    Add-OrReplaceWorkbookQuery -Workbook $workbook -QueryName 'PDF2Excel_FileSummary' -Formula (Get-FileSummaryQueryFormula)
    Add-OrReplaceWorkbookQuery -Workbook $workbook -QueryName 'PDF2Excel_ErrorSummary' -Formula (Get-ErrorSummaryQueryFormula)

    Write-Log 'Result / Errors / Summary シートへ読み込んでいます。'
    Load-WorkbookQueryToWorksheet -Workbook $workbook -WorksheetName 'Result' -QueryName 'PDF2Excel_Result' -TableName 'tblResult'
    Load-WorkbookQueryToWorksheet -Workbook $workbook -WorksheetName 'Errors' -QueryName 'PDF2Excel_Errors' -TableName 'tblErrors'
    Load-WorkbookQueryToWorksheet -Workbook $workbook -WorksheetName 'Summary' -QueryName 'PDF2Excel_FileSummary' -TableName 'tblFileSummary' -DestinationAddress 'A14' -ClearSheet:$false
    Load-WorkbookQueryToWorksheet -Workbook $workbook -WorksheetName 'Summary' -QueryName 'PDF2Excel_ErrorSummary' -TableName 'tblErrorSummary' -DestinationAddress 'M14' -ClearSheet:$false
    $summarySheet = $workbook.Worksheets.Item('Summary')

    $resultSummarySheet = $null
    $errorsSummarySheet = $null
    try {
        $resultSummarySheet = $workbook.Worksheets.Item('Result')
        $errorsSummarySheet = $workbook.Worksheets.Item('Errors')
        $resultRowCount = Get-DataRowCount -Worksheet $resultSummarySheet -ExpectedTableName 'tblResult'
        $errorRowCount = Get-DataRowCount -Worksheet $errorsSummarySheet -ExpectedTableName 'tblErrors'
        $failedPdfCount = $errorRowCount
        $successPdfCount = [Math]::Max($sourceFiles.Count - $failedPdfCount, 0)
    } finally {
        $errorsSummarySheet | Release-ComObject
        $resultSummarySheet | Release-ComObject
    }

    $elapsedSeconds = [int][Math]::Ceiling(((Get-Date) - $script:runStartedAt).TotalSeconds)
    Set-ControlMetrics -Worksheet $controlSheet -SourcePdfCount $sourceFiles.Count -ResultRowCount $resultRowCount -ErrorRowCount $errorRowCount -SuccessPdfCount $successPdfCount -FailedPdfCount $failedPdfCount -ElapsedSeconds $elapsedSeconds
    Set-SummaryMetrics -Worksheet $summarySheet -SourcePdfCount $sourceFiles.Count -SuccessPdfCount $successPdfCount -FailedPdfCount $failedPdfCount -ResultRowCount $resultRowCount -ErrorRowCount $errorRowCount -ElapsedSeconds $elapsedSeconds -Profile $profile -OutputPath $OutputFile
    $controlSheet.Range('B6').Value2 = '出力準備完了'
    $workbook.Save()

    $usedMacro = $false
    try {
        Write-Log 'Excel 出力処理を実行しています。'
        Try-RunMacro -Excel $excel -WorkbookPath $runtimeWorkbookPath -MacroName 'ExportResultAsXlsx'
        $usedMacro = $true
    } catch {
        Write-Log "VBA マクロが使えなかったため、PowerShell 側で xlsx を保存します。理由: $($_.Exception.Message)" 'WARN'
    }

    if (-not $usedMacro) {
        Write-Log 'PowerShell 側で xlsx を保存しています。'
        Export-WorkbookDirectly -Workbook $workbook -OutputPath $OutputFile
    }

    Write-Log "処理が完了しました: $OutputFile"
    Show-RunSummary -SourceFiles $sourceFiles -OutputPath $OutputFile -Profile $profile -ResultRows $resultRowCount -ErrorRows $errorRowCount -SuccessPdfCount $successPdfCount -FailedPdfCount $failedPdfCount -ElapsedSeconds $elapsedSeconds

    if ($OpenOutput) {
        Invoke-Item -LiteralPath $OutputFile
    }

    if ($OpenOutputFolder) {
        Invoke-Item -LiteralPath $outputParent
    }
} catch {
    $runError = Resolve-RunErrorInfo -Message $_.Exception.Message
    Write-Log ("処理に失敗しました: [{0}] [{1}] {2}" -f $runError.ErrorCategory, $runError.ErrorCode, $_.Exception.Message) 'ERROR'
    if ($controlSheet) {
        try {
            $controlSheet.Range('B6').Value2 = '失敗'
            $workbook.Save()
        } catch {
        }
    }
    throw
} finally {
    if ($workbook) {
        try {
            $workbook.Close($false)
        } catch {
        }
    }
    if ($excel) {
        try {
            $excel.Quit()
        } catch {
        }
    }

    foreach ($comObject in @($summarySheet, $controlSheet, $workbook, $excel)) {
        try {
            $comObject | Release-ComObject
        } catch {
        }
    }

    if (Test-Path -LiteralPath $script:runWorkspaceDir) {
        $deleted = Remove-PathWithRetry -Path $script:runWorkspaceDir
        if (-not $deleted) {
            Write-Log "Run workspace could not be removed: $script:runWorkspaceDir" 'WARN'
        }
    }
    Release-RunLock
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
