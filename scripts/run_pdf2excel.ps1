param(
    [string]$InputFolder,
    [string]$OutputFile,
    [switch]$KeepInput,
    [switch]$RebuildTemplate,
    [switch]$OpenOutput
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms

$baseDir = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$scriptDir = Join-Path $baseDir 'scripts'
$templateDir = Join-Path $baseDir 'template'
$inputDir = Join-Path $baseDir 'input'
$outputDir = Join-Path $baseDir 'output'
$runtimeDir = Join-Path $outputDir 'runtime'
$logsDir = Join-Path $baseDir 'logs'
$templatePath = Join-Path $templateDir 'PDF2Excel_Converter.xlsm'
$buildTemplateScript = Join-Path $scriptDir 'build_excel_template.ps1'

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$script:logPath = Join-Path $logsDir "run_$timestamp.log"

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )

    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Add-Content -LiteralPath $script:logPath -Value $line -Encoding UTF8
    Write-Host $line
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
            Remove-Item -LiteralPath $Path -Force -ErrorAction Stop
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
    foreach ($path in @($inputDir, $outputDir, $runtimeDir, $logsDir, $templateDir, (Join-Path $templateDir 'vba'))) {
        Ensure-Directory -Path $path
    }
}

function Compact-RuntimeArtifacts {
    if (-not (Test-Path -LiteralPath $runtimeDir)) {
        return
    }

    Get-ChildItem -LiteralPath $runtimeDir -File -ErrorAction SilentlyContinue | ForEach-Object {
        $deleted = Remove-PathWithRetry -Path $_.FullName
        if (-not $deleted) {
            Write-Log "Runtime artifact could not be removed: $($_.FullName)" 'WARN'
        }
    }
}

function Select-PdfFiles {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = 'Select PDF files'
    $dialog.Filter = 'PDF files (*.pdf)|*.pdf'
    $dialog.Multiselect = $true

    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        throw 'No PDF files were selected.'
    }

    return @($dialog.FileNames)
}

function Resolve-InputPdfFiles {
    param([string]$SourceFolder)

    if ([string]::IsNullOrWhiteSpace($SourceFolder)) {
        return Select-PdfFiles
    }

    if (-not (Test-Path -LiteralPath $SourceFolder)) {
        throw "InputFolder does not exist: $SourceFolder"
    }

    $folderPath = (Resolve-Path -LiteralPath $SourceFolder).Path
    $files = Get-ChildItem -LiteralPath $folderPath -Filter '*.pdf' -File | Sort-Object Name | Select-Object -ExpandProperty FullName
    if (-not $files) {
        throw "No PDF files were found in the specified folder: $folderPath"
    }

    return @($files)
}

function Stage-PdfFiles {
    param(
        [Parameter(Mandatory = $true)][string[]]$Files,
        [switch]$KeepExisting
    )

    $selectedTargets = @{}
    $normalizedPairs = @()
    foreach ($file in $Files) {
        $sourcePath = [System.IO.Path]::GetFullPath((Resolve-Path -LiteralPath $file).Path)
        $targetPath = Join-Path $inputDir ([System.IO.Path]::GetFileName($sourcePath))
        $targetKey = $targetPath.ToLowerInvariant()

        if ($selectedTargets.ContainsKey($targetKey) -and $selectedTargets[$targetKey] -ne $sourcePath) {
            throw "Duplicate PDF file names are not supported: $([System.IO.Path]::GetFileName($sourcePath))"
        }

        $selectedTargets[$targetKey] = $sourcePath
        $normalizedPairs += [pscustomobject]@{
            SourcePath = $sourcePath
            TargetPath = $targetPath
            TargetKey  = $targetKey
        }
    }

    if (-not $KeepExisting) {
        Get-ChildItem -LiteralPath $inputDir -Filter '*.pdf' -File -ErrorAction SilentlyContinue | ForEach-Object {
            $existingPath = [System.IO.Path]::GetFullPath($_.FullName)
            if (-not $selectedTargets.ContainsKey($existingPath.ToLowerInvariant())) {
                Remove-Item -LiteralPath $existingPath -Force
            }
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
        throw 'Failed to stage PDF files.'
    }

    return $staged
}

function Ensure-Template {
    param([switch]$ForceRebuild)

    if ($ForceRebuild -or -not (Test-Path -LiteralPath $templatePath)) {
        Write-Log "Creating Excel template: $templatePath"
        & powershell -NoProfile -ExecutionPolicy Bypass -File $buildTemplateScript -TemplatePath $templatePath
    }

    if (-not (Test-Path -LiteralPath $templatePath)) {
        throw "Template creation failed: $templatePath"
    }
}

function Copy-TemplateToRuntime {
    $runtimePath = Join-Path $runtimeDir "PDF2Excel_runtime_$timestamp.xlsm"
    Copy-Item -LiteralPath $templatePath -Destination $runtimePath -Force
    return $runtimePath
}

function Escape-MString {
    param([Parameter(Mandatory = $true)][string]$Value)

    return $Value.Replace('"', '""')
}

function Get-StagingQueryFormula {
    param([Parameter(Mandatory = $true)][string]$InputPath)

    $escapedPath = Escape-MString -Value $InputPath
@"
let
    ExpectedColumns = 30,
    Source = Folder.Files("$escapedPath"),
    PdfFiles = Table.SelectRows(Source, each Text.Lower([Extension]) = ".pdf"),
    KeepColumns = Table.SelectColumns(PdfFiles, {"Name", "Extension", "Folder Path", "Content"}),
    WithProcessed = Table.AddColumn(
        KeepColumns,
        "Processed",
        each
            let
                fileName = [Name],
                pdfTry = try Pdf.Tables([Content]),
                pdfTables = if pdfTry[HasError] then null else pdfTry[Value],
                pdfError = if pdfTry[HasError] then Error.Message(pdfTry[Error]) else null,
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
                                score =
                                    if dataValue = null or rowCount = null then
                                        999999
                                    else
                                        Number.Abs(columnCount - ExpectedColumns) * 1000 + Number.Abs(rowCount - 100),
                                tableId = try Text.From(Record.Field(_, "Id")) otherwise "",
                                tableKind = try Text.From(Record.Field(_, "Kind")) otherwise "",
                                tableName = try Text.From(Record.Field(_, "Name")) otherwise ""
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
                viableCandidates = List.Select(scoredCandidates, each [Data] <> null and [RowCount] <> null and [RowCount] > 1),
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
                failureReason =
                    if pdfTry[HasError] then
                        "Pdf.Tables failed: " & pdfError
                    else if chosen = null then
                        "No suitable table was detected."
                    else if chosenColumns <> null and chosenColumns > ExpectedColumns then
                        "Column count exceeded 30: " & Text.From(chosenColumns)
                    else
                        null,
                rawData =
                    if failureReason <> null or chosen = null then
                        null
                    else
                        Table.Skip(chosen[Data], 1),
                originalColumns = if rawData = null then {} else Table.ColumnNames(rawData),
                renamed =
                    if rawData = null then
                        null
                    else
                        Table.RenameColumns(
                            rawData,
                            List.Transform(List.Positions(originalColumns), each {originalColumns{_}, "Column" & Text.From(_ + 1)}),
                            MissingField.Ignore
                        ),
                renamedCount = if renamed = null then 0 else Table.ColumnCount(renamed),
                missingColumns =
                    if renamed = null or renamedCount >= ExpectedColumns then
                        {}
                    else
                        List.Transform({renamedCount + 1 .. ExpectedColumns}, each "Column" & Text.From(_)),
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
                        Table.SelectColumns(
                            padded,
                            List.Transform({1 .. ExpectedColumns}, each "Column" & Text.From(_)),
                            MissingField.UseNull
                        ),
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
                        Table.AddColumn(textified, "SourceFile", each fileName, type text),
                reordered =
                    if withFileName = null then
                        null
                    else
                        Table.ReorderColumns(
                            withFileName,
                            List.Combine({{"SourceFile"}, List.Transform({1 .. ExpectedColumns}, each "Column" & Text.From(_))})
                        )
            in
                [
                    IsError = failureReason <> null,
                    Reason = failureReason,
                    CandidateColumns = chosenColumns,
                    CandidateRows = chosenRows,
                    Data = reordered
                ],
        type record
    ),
    Expanded = Table.ExpandRecordColumn(
        WithProcessed,
        "Processed",
        {"IsError", "Reason", "CandidateColumns", "CandidateRows", "Data"},
        {"IsError", "Reason", "CandidateColumns", "CandidateRows", "Data"}
    )
in
    Expanded
"@
}

function Get-ResultQueryFormula {
@"
let
    ExpectedColumns = {"SourceFile", "Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29", "Column30"},
    Source = PDF2Excel_Staging,
    SuccessRows = Table.SelectRows(Source, each [IsError] <> true and [Data] <> null),
    NestedTables = List.RemoveNulls(Table.Column(SuccessRows, "Data")),
    Combined = if List.Count(NestedTables) = 0 then #table(ExpectedColumns, {}) else Table.Combine(NestedTables),
    Reordered = Table.ReorderColumns(Combined, ExpectedColumns, MissingField.UseNull)
in
    Reordered
"@
}

function Get-ErrorsQueryFormula {
@"
let
    Source = PDF2Excel_Staging,
    ErrorRows = Table.SelectRows(Source, each [IsError] = true),
    Selected = Table.SelectColumns(ErrorRows, {"Name", "Reason", "CandidateColumns", "CandidateRows"}, MissingField.UseNull),
    Renamed = Table.RenameColumns(Selected, {{"Name", "FileName"}, {"Reason", "ErrorReason"}})
in
    Renamed
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

function Load-WorkbookQueryToWorksheet {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)][string]$WorksheetName,
        [Parameter(Mandatory = $true)][string]$QueryName,
        [Parameter(Mandatory = $true)][string]$TableName,
        [string]$DestinationAddress = 'A1'
    )

    $worksheet = $null
    $listObject = $null
    $queryTable = $null

    try {
        $worksheet = $Workbook.Worksheets.Item($WorksheetName)
        while ($worksheet.ListObjects.Count -gt 0) {
            $listObject = $worksheet.ListObjects.Item(1)
            $listObject.Delete()
            $listObject | Release-ComObject
            $listObject = $null
        }
        $worksheet.Cells.Clear() | Out-Null

        $source = @("OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=`$Workbook$;Location=$QueryName;Extended Properties=`"`"`"")
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
        [Parameter(Mandatory = $true)][string]$LogPath
    )

    $Worksheet.Range('B2').Value2 = $StagingInputFolder
    $Worksheet.Range('B3').Value2 = $OutputPath
    $Worksheet.Range('B4').Value2 = $LogPath
    $Worksheet.Range('B5').Value2 = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    $Worksheet.Range('B6').Value2 = 'Prepared'
}

function Try-RunMacro {
    param(
        [Parameter(Mandatory = $true)]$Excel,
        [Parameter(Mandatory = $true)][string]$WorkbookPath,
        [Parameter(Mandatory = $true)][string]$MacroName
    )

    $macroTarget = "'$WorkbookPath'!$MacroName"
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

Ensure-Workspace
Set-Content -LiteralPath $script:logPath -Value "PDF2Excel run started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Encoding UTF8
Compact-RuntimeArtifacts

$excel = $null
$workbook = $null
$runtimeWorkbookPath = $null
$controlSheet = $null

try {
    Write-Log "Resolving input PDF files."
    $sourceFiles = Resolve-InputPdfFiles -SourceFolder $InputFolder
    Write-Log ("PDF files detected: {0}" -f $sourceFiles.Count)

    $stagedFiles = Stage-PdfFiles -Files $sourceFiles -KeepExisting:$KeepInput
    Write-Log ("Staging complete: {0}" -f ($stagedFiles -join ', '))
    Start-Sleep -Seconds 2

    Ensure-Template -ForceRebuild:$RebuildTemplate
    $runtimeWorkbookPath = Copy-TemplateToRuntime

    if ([string]::IsNullOrWhiteSpace($OutputFile)) {
        $OutputFile = Join-Path $outputDir "PDF2Excel_$timestamp.xlsx"
    } else {
        $OutputFile = [System.IO.Path]::GetFullPath($OutputFile)
    }

    $outputParent = Split-Path -Path $OutputFile -Parent
    Ensure-Directory -Path $outputParent

    Write-Log "Launching Excel."
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 1

    $workbook = $excel.Workbooks.Open($runtimeWorkbookPath)
    $controlSheet = $workbook.Worksheets.Item('Control')

    Set-ControlValues -Worksheet $controlSheet -StagingInputFolder $inputDir -OutputPath $OutputFile -LogPath $script:logPath

    Write-Log 'Configuring Power Query.'
    Add-OrReplaceWorkbookQuery -Workbook $workbook -QueryName 'PDF2Excel_Staging' -Formula (Get-StagingQueryFormula -InputPath $inputDir)
    Add-OrReplaceWorkbookQuery -Workbook $workbook -QueryName 'PDF2Excel_Result' -Formula (Get-ResultQueryFormula)
    Add-OrReplaceWorkbookQuery -Workbook $workbook -QueryName 'PDF2Excel_Errors' -Formula (Get-ErrorsQueryFormula)

    Write-Log 'Loading queries into Result and Errors sheets.'

    $resultLoadSheet = $null
    $resultLoadListObject = $null
    $resultLoadQueryTable = $null
    try {
        $resultLoadSheet = $workbook.Worksheets.Item('Result')
        $resultLoadSheet.Cells.Clear() | Out-Null
        $resultSource = @('OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=PDF2Excel_Result;Extended Properties=""')
        $resultLoadListObject = $resultLoadSheet.ListObjects.Add(0, $resultSource, $null, 1, $resultLoadSheet.Range('A1'))
        $resultLoadListObject.Name = 'tblResult'
        $resultLoadQueryTable = $resultLoadListObject.QueryTable
        $resultLoadQueryTable.CommandType = 2
        $resultLoadQueryTable.CommandText = @('SELECT * FROM [PDF2Excel_Result]')
        $resultLoadQueryTable.BackgroundQuery = $false
        $resultLoadQueryTable.Refresh($false) | Out-Null
        $resultLoadSheet.Columns.AutoFit() | Out-Null
    } finally {
        $resultLoadQueryTable | Release-ComObject
        $resultLoadListObject | Release-ComObject
        $resultLoadSheet | Release-ComObject
    }

    $errorsLoadSheet = $null
    $errorsLoadListObject = $null
    $errorsLoadQueryTable = $null
    try {
        $errorsLoadSheet = $workbook.Worksheets.Item('Errors')
        $errorsLoadSheet.Cells.Clear() | Out-Null
        $errorsSource = @('OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=PDF2Excel_Errors;Extended Properties=""')
        $errorsLoadListObject = $errorsLoadSheet.ListObjects.Add(0, $errorsSource, $null, 1, $errorsLoadSheet.Range('A1'))
        $errorsLoadListObject.Name = 'tblErrors'
        $errorsLoadQueryTable = $errorsLoadListObject.QueryTable
        $errorsLoadQueryTable.CommandType = 2
        $errorsLoadQueryTable.CommandText = @('SELECT * FROM [PDF2Excel_Errors]')
        $errorsLoadQueryTable.BackgroundQuery = $false
        $errorsLoadQueryTable.Refresh($false) | Out-Null
        $errorsLoadSheet.Columns.AutoFit() | Out-Null
    } finally {
        $errorsLoadQueryTable | Release-ComObject
        $errorsLoadListObject | Release-ComObject
        $errorsLoadSheet | Release-ComObject
    }

    $controlSheet.Range('B6').Value2 = 'Prepared'
    $workbook.Save()

    $usedMacro = $false
    try {
        Write-Log 'Running VBA macros.'
        Try-RunMacro -Excel $excel -WorkbookPath $runtimeWorkbookPath -MacroName 'ExportResultAsXlsx'
        $usedMacro = $true
    } catch {
        Write-Log "VBA macros were unavailable. Saving xlsx directly from PowerShell. Reason: $($_.Exception.Message)" 'WARN'
    }

    if (-not $usedMacro) {
        Write-Log 'Saving xlsx directly from PowerShell.'
        Export-WorkbookDirectly -Workbook $workbook -OutputPath $OutputFile
    }

    Write-Log "Completed: $OutputFile"

    if ($OpenOutput) {
        Invoke-Item -LiteralPath $OutputFile
    }
} catch {
    Write-Log "Processing failed: $($_.Exception.Message)" 'ERROR'
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

    foreach ($comObject in @($controlSheet, $workbook, $excel)) {
        try {
            $comObject | Release-ComObject
        } catch {
        }
    }

    if ($runtimeWorkbookPath) {
        $deleted = Remove-PathWithRetry -Path $runtimeWorkbookPath
        if (-not $deleted) {
            Write-Log "Runtime workbook could not be removed: $runtimeWorkbookPath" 'WARN'
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
