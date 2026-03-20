param(
    [ValidateSet('v1', 'v2')]
    [string]$VersionMode,
    [string]$ProfileName,
    [string]$DisplayName,
    [string]$OutputPath,
    [string]$SamplePdfPath,
    [string]$ProfileBaseDirOverride,
    [switch]$Wizard,
    [switch]$SkipMain,
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'pdf2excel.common.ps1')

[Console]::InputEncoding = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding = [Console]::OutputEncoding

$projectRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$profileBaseDir = if ([string]::IsNullOrWhiteSpace($ProfileBaseDirOverride)) {
    Join-Path $projectRoot ("config\profiles\{0}" -f $VersionMode)
} else {
    [System.IO.Path]::GetFullPath($ProfileBaseDirOverride)
}
$profileObject = $null
$script:readHostFailureLimit = 3

function Get-DefaultProfileScaffoldValues {
    param(
        [Parameter(Mandatory = $true)][string]$ResolvedVersionMode,
        [string]$ResolvedProfileName,
        [string]$ResolvedDisplayName
    )

    return [ordered]@{
        ProfileName                = $ResolvedProfileName
        DisplayName                = if ([string]::IsNullOrWhiteSpace($ResolvedDisplayName)) {
            if ($ResolvedVersionMode -eq 'v2') { '生データ転記サンプルプロファイル' } else { "$ResolvedProfileName プロファイル" }
        } else {
            $ResolvedDisplayName
        }
        Description                = if ($ResolvedVersionMode -eq 'v2') { '生データ転記向けのサンプルプロファイルです。列定義と review 列を調整して使います。' } else { '標準変換向けのサンプルプロファイルです。列数やヘッダー行数を調整して使います。' }
        ExpectedColumns            = if ($ResolvedVersionMode -eq 'v2') { 30 } else { 10 }
        HeaderRowsToSkip           = 1
        TargetRowCount             = if ($ResolvedVersionMode -eq 'v2') { 5 } else { 100 }
        AllowMoreColumns           = $false
        PreferredTableKinds        = @('Table')
        PreferredTableNameContains = @()
        PreferredTableIdContains   = @()
        SourceFileColumnName       = if ($ResolvedVersionMode -eq 'v2') { '元ファイル名' } else { 'SourceFile' }
        DataColumnPrefix           = if ($ResolvedVersionMode -eq 'v2') { '項目' } else { 'Column' }
        OutputColumnNames          = @(
            Get-ProfileDataColumnNames -Profile ([pscustomobject]@{
                    ExpectedColumns  = if ($ResolvedVersionMode -eq 'v2') { 30 } else { 10 }
                    DataColumnPrefix = if ($ResolvedVersionMode -eq 'v2') { '項目' } else { 'Column' }
                })
        )
        MultiPageMergeMode         = if ($ResolvedVersionMode -eq 'v2') { 'sameHeader' } else { 'single' }
        NormalizedTimeColumns      = @()
        ReviewPersonColumn         = $null
        ReviewSiteColumn           = $null
        ReviewInTimeColumn         = $null
        ReviewOutTimeColumn        = $null
    }
}

function Read-VersionMode {
    while ($true) {
        $selection = Read-HostSafely -Prompt '対象バージョンを選んでください (1=VER1, 2=VER2)'
        switch ($selection) {
            '1' { return 'v1' }
            '2' { return 'v2' }
            default { Write-Host '1 または 2 を入力してください。' }
        }
    }
}

function Read-HostSafely {
    param([Parameter(Mandatory = $true)][string]$Prompt)

    $readHostCommand = Get-Command Read-Host -ErrorAction SilentlyContinue
    if ([Console]::IsInputRedirected -and ($null -eq $readHostCommand -or $readHostCommand.CommandType -ne 'Function')) {
        Write-Host -NoNewline ("{0}: " -f $Prompt)
        $line = [Console]::In.ReadLine()
        if ($null -eq $line) {
            return ''
        }
        return $line
    }

    $failureCount = 0
    while ($true) {
        try {
            return Read-Host $Prompt
        } catch {
            $failureCount += 1
            $message = if ([string]::IsNullOrWhiteSpace($_.Exception.Message)) { '入力を取得できませんでした。' } else { $_.Exception.Message }
            if ($failureCount -ge $script:readHostFailureLimit) {
                throw "入力を取得できないため処理を中断しました: $message"
            }

            Write-Host ("入力を取得できませんでした。再試行します ({0}/{1}): {2}" -f $failureCount, $script:readHostFailureLimit, $message)
            Start-Sleep -Milliseconds 200
        }
    }
}

function Read-ProfileName {
    param([string]$DefaultValue)

    while ($true) {
        $prompt = if ([string]::IsNullOrWhiteSpace($DefaultValue)) { '内部名/ファイル名に使うプロファイル名を入力してください (ASCII のみ)' } else { "内部名/ファイル名に使うプロファイル名を入力してください (Enter で $DefaultValue)" }
        $value = Read-HostSafely -Prompt $prompt
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

    $value = Read-HostSafely -Prompt $(if ([string]::IsNullOrWhiteSpace($DefaultValue)) { $Prompt } else { "$Prompt (Enter で $DefaultValue)" })
    if ([string]::IsNullOrWhiteSpace($value)) {
        return $DefaultValue
    }

    return $value
}

function Read-ValidatedIntValue {
    param(
        [Parameter(Mandatory = $true)][string]$Prompt,
        [int]$DefaultValue,
        [int]$MinValue = 0
    )

    while ($true) {
        $rawValue = Read-OptionalValue -Prompt $Prompt -DefaultValue $DefaultValue
        $parsedValue = 0
        if (-not [int]::TryParse([string]$rawValue, [ref]$parsedValue)) {
            Write-Host '整数を入力してください。'
            continue
        }
        if ($parsedValue -lt $MinValue) {
            Write-Host ("{0} 以上の整数を入力してください。" -f $MinValue)
            continue
        }

        return $parsedValue
    }
}

function Read-NullableIntValue {
    param(
        [Parameter(Mandatory = $true)][string]$Prompt,
        [Nullable[int]]$DefaultValue,
        [int]$MinValue = 1
    )

    while ($true) {
        $defaultText = if ($null -eq $DefaultValue) { '' } else { [string]$DefaultValue }
        $value = Read-HostSafely -Prompt $(if ([string]::IsNullOrWhiteSpace($defaultText)) { "$Prompt (未設定のままにする場合は Enter)" } else { "$Prompt (Enter で $defaultText / 未設定にする場合は -)" })
        if ([string]::IsNullOrWhiteSpace($value)) {
            return $DefaultValue
        }
        if ($value.Trim() -eq '-') {
            return $null
        }

        $parsedValue = 0
        if (-not [int]::TryParse($value, [ref]$parsedValue)) {
            Write-Host '整数を入力してください。'
            continue
        }
        if ($parsedValue -lt $MinValue) {
            Write-Host ("{0} 以上の整数を入力してください。" -f $MinValue)
            continue
        }

        return $parsedValue
    }
}

function Read-YesNoValue {
    param(
        [Parameter(Mandatory = $true)][string]$Prompt,
        [bool]$DefaultValue = $false
    )

    while ($true) {
        $defaultLabel = if ($DefaultValue) { 'Y' } else { 'N' }
        $value = Read-HostSafely -Prompt "$Prompt (Y/N, Enter で $defaultLabel)"
        if ([string]::IsNullOrWhiteSpace($value)) {
            return $DefaultValue
        }

        switch ($value.Trim().ToUpperInvariant()) {
            'Y' { return $true }
            'YES' { return $true }
            'N' { return $false }
            'NO' { return $false }
            default { Write-Host 'Y または N を入力してください。' }
        }
    }
}

function ConvertTo-StringArrayFromCommaSeparated {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return @()
    }

    return @(
        $Value -split '[,、]' |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
}

function ConvertTo-NormalizedTimeColumnEntries {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return @()
    }

    $entries = @()
    foreach ($token in @($Value -split '[,、]' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        $parts = $token -split ':', 2
        if ($parts.Count -ne 2) {
            throw "正規化時刻列は '列番号:表示名' 形式で入力してください: $token"
        }

        $sourceColumn = 0
        if (-not [int]::TryParse($parts[0].Trim(), [ref]$sourceColumn) -or $sourceColumn -lt 1) {
            throw "列番号は 1 以上の整数で入力してください: $token"
        }

        $displayName = $parts[1].Trim()
        if ([string]::IsNullOrWhiteSpace($displayName)) {
            throw "表示名を省略できません: $token"
        }

        $entries += [ordered]@{
            sourceColumn = $sourceColumn
            displayName  = $displayName
        }
    }

    return @($entries)
}

function Read-NormalizedTimeColumns {
    param($DefaultEntries)

    while ($true) {
        $defaultText = if ($null -eq $DefaultEntries -or @($DefaultEntries).Count -eq 0) {
            ''
        } else {
            (@($DefaultEntries) | ForEach-Object { '{0}:{1}' -f $_.sourceColumn, $_.displayName }) -join ','
        }

        try {
            $rawValue = Read-HostSafely -Prompt $(if ([string]::IsNullOrWhiteSpace($defaultText)) { "正規化時刻列を入力してください (列番号:表示名 をカンマ区切り / 未設定は Enter)" } else { "正規化時刻列を入力してください (Enter で $defaultText / 未設定は -)" })
            if ([string]::IsNullOrWhiteSpace($rawValue)) {
                return if ([string]::IsNullOrWhiteSpace($defaultText)) { @() } else { @($DefaultEntries) }
            }
            if ($rawValue.Trim() -eq '-') {
                return @()
            }

            return @(ConvertTo-NormalizedTimeColumnEntries -Value $rawValue)
        } catch {
            Write-Host $_.Exception.Message
        }
    }
}

function Read-MultiPageMergeModeValue {
    param([string]$DefaultValue = 'sameHeader')

    while ($true) {
        $value = Read-HostSafely -Prompt "複数ページ結合モードを入力してください (single / sameHeader, Enter で $DefaultValue)"
        if ([string]::IsNullOrWhiteSpace($value)) {
            return $DefaultValue
        }

        switch ($value.Trim()) {
            'single' { return 'single' }
            'sameHeader' { return 'sameHeader' }
            default { Write-Host 'single または sameHeader を入力してください。' }
        }
    }
}

function Read-ExistingPdfPath {
    param([string]$DefaultValue)

    while ($true) {
        $rawValue = Read-OptionalValue -Prompt '下書き作成に使うサンプル PDF のパスを入力してください (Enter で手入力のみ)' -DefaultValue $DefaultValue
        if ([string]::IsNullOrWhiteSpace($rawValue)) {
            return ''
        }

        $resolvedPath = try {
            [System.IO.Path]::GetFullPath($rawValue)
        } catch {
            $rawValue
        }

        if (-not (Test-Path -LiteralPath $resolvedPath)) {
            Write-Host "ファイルが見つかりません: $resolvedPath"
            continue
        }
        if ([System.IO.Path]::GetExtension($resolvedPath) -ine '.pdf') {
            Write-Host 'PDF ファイルを指定してください。'
            continue
        }

        return $resolvedPath
    }
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
        [string]$DestinationAddress = 'A1'
    )

    $worksheet = $null
    $listObject = $null
    $queryTable = $null

    try {
        $worksheet = Get-OrCreateWorksheet -Workbook $Workbook -WorksheetName $WorksheetName
        while ($worksheet.ListObjects.Count -gt 0) {
            $existing = $worksheet.ListObjects.Item(1)
            $existing.Delete()
            $existing | Release-ComObject
        }
        $worksheet.Cells.Clear() | Out-Null

        $source = 'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=' + $QueryName + ';Extended Properties=""'
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
            $queryTable.CommandText = "SELECT * FROM [$QueryName]"
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

function Get-SamplePdfSummaryQueryFormula {
    param([Parameter(Mandatory = $true)][string]$PdfPath)

    $pdfDirectoryLiteral = ConvertTo-MTextLiteral -Value ([System.IO.Path]::GetDirectoryName($PdfPath))
    $pdfFileNameLiteral = ConvertTo-MTextLiteral -Value ([System.IO.Path]::GetFileName($PdfPath))
@"
let
    FolderSource = Folder.Files($pdfDirectoryLiteral),
    TargetFile = Table.SelectRows(FolderSource, each [Name] = $pdfFileNameLiteral),
    PdfBinary = if Table.RowCount(TargetFile) = 0 then error "Target PDF not found." else TargetFile{0}[Content],
    Source = Pdf.Tables(PdfBinary, [Implementation = "1.3", MultiPageTables = false]),
    Indexed = Table.AddIndexColumn(Source, "CandidateIndex", 1, 1, Int64.Type),
    Summary =
        Table.TransformColumns(
            Indexed,
            {
                {"Id", each try Text.From(_) otherwise "", type text},
                {"Kind", each try Text.From(_) otherwise "", type text},
                {"Name", each try Text.From(_) otherwise "", type text}
            }
        )
in
    Table.SelectColumns(Summary, {"CandidateIndex", "Id", "Kind", "Name"})
"@
}

function Get-SamplePdfPreviewQueryFormula {
    param(
        [Parameter(Mandatory = $true)][string]$PdfPath,
        [Parameter(Mandatory = $true)][int]$CandidateIndex
    )

    $pdfDirectoryLiteral = ConvertTo-MTextLiteral -Value ([System.IO.Path]::GetDirectoryName($PdfPath))
    $pdfFileNameLiteral = ConvertTo-MTextLiteral -Value ([System.IO.Path]::GetFileName($PdfPath))
@"
let
    OutputColumns = {"ColumnCount", "RowCount", "RowIndex", "ColumnIndex", "CellText"},
    FolderSource = Folder.Files($pdfDirectoryLiteral),
    TargetFile = Table.SelectRows(FolderSource, each [Name] = $pdfFileNameLiteral),
    PdfBinary = if Table.RowCount(TargetFile) = 0 then error "Target PDF not found." else TargetFile{0}[Content],
    Source = Pdf.Tables(PdfBinary, [Implementation = "1.3", MultiPageTables = false]),
    Indexed = Table.AddIndexColumn(Source, "CandidateIndex", 1, 1, Int64.Type),
    Selected = Table.SelectRows(Indexed, each [CandidateIndex] = $CandidateIndex),
    DataValue = if Table.RowCount(Selected) = 0 then null else try Selected{0}[Data] otherwise null,
    ColumnCount = if DataValue = null then 0 else Table.ColumnCount(DataValue),
    RowCount = if DataValue = null then 0 else Table.RowCount(DataValue),
    PreviewRows = if DataValue = null then {} else Table.ToRows(Table.FirstN(DataValue, 4)),
    PreviewRecords =
        List.Combine(
            List.Transform(
                List.Positions(PreviewRows),
                (rowOffset) =>
                    let
                        currentRow = PreviewRows{rowOffset}
                    in
                        List.Transform(
                            List.Positions(currentRow),
                            (columnOffset) =>
                                [
                                    ColumnCount = ColumnCount,
                                    RowCount = RowCount,
                                    RowIndex = rowOffset + 1,
                                    ColumnIndex = columnOffset + 1,
                                    CellText = try Text.From(currentRow{columnOffset}) otherwise ""
                                ]
                        )
            )
        )
in
    if List.Count(PreviewRecords) = 0 then #table(OutputColumns, {}) else Table.FromRecords(PreviewRecords, OutputColumns)
"@
}

function ConvertTo-ComparableHeaderToken {
    param([AllowNull()]$Value)

    return (Format-ProfileColumnDisplayName -Value $Value).ToUpperInvariant()
}

function Get-CandidateHeaderRowAnalysis {
    param([Parameter(Mandatory = $true)]$Candidate)

    $rowEntries = @($Candidate.Rows.GetEnumerator() | Sort-Object { [int]$_.Key })
    if ($rowEntries.Count -eq 0) {
        return $null
    }

    $searchRows = $rowEntries | Select-Object -First ([Math]::Min(4, $rowEntries.Count))
    $headerKeywords = @('日付', '曜日', '氏名', '名前', '所属', '班', '現場', '工事', '入場', '退場', '開始', '終了', '作業', '内容', '備考', '摘要', '休憩', '小計', '確認', '管理', 'IN', 'OUT')
    $best = $null
    foreach ($entry in $searchRows) {
        $cells = @($entry.Value)
        $displayNames = @(Get-UniqueColumnNameList -Names $cells)
        $normalizedCells = @($cells | ForEach-Object { [string](Format-ProfileColumnDisplayName -Value $_) })
        $meaningfulCells = @($normalizedCells | Where-Object { $_ -ne '(空欄)' })
        $nonEmptyCount = $meaningfulCells.Count
        $headerKeywordCount = @(
            $meaningfulCells |
                Where-Object {
                    $cell = $_.ToUpperInvariant()
                    @($headerKeywords | Where-Object { $cell.Contains($_.ToUpperInvariant()) }).Count -gt 0
                }
        ).Count
        $timeLikeCount = @($meaningfulCells | Where-Object { $_ -match '[:：]' -or $_ -match '時' }).Count
        $digitLikeCount = @($meaningfulCells | Where-Object { $_ -match '\d' }).Count
        $dateLikeCount = @($meaningfulCells | Where-Object { $_ -match '\d{1,4}[/-]\d{1,2}' -or $_ -match '\d+月' -or $_ -match '曜' }).Count
        $score =
            ($headerKeywordCount * 500) +
            ($nonEmptyCount * 80) -
            ($timeLikeCount * 150) -
            ($dateLikeCount * 120) -
            ($digitLikeCount * 30) +
            [int]$entry.Key
        $candidateResult = [pscustomobject]@{
            HeaderRowIndex = [int]$entry.Key
            HeaderNames    = $displayNames
            Signature      = (@($displayNames | ForEach-Object { ConvertTo-ComparableHeaderToken -Value $_ }) -join '|')
            Score          = $score
        }
        if ($null -eq $best -or $candidateResult.Score -gt $best.Score) {
            $best = $candidateResult
        }
    }

    return $best
}

function Get-PreferredTableNameHints {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return @()
    }

    $ignored = @('TABLE', 'PAGE', 'DATA', 'PAGES', '表')
    return @(
        [regex]::Matches($Value, '[一-龠ぁ-んァ-ヶA-Za-z0-9]{2,}') |
            ForEach-Object { $_.Value.Trim() } |
            Where-Object {
                -not [string]::IsNullOrWhiteSpace($_) -and
                ($ignored -notcontains $_.ToUpperInvariant()) -and
                ($_.ToUpperInvariant() -notmatch '^PAGE\d+$')
            } |
            Select-Object -Unique -First 3
    )
}

function Find-ColumnIndexByKeywords {
    param(
        [Parameter(Mandatory = $true)][string[]]$HeaderNames,
        [Parameter(Mandatory = $true)][string[]]$Keywords
    )

    for ($index = 0; $index -lt $HeaderNames.Count; $index += 1) {
        $normalized = ConvertTo-ComparableHeaderToken -Value $HeaderNames[$index]
        foreach ($keyword in $Keywords) {
            if ($normalized.Contains($keyword.ToUpperInvariant())) {
                return ($index + 1)
            }
        }
    }

    return $null
}

function ConvertTo-NormalizedTimeDraft {
    param([Parameter(Mandatory = $true)][string[]]$HeaderNames)

    $entries = @()
    for ($index = 0; $index -lt $HeaderNames.Count; $index += 1) {
        $headerName = [string](Format-ProfileColumnDisplayName -Value $HeaderNames[$index])
        $normalized = [string](ConvertTo-ComparableHeaderToken -Value $headerName)
        $isInColumn =
            $normalized.IndexOf('入場', [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -or
            $normalized.IndexOf('開始', [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -or
            $normalized.IndexOf('IN', [System.StringComparison]::OrdinalIgnoreCase) -ge 0
        $isOutColumn =
            $normalized.IndexOf('退場', [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -or
            $normalized.IndexOf('終了', [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -or
            $normalized.IndexOf('OUT', [System.StringComparison]::OrdinalIgnoreCase) -ge 0
        if ($isInColumn) {
            $entries += [ordered]@{
                    sourceColumn = $index + 1
                    displayName  = '正規化{0}' -f $headerName
                }
        } elseif ($isOutColumn) {
            $entries += [ordered]@{
                    sourceColumn = $index + 1
                    displayName  = '正規化{0}' -f $headerName
                }
        }
    }

    return @($entries)
}

function Get-SamplePdfProfileDraft {
    param([Parameter(Mandatory = $true)][string]$PdfPath)

    $excel = $null
    $workbook = $null
    $analysisSheet = $null
    $tempWorkbookPath = ''
    $templatePath = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\template\PDF2Excel_V2_Converter.xlsm'))

    try {
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            $excel.AskToUpdateLinks = $false
        } catch {
            throw "Excel の起動に失敗しました: $($_.Exception.Message)"
        }
        try {
            if (Test-Path -LiteralPath $templatePath) {
                $tempWorkbookPath = Join-Path ([System.IO.Path]::GetTempPath()) ('pdf2excel_profile_sample_{0}.xlsm' -f ([guid]::NewGuid().ToString('N')))
                Copy-Item -LiteralPath $templatePath -Destination $tempWorkbookPath -Force
                $workbook = $excel.Workbooks.Open($tempWorkbookPath)
            } else {
                $workbook = $excel.Workbooks.Add()
                $tempWorkbookPath = [System.IO.Path]::ChangeExtension([System.IO.Path]::GetTempFileName(), '.xlsx')
                $workbook.SaveAs($tempWorkbookPath, 51)
            }
        } catch {
            throw "サンプル解析用ブックの準備に失敗しました: $($_.Exception.Message)"
        }
        try {
            Add-OrReplaceWorkbookQuery -Workbook $workbook -QueryName 'PDF2Excel_ProfileSampleSummary' -Formula (Get-SamplePdfSummaryQueryFormula -PdfPath $PdfPath)
        } catch {
            throw "サンプル候補一覧クエリの追加に失敗しました: $($_.Exception.Message)"
        }
        try {
            Load-WorkbookQueryToWorksheet -Workbook $workbook -WorksheetName 'ProfileSampleSummary' -QueryName 'PDF2Excel_ProfileSampleSummary' -TableName 'tblProfileSampleSummary'
        } catch {
            throw "サンプル候補一覧の読込に失敗しました: $($_.Exception.Message)"
        }
        $analysisSheet = $workbook.Worksheets.Item('ProfileSampleSummary')
        $usedRange = $analysisSheet.UsedRange
        $rowCount = [int]$usedRange.Rows.Count
        if ($rowCount -lt 2) {
            throw 'サンプル PDF から表候補を取得できませんでした。'
        }

        $summaryCandidates = @()
        for ($row = 2; $row -le $rowCount; $row += 1) {
            $candidateIndex = [int]$analysisSheet.Cells.Item($row, 1).Value2
            if ($candidateIndex -lt 1) {
                continue
            }
            $summaryCandidates += [pscustomobject]@{
                CandidateIndex = $candidateIndex
                Id             = [string]$analysisSheet.Cells.Item($row, 2).Value2
                Kind           = [string]$analysisSheet.Cells.Item($row, 3).Value2
                Name           = [string]$analysisSheet.Cells.Item($row, 4).Value2
            }
        }

        $candidates = @()
        foreach ($candidate in @(
                $summaryCandidates |
                    Sort-Object `
                        @{ Expression = { if ($_.Kind -eq 'Table') { 0 } else { 1 } } }, `
                        @{ Expression = 'CandidateIndex'; Descending = $false }
            )) {
            try {
                Add-OrReplaceWorkbookQuery -Workbook $workbook -QueryName 'PDF2Excel_ProfileSamplePreview' -Formula (Get-SamplePdfPreviewQueryFormula -PdfPath $PdfPath -CandidateIndex $candidate.CandidateIndex)
            } catch {
                throw "サンプル候補プレビュークエリの追加に失敗しました (CandidateIndex=$($candidate.CandidateIndex)): $($_.Exception.Message)"
            }
            try {
                Load-WorkbookQueryToWorksheet -Workbook $workbook -WorksheetName 'ProfileSamplePreview' -QueryName 'PDF2Excel_ProfileSamplePreview' -TableName 'tblProfileSamplePreview'
            } catch {
                throw "サンプル候補プレビューの読込に失敗しました (CandidateIndex=$($candidate.CandidateIndex)): $($_.Exception.Message)"
            }
            $previewSheet = $workbook.Worksheets.Item('ProfileSamplePreview')
            $previewRows = @{}
            $previewUsedRange = $previewSheet.UsedRange
            $previewRowCount = [int]$previewUsedRange.Rows.Count
            $candidateColumnCount = 0
            $candidateRowCount = 0
            for ($previewRow = 2; $previewRow -le $previewRowCount; $previewRow += 1) {
                $candidateColumnCount = [int]$previewSheet.Cells.Item($previewRow, 1).Value2
                $candidateRowCount = [int]$previewSheet.Cells.Item($previewRow, 2).Value2
                $rowIndex = [int]$previewSheet.Cells.Item($previewRow, 3).Value2
                $columnIndex = [int]$previewSheet.Cells.Item($previewRow, 4).Value2
                $cellText = [string]$previewSheet.Cells.Item($previewRow, 5).Value2
                if ($rowIndex -lt 1 -or $columnIndex -lt 1) {
                    continue
                }
                if (-not $previewRows.ContainsKey($rowIndex)) {
                    $previewRows[$rowIndex] = @()
                }
                while ($previewRows[$rowIndex].Count -lt ($columnIndex - 1)) {
                    $previewRows[$rowIndex] += ''
                }
                $previewRows[$rowIndex] += $cellText
            }

            $headerAnalysis = Get-CandidateHeaderRowAnalysis -Candidate ([pscustomobject]@{
                    CandidateIndex = $candidate.CandidateIndex
                    Id             = $candidate.Id
                    Kind           = $candidate.Kind
                    Name           = $candidate.Name
                    ColumnCount    = $candidateColumnCount
                    RowCount       = $candidateRowCount
                    Rows           = $previewRows
                })
            if ($null -eq $headerAnalysis) {
                continue
            }

            $candidates += [pscustomobject]@{
                CandidateIndex = $candidate.CandidateIndex
                Id             = $candidate.Id
                Kind           = $candidate.Kind
                Name           = $candidate.Name
                ColumnCount    = $candidateColumnCount
                RowCount       = $candidateRowCount
                HeaderRowIndex = $headerAnalysis.HeaderRowIndex
                HeaderNames    = @($headerAnalysis.HeaderNames)
                Signature      = $headerAnalysis.Signature
                DataRowCount   = [Math]::Max($candidateRowCount - $headerAnalysis.HeaderRowIndex, 1)
            }
        }

        if ($candidates.Count -eq 0) {
            throw 'サンプル PDF の候補表を解析できませんでした。'
        }

        $bestGroup = $candidates |
            Group-Object -Property { '{0}|{1}' -f $_.Signature, $_.ColumnCount } |
            Sort-Object @{ Expression = { $_.Count }; Descending = $true }, @{ Expression = { ($_.Group | Measure-Object -Property DataRowCount -Maximum).Maximum }; Descending = $true } |
            Select-Object -First 1

        $representative = $bestGroup.Group | Sort-Object @{ Expression = 'DataRowCount'; Descending = $true }, @{ Expression = 'CandidateIndex'; Descending = $false } | Select-Object -First 1
        $targetRowCount = [int][Math]::Max([Math]::Round((($bestGroup.Group | Measure-Object -Property DataRowCount -Average).Average), 0, [MidpointRounding]::AwayFromZero), 1)
        $reviewPersonColumn = Find-ColumnIndexByKeywords -HeaderNames $representative.HeaderNames -Keywords @('氏名', '名前', '作業員')
        $reviewSiteColumn = Find-ColumnIndexByKeywords -HeaderNames $representative.HeaderNames -Keywords @('現場', '工事', '施工場所')
        $reviewInTimeColumn = Find-ColumnIndexByKeywords -HeaderNames $representative.HeaderNames -Keywords @('入場', '開始', 'IN')
        $reviewOutTimeColumn = Find-ColumnIndexByKeywords -HeaderNames $representative.HeaderNames -Keywords @('退場', '終了', 'OUT')

        return [pscustomobject]@{
            SamplePdfPath            = $PdfPath
            ExpectedColumns          = [int]$representative.ColumnCount
            HeaderRowsToSkip         = [int]$representative.HeaderRowIndex
            TargetRowCount           = $targetRowCount
            PreferredTableNameHints  = @(Get-PreferredTableNameHints -Value $representative.Name)
            OutputColumnNames        = @($representative.HeaderNames)
            MultiPageMergeMode       = if ($bestGroup.Count -gt 1) { 'sameHeader' } else { 'single' }
            ReviewPersonColumn       = $reviewPersonColumn
            ReviewSiteColumn         = $reviewSiteColumn
            ReviewInTimeColumn       = $reviewInTimeColumn
            ReviewOutTimeColumn      = $reviewOutTimeColumn
            NormalizedTimeColumns    = @(ConvertTo-NormalizedTimeDraft -HeaderNames $representative.HeaderNames)
        }
    } finally {
        if ($workbook) {
            try {
                $workbook.Close($false)
            } catch {
            }
        }
        foreach ($comObject in @($analysisSheet, $workbook, $excel)) {
            $comObject | Release-ComObject
        }
        if (-not [string]::IsNullOrWhiteSpace($tempWorkbookPath) -and (Test-Path -LiteralPath $tempWorkbookPath)) {
            Remove-Item -LiteralPath $tempWorkbookPath -Force -ErrorAction SilentlyContinue
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Show-ProfileSummary {
    param(
        [Parameter(Mandatory = $true)]$Profile,
        [Parameter(Mandatory = $true)][string]$ResolvedVersionMode,
        [Parameter(Mandatory = $true)][string]$ResolvedOutputPath
    )

    Write-Host ''
    Write-Host '作成内容の確認'
    Write-Host ("  VersionMode               : {0}" -f $ResolvedVersionMode)
    Write-Host ("  ProfileName               : {0}" -f $Profile.name)
    Write-Host ("  DisplayName               : {0}" -f $Profile.displayName)
    Write-Host ("  OutputPath                : {0}" -f $ResolvedOutputPath)
    Write-Host ("  Description               : {0}" -f $Profile.description)
    Write-Host ("  ExpectedColumns           : {0}" -f $Profile.expectedColumns)
    Write-Host ("  HeaderRowsToSkip          : {0}" -f $Profile.headerRowsToSkip)
    Write-Host ("  TargetRowCount            : {0}" -f $Profile.targetRowCount)
    Write-Host ("  AllowMoreColumns          : {0}" -f $Profile.allowMoreColumns)
    Write-Host ("  SourceFileColumnName      : {0}" -f $Profile.sourceFileColumnName)
    Write-Host ("  DataColumnPrefix          : {0}" -f $Profile.dataColumnPrefix)
    $outputColumnSummary = if (@($Profile.outputColumnNames).Count -eq 0) { '未設定' } else { (@($Profile.outputColumnNames) -join ', ') }
    Write-Host ("  OutputColumnNames         : {0}" -f $outputColumnSummary)
    Write-Host ("  PreferredTableNameContains: {0}" -f ((@($Profile.preferredTableNameContains) -join ', ')))
    if ($ResolvedVersionMode -eq 'v2') {
        Write-Host ("  MultiPageMergeMode        : {0}" -f $Profile.multiPageMergeMode)
        Write-Host ("  ReviewPersonColumn        : {0}" -f $Profile.reviewPersonColumn)
        Write-Host ("  ReviewSiteColumn          : {0}" -f $Profile.reviewSiteColumn)
        Write-Host ("  ReviewInTimeColumn        : {0}" -f $Profile.reviewInTimeColumn)
        Write-Host ("  ReviewOutTimeColumn       : {0}" -f $Profile.reviewOutTimeColumn)
        $normalizedTimeSummary = if (@($Profile.normalizedTimeColumns).Count -eq 0) { '未設定' } else { (@($Profile.normalizedTimeColumns) | ForEach-Object { '{0}:{1}' -f $_.sourceColumn, $_.displayName }) -join ', ' }
        Write-Host ("  NormalizedTimeColumns     : {0}" -f $normalizedTimeSummary)
    }
    Write-Host ''
}

function Confirm-WizardSummary {
    while ($true) {
        $value = Read-HostSafely -Prompt 'この内容で作成しますか? (Y/N)'
        switch ($value.Trim().ToUpperInvariant()) {
            'Y' { return $true }
            'YES' { return $true }
            'N' { return $false }
            'NO' { return $false }
            default { Write-Host 'Y または N を入力してください。' }
        }
    }
}

function New-ProfileScaffoldObject {
    param(
        [Parameter(Mandatory = $true)][string]$ResolvedVersionMode,
        [Parameter(Mandatory = $true)][string]$ResolvedProfileName,
        [Parameter(Mandatory = $true)][string]$ResolvedDisplayName,
        [string]$Description,
        [int]$ExpectedColumns = 0,
        [int]$HeaderRowsToSkip = 0,
        [int]$TargetRowCount = 0,
        [bool]$AllowMoreColumns = $false,
        [string[]]$PreferredTableNameContains = @(),
        [string]$SourceFileColumnName,
        [string]$DataColumnPrefix,
        [string[]]$OutputColumnNames = @(),
        [string]$MultiPageMergeMode,
        $NormalizedTimeColumns = @(),
        [Nullable[int]]$ReviewPersonColumn = $null,
        [Nullable[int]]$ReviewSiteColumn = $null,
        [Nullable[int]]$ReviewInTimeColumn = $null,
        [Nullable[int]]$ReviewOutTimeColumn = $null
    )

    $defaults = Get-DefaultProfileScaffoldValues -ResolvedVersionMode $ResolvedVersionMode -ResolvedProfileName $ResolvedProfileName -ResolvedDisplayName $ResolvedDisplayName
    $resolvedExpectedColumns = if ($ExpectedColumns -gt 0) { $ExpectedColumns } else { $defaults.ExpectedColumns }
    $resolvedDataColumnPrefix = if ([string]::IsNullOrWhiteSpace($DataColumnPrefix)) { $defaults.DataColumnPrefix } else { $DataColumnPrefix }
    $resolvedOutputColumnNames = @(
        Get-ProfileDataColumnNames -Profile ([pscustomobject]@{
                ExpectedColumns  = $resolvedExpectedColumns
                DataColumnPrefix = $resolvedDataColumnPrefix
                OutputColumnNames = if (@($OutputColumnNames).Count -gt 0) { @($OutputColumnNames) } else { @($defaults.OutputColumnNames) }
            })
    )

    $profile = [ordered]@{
        name                       = $ResolvedProfileName
        displayName                = $ResolvedDisplayName
        description                = if ([string]::IsNullOrWhiteSpace($Description)) { $defaults.Description } else { $Description }
        expectedColumns            = $resolvedExpectedColumns
        headerRowsToSkip           = $HeaderRowsToSkip
        targetRowCount             = if ($TargetRowCount -gt 0) { $TargetRowCount } else { $defaults.TargetRowCount }
        allowMoreColumns           = $AllowMoreColumns
        preferredTableKinds        = @('Table')
        preferredTableNameContains = @($PreferredTableNameContains)
        preferredTableIdContains   = @()
        sourceFileColumnName       = if ([string]::IsNullOrWhiteSpace($SourceFileColumnName)) { $defaults.SourceFileColumnName } else { $SourceFileColumnName }
        dataColumnPrefix           = $resolvedDataColumnPrefix
        outputColumnNames          = @($resolvedOutputColumnNames)
    }

    if ($ResolvedVersionMode -eq 'v2') {
        $profile.multiPageMergeMode = if ([string]::IsNullOrWhiteSpace($MultiPageMergeMode)) { $defaults.MultiPageMergeMode } else { $MultiPageMergeMode }
        $profile.normalizedTimeColumns = @($NormalizedTimeColumns)
        $profile.reviewPersonColumn = $ReviewPersonColumn
        $profile.reviewSiteColumn = $ReviewSiteColumn
        $profile.reviewInTimeColumn = $ReviewInTimeColumn
        $profile.reviewOutTimeColumn = $ReviewOutTimeColumn
    }

    return $profile
}

function Invoke-V2Wizard {
    param(
        [string]$RequestedProfileName,
        [string]$RequestedDisplayName,
        [string]$RequestedOutputPath,
        [string]$RequestedSamplePdfPath
    )

    $resolvedProfileName = Read-ProfileName -DefaultValue $RequestedProfileName
    $defaults = Get-DefaultProfileScaffoldValues -ResolvedVersionMode 'v2' -ResolvedProfileName $resolvedProfileName -ResolvedDisplayName $RequestedDisplayName
    $resolvedSamplePdfPath = Read-ExistingPdfPath -DefaultValue $RequestedSamplePdfPath
    if (-not [string]::IsNullOrWhiteSpace($resolvedSamplePdfPath)) {
        Write-Host 'サンプル PDF から下書きを解析しています...'
        try {
            $sampleDraft = Get-SamplePdfProfileDraft -PdfPath $resolvedSamplePdfPath
            $defaults.ExpectedColumns = $sampleDraft.ExpectedColumns
            $defaults.HeaderRowsToSkip = $sampleDraft.HeaderRowsToSkip
            $defaults.TargetRowCount = $sampleDraft.TargetRowCount
            $defaults.PreferredTableNameContains = @($sampleDraft.PreferredTableNameHints)
            $defaults.OutputColumnNames = @($sampleDraft.OutputColumnNames)
            $defaults.MultiPageMergeMode = $sampleDraft.MultiPageMergeMode
            $defaults.ReviewPersonColumn = $sampleDraft.ReviewPersonColumn
            $defaults.ReviewSiteColumn = $sampleDraft.ReviewSiteColumn
            $defaults.ReviewInTimeColumn = $sampleDraft.ReviewInTimeColumn
            $defaults.ReviewOutTimeColumn = $sampleDraft.ReviewOutTimeColumn
            $defaults.NormalizedTimeColumns = @($sampleDraft.NormalizedTimeColumns)
            Write-Host ("解析結果を初期値に反映しました: 列数={0}, ヘッダー行={1}, 結合モード={2}" -f $sampleDraft.ExpectedColumns, $sampleDraft.HeaderRowsToSkip, $sampleDraft.MultiPageMergeMode)
        } catch {
            Write-Host ("サンプル PDF の解析に失敗したため、手入力を続けます: {0}" -f $_.Exception.Message)
        }
    }
    $resolvedDisplayName = Read-OptionalValue -Prompt '表示名を入力してください' -DefaultValue $defaults.DisplayName
    $defaultOutputPath = if ([string]::IsNullOrWhiteSpace($RequestedOutputPath)) {
        Join-Path $profileBaseDir ("{0}.json" -f $resolvedProfileName)
    } else {
        $RequestedOutputPath
    }
    $resolvedOutputPath = [System.IO.Path]::GetFullPath((Read-OptionalValue -Prompt '保存先を入力してください' -DefaultValue $defaultOutputPath))
    $resolvedDescription = Read-OptionalValue -Prompt '説明を入力してください' -DefaultValue $defaults.Description
    $expectedColumns = Read-ValidatedIntValue -Prompt '想定列数を入力してください' -DefaultValue $defaults.ExpectedColumns -MinValue 1
    $headerRowsToSkip = Read-ValidatedIntValue -Prompt 'ヘッダー除外行数を入力してください' -DefaultValue $defaults.HeaderRowsToSkip -MinValue 0
    $targetRowCount = Read-ValidatedIntValue -Prompt '想定行数を入力してください' -DefaultValue $defaults.TargetRowCount -MinValue 1
    $allowMoreColumns = Read-YesNoValue -Prompt '想定列数より多い列を許可しますか' -DefaultValue $defaults.AllowMoreColumns
    $multiPageMergeMode = Read-MultiPageMergeModeValue -DefaultValue $defaults.MultiPageMergeMode
    $sourceFileColumnName = Read-OptionalValue -Prompt '元ファイル列名を入力してください' -DefaultValue $defaults.SourceFileColumnName
    $dataColumnPrefix = Read-OptionalValue -Prompt 'データ列接頭辞を入力してください' -DefaultValue $defaults.DataColumnPrefix
    $preferredTableNameDefault = if (@($defaults.PreferredTableNameContains).Count -eq 0) { '' } else { @($defaults.PreferredTableNameContains) -join ',' }
    $preferredTableNameContains = ConvertTo-StringArrayFromCommaSeparated -Value (Read-OptionalValue -Prompt '優先表名キーワードを入力してください (カンマ区切り / 未設定は Enter)' -DefaultValue $preferredTableNameDefault)
    $reviewPersonColumn = Read-NullableIntValue -Prompt 'Review 氏名列番号を入力してください' -DefaultValue $defaults.ReviewPersonColumn -MinValue 1
    $reviewSiteColumn = Read-NullableIntValue -Prompt 'Review 現場列番号を入力してください' -DefaultValue $defaults.ReviewSiteColumn -MinValue 1
    $reviewInTimeColumn = Read-NullableIntValue -Prompt 'Review 入場列番号を入力してください' -DefaultValue $defaults.ReviewInTimeColumn -MinValue 1
    $reviewOutTimeColumn = Read-NullableIntValue -Prompt 'Review 退場列番号を入力してください' -DefaultValue $defaults.ReviewOutTimeColumn -MinValue 1
    $normalizedTimeColumns = @(Read-NormalizedTimeColumns -DefaultEntries $defaults.NormalizedTimeColumns)

    $resolvedProfile = New-ProfileScaffoldObject `
        -ResolvedVersionMode 'v2' `
        -ResolvedProfileName $resolvedProfileName `
        -ResolvedDisplayName $resolvedDisplayName `
        -Description $resolvedDescription `
        -ExpectedColumns $expectedColumns `
        -HeaderRowsToSkip $headerRowsToSkip `
        -TargetRowCount $targetRowCount `
        -AllowMoreColumns:$allowMoreColumns `
        -PreferredTableNameContains $preferredTableNameContains `
        -SourceFileColumnName $sourceFileColumnName `
        -DataColumnPrefix $dataColumnPrefix `
        -OutputColumnNames $defaults.OutputColumnNames `
        -MultiPageMergeMode $multiPageMergeMode `
        -NormalizedTimeColumns $normalizedTimeColumns `
        -ReviewPersonColumn $reviewPersonColumn `
        -ReviewSiteColumn $reviewSiteColumn `
        -ReviewInTimeColumn $reviewInTimeColumn `
        -ReviewOutTimeColumn $reviewOutTimeColumn

    Show-ProfileSummary -Profile $resolvedProfile -ResolvedVersionMode 'v2' -ResolvedOutputPath $resolvedOutputPath
    if (-not (Confirm-WizardSummary)) {
        throw 'プロファイル作成をキャンセルしました。'
    }

    return [pscustomobject]@{
        ProfileObject = $resolvedProfile
        OutputPath    = $resolvedOutputPath
    }
}

if ($SkipMain) {
    return
}

if ([string]::IsNullOrWhiteSpace($VersionMode)) {
    $VersionMode = Read-VersionMode
}

if ($Wizard -and $VersionMode -eq 'v2') {
    $wizardResult = Invoke-V2Wizard -RequestedProfileName $ProfileName -RequestedDisplayName $DisplayName -RequestedOutputPath $OutputPath -RequestedSamplePdfPath $SamplePdfPath
    $ProfileName = [string]$wizardResult.ProfileObject.name
    $DisplayName = [string]$wizardResult.ProfileObject.displayName
    $OutputPath = [string]$wizardResult.OutputPath
    $profileObject = $wizardResult.ProfileObject
} else {
    $ProfileName = Read-ProfileName -DefaultValue $ProfileName
    $defaultDisplayName = if ($VersionMode -eq 'v2') { '生データ転記サンプルプロファイル' } else { "$ProfileName プロファイル" }
    if ([string]::IsNullOrWhiteSpace($DisplayName)) {
        $DisplayName = Read-OptionalValue -Prompt '表示名を入力してください' -DefaultValue $defaultDisplayName
    }

    $defaultOutputPath = Join-Path $profileBaseDir ("{0}.json" -f $ProfileName)
    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        $OutputPath = Read-OptionalValue -Prompt '保存先を入力してください' -DefaultValue $defaultOutputPath
    }
    $OutputPath = [System.IO.Path]::GetFullPath($OutputPath)
}

$outputDir = Split-Path -Path $OutputPath -Parent
Ensure-Directory -Path $outputDir

if ((Test-Path -LiteralPath $OutputPath) -and -not $Force) {
    throw "プロファイルは既に存在します。上書きする場合は -Force を指定してください: $OutputPath"
}

if ($null -eq $profileObject) {
    $profileObject = New-ProfileScaffoldObject -ResolvedVersionMode $VersionMode -ResolvedProfileName $ProfileName -ResolvedDisplayName $DisplayName
}

$json = $profileObject | ConvertTo-Json -Depth 6
[System.IO.File]::WriteAllText($OutputPath, $json + [Environment]::NewLine, [System.Text.UTF8Encoding]::new($true))

Write-Host ''
Write-Host ("{0} のプロファイル雛形を作成しました。" -f (Get-VersionDisplayName -VersionMode $VersionMode))
Write-Host ("内部ID: {0}" -f $ProfileName)
Write-Host ("表示名: {0}" -f $DisplayName)
Write-Host ("保存先: {0}" -f $OutputPath)
Write-Host ''
Write-Host '次に確認してください:'
Write-Host '- expectedColumns / headerRowsToSkip / targetRowCount'
if ($VersionMode -eq 'v2') {
    Write-Host '- multiPageMergeMode'
    Write-Host '- outputColumnNames'
    Write-Host '- preferredTableNameContains'
    Write-Host '- normalizedTimeColumns'
    Write-Host '- reviewPersonColumn / reviewSiteColumn / reviewInTimeColumn / reviewOutTimeColumn'
}
