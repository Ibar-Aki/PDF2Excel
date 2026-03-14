Set-StrictMode -Version Latest

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

function Reset-Directory {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (Test-Path -LiteralPath $Path) {
        try {
            Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
        } catch [System.IO.DirectoryNotFoundException] {
        }
    }

    New-Item -ItemType Directory -Path $Path -Force | Out-Null
}

function New-EmptyFile {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType File -Path $Path | Out-Null
    }
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

function Get-ProfileOutputColumnNames {
    param([Parameter(Mandatory = $true)]$Profile)

    $names = @($Profile.SourceFileColumnName)
    foreach ($index in 1..$Profile.ExpectedColumns) {
        $names += '{0}{1}' -f $Profile.DataColumnPrefix, $index
    }

    return $names
}

function Get-ControlSheetStaticCells {
    return [ordered]@{
        'A1'  = '項目'
        'B1'  = '内容'
        'A2'  = '入力フォルダ'
        'A3'  = '出力ファイル'
        'A4'  = 'ログファイル'
        'A5'  = '最終実行日時'
        'A6'  = '状態'
        'A7'  = '対象PDF数'
        'A8'  = '取込データ行数'
        'A9'  = 'エラー件数'
        'A10' = '成功PDF数'
        'A11' = '失敗PDF数'
        'A12' = '処理時間(秒)'
        'A13' = '使用プロファイル'
        'A14' = 'プロファイル説明'
        'A16' = 'かんたんな使い方'
        'B16' = '1. run_pdf2excel.bat を実行  2. PDF を選択  3. 実行前チェックを確認  4. Result / Summary / Errors を確認'
        'A17' = '確認ポイント'
        'B17' = 'Summary は件数の全体像、Result は変換成功データ、Errors は失敗した PDF と理由です。'
        'A18' = '注意'
        'B18' = '文字を選択できるテキスト PDF と、ほぼ同じレイアウトの帳票を想定しています。'
    }
}

function Set-ControlSheetStaticCells {
    param([Parameter(Mandatory = $true)]$Worksheet)

    foreach ($entry in (Get-ControlSheetStaticCells).GetEnumerator()) {
        $Worksheet.Range($entry.Key).Value2 = $entry.Value
    }
}

function Get-SummarySheetStaticCells {
    return [ordered]@{
        'A1'  = 'Summary'
        'A2'  = '実行結果の集計と、PDFごとの内訳を表示します。'
        'A4'  = '項目'
        'B4'  = '内容'
        'A13' = 'PDF別サマリー'
        'M13' = 'エラー分類別件数'
    }
}

function Set-SummarySheetStaticCells {
    param([Parameter(Mandatory = $true)]$Worksheet)

    foreach ($entry in (Get-SummarySheetStaticCells).GetEnumerator()) {
        $Worksheet.Range($entry.Key).Value2 = $entry.Value
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
