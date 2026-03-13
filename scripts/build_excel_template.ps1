param(
    [string]$TemplatePath = (Join-Path (Join-Path $PSScriptRoot '..') 'template\PDF2Excel_Converter.xlsm'),
    [string]$VbaModulePath = (Join-Path (Join-Path $PSScriptRoot '..') 'template\vba\PDF2ExcelMacros.bas')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

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

function Ensure-WorkbookHasRequiredSheets {
    param([Parameter(Mandatory = $true)]$Workbook)

    while ($Workbook.Worksheets.Count -lt 4) {
        $null = $Workbook.Worksheets.Add()
    }

    while ($Workbook.Worksheets.Count -gt 4) {
        $Workbook.Worksheets.Item($Workbook.Worksheets.Count).Delete()
    }
}

function Set-ControlSheetLayout {
    param([Parameter(Mandatory = $true)]$Worksheet)

    $Worksheet.Cells.Clear() | Out-Null
    $Worksheet.Name = 'Control'
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
    $Worksheet.Range('A16').Value2 = 'かんたんな使い方'
    $Worksheet.Range('B16').Value2 = '1. run_pdf2excel.bat を実行  2. PDF を選択  3. 実行前チェックを確認  4. Result / Summary / Errors を確認'
    $Worksheet.Range('A17').Value2 = '確認ポイント'
    $Worksheet.Range('B17').Value2 = 'Summary は件数の全体像、Result は変換成功データ、Errors は失敗した PDF と理由です。'
    $Worksheet.Range('A18').Value2 = '注意'
    $Worksheet.Range('B18').Value2 = '文字を選択できるテキスト PDF と、ほぼ同じレイアウトの帳票を想定しています。'
    $Worksheet.Range('A1:B1').Font.Bold = $true
    $Worksheet.Range('A1:B1').Interior.Color = 15773696
    $Worksheet.Range('A16:A18').Font.Bold = $true
    $Worksheet.Columns.Item('A').ColumnWidth = 18
    $Worksheet.Columns.Item('B').ColumnWidth = 92
    $Worksheet.Range('A1:B18').VerticalAlignment = -4160
    $Worksheet.Range('A1:B18').WrapText = $true
    $Worksheet.Range('B6').Value2 = '待機中'
    $Worksheet.Application.ActiveWindow.SplitRow = 1
    $Worksheet.Application.ActiveWindow.FreezePanes = $true
}

function Set-DataSheetLayout {
    param(
        [Parameter(Mandatory = $true)]$Worksheet,
        [Parameter(Mandatory = $true)][string]$SheetName,
        [Parameter(Mandatory = $true)][string]$Description
    )

    $Worksheet.Cells.Clear() | Out-Null
    $Worksheet.Name = $SheetName
    $Worksheet.Range('A1').Value2 = $SheetName
    $Worksheet.Range('A2').Value2 = $Description
    $Worksheet.Range('A1').Font.Bold = $true
    $Worksheet.Range('A1').Font.Size = 14
    $Worksheet.Columns.Item('A').ColumnWidth = 24
    $Worksheet.Columns.Item('B').ColumnWidth = 24
}

function Set-SummarySheetLayout {
    param([Parameter(Mandatory = $true)]$Worksheet)

    $Worksheet.Cells.Clear() | Out-Null
    $Worksheet.Name = 'Summary'
    $Worksheet.Range('A1').Value2 = 'Summary'
    $Worksheet.Range('A2').Value2 = '実行結果の集計と、PDFごとの内訳を表示します。'
    $Worksheet.Range('A4').Value2 = '項目'
    $Worksheet.Range('B4').Value2 = '内容'
    $Worksheet.Range('A7').Value2 = 'PDF別サマリー'
    $Worksheet.Range('M7').Value2 = 'エラー分類別件数'
    $Worksheet.Range('A1').Font.Bold = $true
    $Worksheet.Range('A1').Font.Size = 14
    $Worksheet.Range('A4:B4').Font.Bold = $true
    $Worksheet.Range('A4:B4').Interior.Color = 15773696
    $Worksheet.Range('A7').Font.Bold = $true
    $Worksheet.Range('M7').Font.Bold = $true
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

function Import-VbaModule {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)][string]$ModulePath
    )

    if (-not (Test-Path -LiteralPath $ModulePath)) {
        Write-Warning "VBA module file was not found: $ModulePath"
        return
    }

    try {
        $components = $Workbook.VBProject.VBComponents
        $moduleCode = Get-Content -LiteralPath $ModulePath -Raw -Encoding UTF8
        $moduleCode = $moduleCode -replace '^\s*Attribute VB_Name = ".*?"\r?\n', ''

        for ($index = $components.Count; $index -ge 1; $index -= 1) {
            $component = $components.Item($index)
            try {
                if ([string]$component.Name -eq 'PDF2ExcelMacros') {
                    $components.Remove($component)
                }
            } finally {
                $component | Release-ComObject
            }
        }

        $moduleComponent = $components.Add(1)
        $moduleComponent.Name = 'PDF2ExcelMacros'
        $moduleComponent.CodeModule.AddFromString($moduleCode)
        $moduleComponent | Release-ComObject
        $components | Release-ComObject
    } catch {
        Write-Warning "VBA module import failed: $($_.Exception.Message)"
    }
}

Ensure-Directory -Path (Split-Path -Path $TemplatePath -Parent)

$excel = $null
$workbook = $null
$sheet1 = $null
$sheet2 = $null
$sheet3 = $null
$sheet4 = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Add()
    Ensure-WorkbookHasRequiredSheets -Workbook $workbook

    $sheet1 = $workbook.Worksheets.Item(1)
    $sheet2 = $workbook.Worksheets.Item(2)
    $sheet3 = $workbook.Worksheets.Item(3)
    $sheet4 = $workbook.Worksheets.Item(4)

    Set-ControlSheetLayout -Worksheet $sheet1
    Set-DataSheetLayout -Worksheet $sheet2 -SheetName 'Result' -Description '変換成功データがここに読み込まれます。A列はPDFファイル名、B列以降はプロファイルに応じた表データです。'
    Set-DataSheetLayout -Worksheet $sheet3 -SheetName 'Errors' -Description '失敗したPDFと、分類されたエラー理由がここに一覧表示されます。'
    Set-SummarySheetLayout -Worksheet $sheet4

    $workbook.SaveAs($TemplatePath, 52)
    Import-VbaModule -Workbook $workbook -ModulePath $VbaModulePath
    $workbook.Save()

    Write-Host "Template created: $TemplatePath"
} finally {
    if ($workbook) {
        try {
            $workbook.Close($true)
        } catch {
        }
    }
    if ($excel) {
        try {
            $excel.Quit()
        } catch {
        }
    }

    $sheet4 | Release-ComObject
    $sheet3 | Release-ComObject
    $sheet2 | Release-ComObject
    $sheet1 | Release-ComObject
    $workbook | Release-ComObject
    $excel | Release-ComObject
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}


