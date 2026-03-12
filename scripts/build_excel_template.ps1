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

function Ensure-WorkbookHasThreeSheets {
    param([Parameter(Mandatory = $true)]$Workbook)

    while ($Workbook.Worksheets.Count -lt 3) {
        $null = $Workbook.Worksheets.Add()
    }

    while ($Workbook.Worksheets.Count -gt 3) {
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
    $Worksheet.Range('A11').Value2 = 'かんたんな使い方'
    $Worksheet.Range('B11').Value2 = '1. run_pdf2excel.bat を実行  2. PDF を選択  3. Result と Errors を確認'
    $Worksheet.Range('A12').Value2 = '確認ポイント'
    $Worksheet.Range('B12').Value2 = 'Result は変換成功データ、Errors は失敗した PDF と理由です。'
    $Worksheet.Range('A13').Value2 = '注意'
    $Worksheet.Range('B13').Value2 = '文字を選択できるテキスト PDF と、ほぼ同じレイアウトの帳票を想定しています。'
    $Worksheet.Range('A1:B1').Font.Bold = $true
    $Worksheet.Range('A1:B1').Interior.Color = 15773696
    $Worksheet.Range('A11:A13').Font.Bold = $true
    $Worksheet.Columns.Item('A').ColumnWidth = 18
    $Worksheet.Columns.Item('B').ColumnWidth = 92
    $Worksheet.Range('A1:B13').VerticalAlignment = -4160
    $Worksheet.Range('A1:B13').WrapText = $true
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
        $null = $components.Import($ModulePath)
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

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Add()
    Ensure-WorkbookHasThreeSheets -Workbook $workbook

    $sheet1 = $workbook.Worksheets.Item(1)
    $sheet2 = $workbook.Worksheets.Item(2)
    $sheet3 = $workbook.Worksheets.Item(3)

    Set-ControlSheetLayout -Worksheet $sheet1
    Set-DataSheetLayout -Worksheet $sheet2 -SheetName 'Result' -Description '変換成功データがここに読み込まれます。A列はPDFファイル名、B列から30列分の表データです。'
    Set-DataSheetLayout -Worksheet $sheet3 -SheetName 'Errors' -Description '失敗したPDFと理由がここに一覧表示されます。'

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

    $sheet3 | Release-ComObject
    $sheet2 | Release-ComObject
    $sheet1 | Release-ComObject
    $workbook | Release-ComObject
    $excel | Release-ComObject
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}


