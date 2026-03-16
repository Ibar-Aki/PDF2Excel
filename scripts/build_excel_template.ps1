param(
    [string]$TemplatePath = (Join-Path (Join-Path $PSScriptRoot '..') 'template\PDF2Excel_V1_Converter.xlsm'),
    [string]$VbaModulePath = (Join-Path (Join-Path $PSScriptRoot '..') 'template\vba\PDF2ExcelMacros.bas'),
    [string]$VbaModuleShiftJisPath = (Join-Path (Join-Path $PSScriptRoot '..') 'template\vba\PDF2ExcelMacros.sjis.bas'),
    [string]$TemplateBuilderModulePath = (Join-Path (Join-Path $PSScriptRoot '..') 'template\vba\PDF2ExcelTemplateBuilder.bas'),
    [string]$TemplateBuilderModuleShiftJisPath = (Join-Path (Join-Path $PSScriptRoot '..') 'template\vba\PDF2ExcelTemplateBuilder.sjis.bas'),
    [ValidateSet('v1', 'v2')]
    [string]$TemplateVariant
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'pdf2excel.common.ps1')

function Sync-VbaModuleEncodingMirror {
    param(
        [Parameter(Mandatory = $true)][string]$Utf8Path,
        [Parameter(Mandatory = $true)][string]$ShiftJisPath
    )

    if (-not (Test-Path -LiteralPath $Utf8Path)) {
        throw "VBA モジュールファイルが見つかりません: $Utf8Path"
    }

    $moduleText = Get-Content -LiteralPath $Utf8Path -Raw -Encoding UTF8
    $encoding = [System.Text.Encoding]::GetEncoding(932)
    [System.IO.File]::WriteAllBytes($ShiftJisPath, $encoding.GetBytes($moduleText))
}

function Resolve-TemplateVariant {
    param(
        [string]$TemplatePath,
        [string]$ExplicitVariant
    )

    if (-not [string]::IsNullOrWhiteSpace($ExplicitVariant)) {
        return $ExplicitVariant.ToLowerInvariant()
    }

    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($TemplatePath).ToLowerInvariant()
    if ($fileName.Contains('_v2_') -or $fileName.EndsWith('_v2') -or $fileName.Contains('v2')) {
        return 'v2'
    }

    return 'v1'
}

function Ensure-WorkbookHasRequiredSheets {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [ValidateSet('v1', 'v2')]
        [Parameter(Mandatory = $true)][string]$TemplateVariant
    )

    $requiredSheetCount = if ($TemplateVariant -eq 'v2') { 5 } else { 4 }

    while ($Workbook.Worksheets.Count -lt $requiredSheetCount) {
        $null = $Workbook.Worksheets.Add()
    }

    while ($Workbook.Worksheets.Count -gt $requiredSheetCount) {
        $Workbook.Worksheets.Item($Workbook.Worksheets.Count).Delete()
    }
}

function Set-ControlSheetLayout {
    param(
        [Parameter(Mandatory = $true)]$Worksheet,
        [ValidateSet('v1', 'v2')]
        [Parameter(Mandatory = $true)][string]$TemplateVariant
    )

    $Worksheet.Cells.Clear() | Out-Null
    $Worksheet.Name = 'Control'
    Set-ControlSheetStaticCells -Worksheet $Worksheet -VersionMode $TemplateVariant
    $Worksheet.Range('A1:B1').Font.Bold = $true
    $Worksheet.Range('A1:B1').Interior.Color = 15773696
    $Worksheet.Range('A15:A18').Font.Bold = $true
    $Worksheet.Columns.Item('A').ColumnWidth = 18
    $Worksheet.Columns.Item('B').ColumnWidth = 92
    $Worksheet.Range('A1:B18').VerticalAlignment = -4160
    $Worksheet.Range('A1:B18').WrapText = $true
    $Worksheet.Range('B6').Value2 = '待機中'
    $Worksheet.Range('B15').Value2 = Get-VersionDisplayName -VersionMode $TemplateVariant
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
    param(
        [Parameter(Mandatory = $true)]$Worksheet,
        [ValidateSet('v1', 'v2')]
        [Parameter(Mandatory = $true)][string]$TemplateVariant
    )

    $Worksheet.Cells.Clear() | Out-Null
    $Worksheet.Name = 'Summary'
    Set-SummarySheetStaticCells -Worksheet $Worksheet -VersionMode $TemplateVariant
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

function Set-ReviewSheetLayout {
    param([Parameter(Mandatory = $true)]$Worksheet)

    $Worksheet.Cells.Clear() | Out-Null
    $Worksheet.Name = 'Review'
    $Worksheet.Range('A1').Value2 = 'Review'
    $Worksheet.Range('A2').Value2 = '確認が必要な生データ転記行を一覧化します。'
    $Worksheet.Range('A4').Value2 = '元ファイル名'
    $Worksheet.Range('B4').Value2 = 'ページ'
    $Worksheet.Range('C4').Value2 = '氏名 raw'
    $Worksheet.Range('D4').Value2 = '現場 raw'
    $Worksheet.Range('E4').Value2 = '入場 raw'
    $Worksheet.Range('F4').Value2 = '退場 raw'
    $Worksheet.Range('G4').Value2 = '確認要理由'
    $Worksheet.Range('A1').Font.Bold = $true
    $Worksheet.Range('A1').Font.Size = 14
    $Worksheet.Range('A4:G4').Font.Bold = $true
    $Worksheet.Range('A4:G4').Interior.Color = 15773696
    $Worksheet.Columns.Item('A').ColumnWidth = 28
    $Worksheet.Columns.Item('B').ColumnWidth = 10
    $Worksheet.Columns.Item('C').ColumnWidth = 18
    $Worksheet.Columns.Item('D').ColumnWidth = 30
    $Worksheet.Columns.Item('E').ColumnWidth = 18
    $Worksheet.Columns.Item('F').ColumnWidth = 18
    $Worksheet.Columns.Item('G').ColumnWidth = 44
}

function Import-VbaModule {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)][string]$ModulePath
    )

    if (-not (Test-Path -LiteralPath $ModulePath)) {
        Write-Warning "VBA モジュールファイルが見つかりません: $ModulePath"
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
        Write-Warning "VBA モジュールの取込に失敗しました: $($_.Exception.Message)"
    }
}

Ensure-Directory -Path (Split-Path -Path $TemplatePath -Parent)
Ensure-Directory -Path (Split-Path -Path $VbaModuleShiftJisPath -Parent)
Ensure-Directory -Path (Split-Path -Path $TemplateBuilderModuleShiftJisPath -Parent)
Sync-VbaModuleEncodingMirror -Utf8Path $VbaModulePath -ShiftJisPath $VbaModuleShiftJisPath
Sync-VbaModuleEncodingMirror -Utf8Path $TemplateBuilderModulePath -ShiftJisPath $TemplateBuilderModuleShiftJisPath

$excel = $null
$workbook = $null
$sheet1 = $null
$sheet2 = $null
$sheet3 = $null
$sheet4 = $null
$sheet5 = $null
$resolvedTemplateVariant = Resolve-TemplateVariant -TemplatePath $TemplatePath -ExplicitVariant $TemplateVariant

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Add()
    Ensure-WorkbookHasRequiredSheets -Workbook $workbook -TemplateVariant $resolvedTemplateVariant

    $sheet1 = $workbook.Worksheets.Item(1)
    $sheet2 = $workbook.Worksheets.Item(2)
    $sheet3 = $workbook.Worksheets.Item(3)
    $sheet4 = $workbook.Worksheets.Item(4)
    if ($resolvedTemplateVariant -eq 'v2') {
        $sheet5 = $workbook.Worksheets.Item(5)
    }

    Set-ControlSheetLayout -Worksheet $sheet1 -TemplateVariant $resolvedTemplateVariant
    Set-DataSheetLayout -Worksheet $sheet2 -SheetName 'Result' -Description '変換成功データがここに読み込まれます。A列はPDFファイル名、B列以降はプロファイルに応じた表データです。'
    Set-DataSheetLayout -Worksheet $sheet3 -SheetName 'Errors' -Description '失敗したPDFと、分類されたエラー理由がここに一覧表示されます。'
    Set-SummarySheetLayout -Worksheet $sheet4 -TemplateVariant $resolvedTemplateVariant
    if ($resolvedTemplateVariant -eq 'v2') {
        Set-ReviewSheetLayout -Worksheet $sheet5
    }

    $workbook.SaveAs($TemplatePath, 52)
    Import-VbaModule -Workbook $workbook -ModulePath $VbaModulePath
    $workbook.Save()

    Write-Host "テンプレートを作成しました: $TemplatePath"
    Write-Host "テンプレート種別: $resolvedTemplateVariant"
    Write-Host "Shift_JIS VBA ミラー: $VbaModuleShiftJisPath"
    Write-Host "Shift_JIS ブートストラップ VBA ミラー: $TemplateBuilderModuleShiftJisPath"
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

    $sheet5 | Release-ComObject
    $sheet4 | Release-ComObject
    $sheet3 | Release-ComObject
    $sheet2 | Release-ComObject
    $sheet1 | Release-ComObject
    $workbook | Release-ComObject
    $excel | Release-ComObject
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
