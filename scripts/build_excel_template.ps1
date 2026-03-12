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
    $Worksheet.Range('A1').Value2 = 'Key'
    $Worksheet.Range('B1').Value2 = 'Value'
    $Worksheet.Range('A2').Value2 = 'InputFolder'
    $Worksheet.Range('A3').Value2 = 'OutputFile'
    $Worksheet.Range('A4').Value2 = 'LogFile'
    $Worksheet.Range('A5').Value2 = 'LastRunAt'
    $Worksheet.Range('A6').Value2 = 'Status'
    $Worksheet.Range('A8').Value2 = 'Usage'
    $Worksheet.Range('B8').Value2 = 'Run run_pdf2excel.bat or scripts\run_pdf2excel.ps1.'
    $Worksheet.Range('A10').Value2 = 'Notes'
    $Worksheet.Range('B10').Value2 = 'Text PDF and similar table layout are required.'
    $Worksheet.Range('A1:B1').Font.Bold = $true
    $Worksheet.Range('A8:A10').Font.Bold = $true
    $Worksheet.Columns.Item('A').ColumnWidth = 18
    $Worksheet.Columns.Item('B').ColumnWidth = 90
    $Worksheet.Range('B6').Value2 = 'Ready'
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
    Set-DataSheetLayout -Worksheet $sheet2 -SheetName 'Result' -Description 'Power Query result will be loaded here.'
    Set-DataSheetLayout -Worksheet $sheet3 -SheetName 'Errors' -Description 'Failed files will be listed here.'

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
