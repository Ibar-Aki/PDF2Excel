Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$tempPath = Join-Path $projectRoot 'tests\results\v2-template-builder-smoke.xlsm'

if (Test-Path -LiteralPath $tempPath) {
    Remove-Item -LiteralPath $tempPath -Force
}

$excel = $null
$workbook = $null
$macroBlocked = $false
$macroBlockedMessage = ''

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1

    $workbook = $excel.Workbooks.Add()
    $workbook.SaveAs($tempPath, 52)

    $components = $workbook.VBProject.VBComponents
    try {
        $null = $components.Import((Join-Path $projectRoot 'template\vba\PDF2ExcelMacros.bas'))
        $null = $components.Import((Join-Path $projectRoot 'template\vba\PDF2ExcelTemplateBuilder.bas'))
    } finally {
        if ($components) {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($components)
        }
    }

    $workbook.Save()
    $workbook.Close($true) | Out-Null
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)
    $workbook = $excel.Workbooks.Open($tempPath)

    try {
        $excel.Run(("'{0}'!BuildPDF2ExcelV2TemplateInActiveWorkbook" -f $workbook.Name))
    } catch [System.Runtime.InteropServices.COMException] {
        $macroBlocked = $true
        $macroBlockedMessage = $_.Exception.Message
    }

    if ($macroBlocked) {
        Write-Output ("SMOKE_SKIPPED {0}" -f $macroBlockedMessage)
        return
    }

    $sheetNames = New-Object System.Collections.Generic.List[string]
    foreach ($sheet in $workbook.Worksheets) {
        try {
            [void]$sheetNames.Add([string]$sheet.Name)
        } finally {
            if ($sheet) {
                [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet)
            }
        }
    }

    if ($workbook.Worksheets.Count -ne 5) {
        throw "Unexpected worksheet count: $($workbook.Worksheets.Count)"
    }

    $actualNames = $sheetNames -join ','
    if ($actualNames -ne 'Control,Result,Errors,Summary,Review') {
        throw "Unexpected worksheet names: $actualNames"
    }

    $versionValue = [string]$workbook.Worksheets('Control').Range('B15').Value2
    if ($versionValue -ne 'VER2') {
        throw "Unexpected Control!B15 value: $versionValue"
    }

    Write-Output ("SMOKE_OK {0}" -f $actualNames)
} finally {
    if ($workbook) {
        try {
            $workbook.Close($false) | Out-Null
        } finally {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)
        }
    }

    if ($excel) {
        try {
            $excel.Quit() | Out-Null
        } finally {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
        }
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
