param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$testsRoot = Join-Path $projectRoot 'tests'
$workRoot = Join-Path $testsRoot 'work'
$fixturesRoot = Join-Path $workRoot 'fixtures'
$validPdfDir = Join-Path $fixturesRoot 'valid'
$mixedPdfDir = Join-Path $fixturesRoot 'mixed'
$duplicateRoot = Join-Path $fixturesRoot 'duplicate'
$duplicateA = Join-Path $duplicateRoot 'a'
$duplicateB = Join-Path $duplicateRoot 'b'
$resultsRoot = Join-Path $testsRoot 'results'
$reportsRoot = Join-Path $projectRoot 'reports'
$runScript = Join-Path $projectRoot 'scripts\run_pdf2excel.ps1'
$buildTemplateScript = Join-Path $projectRoot 'scripts\build_excel_template.ps1'
$batScript = Join-Path $projectRoot 'run_pdf2excel.bat'
$templatePath = Join-Path $projectRoot 'template\PDF2Excel_Converter.xlsm'
$jsonReportPath = Join-Path $resultsRoot 'integration-test-results.json'
$markdownReportPath = Join-Path $reportsRoot 'test-report.md'
$timeStarted = Get-Date

function Ensure-Directory {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Reset-Directory {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Recurse -Force
    }
    New-Item -ItemType Directory -Path $Path -Force | Out-Null
}

function Get-ExcelProcessIds {
    $processes = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
    if ($processes.Count -eq 0) {
        return @()
    }
    return @($processes | Select-Object -ExpandProperty Id)
}

function Wait-For-ExcelBaseline {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][int[]]$BaselineIds,
        [int]$TimeoutSeconds = 10
    )

    for ($attempt = 0; $attempt -lt $TimeoutSeconds; $attempt += 1) {
        $current = @(Get-ExcelProcessIds)
        $leaked = @($current | Where-Object { $BaselineIds -notcontains $_ })
        if ($leaked.Count -eq 0) {
            return @()
        }
        Start-Sleep -Seconds 1
    }

    $finalCurrent = @(Get-ExcelProcessIds)
    return @($finalCurrent | Where-Object { $BaselineIds -notcontains $_ })
}

function Assert-True {
    param(
        [Parameter(Mandatory = $true)][bool]$Condition,
        [Parameter(Mandatory = $true)][string]$Message
    )

    if (-not $Condition) {
        throw $Message
    }
}

function New-ExcelPdfFixture {
    param(
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)][string]$Prefix,
        [int]$DataRows = 4,
        [int]$Columns = 30
    )

    $excel = $null
    $workbook = $null
    $worksheet = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)

        for ($column = 1; $column -le $Columns; $column += 1) {
            $worksheet.Cells.Item(1, $column).Value2 = "H$column"
        }

        for ($row = 2; $row -le ($DataRows + 1); $row += 1) {
            for ($column = 1; $column -le $Columns; $column += 1) {
                if ((($row + $column) % 4) -ne 0) {
                    $worksheet.Cells.Item($row, $column).Value2 = "$Prefix-R$($row - 1)-C$column"
                }
            }
        }

        $worksheet.PageSetup.Orientation = 2
        $worksheet.PageSetup.Zoom = $false
        $worksheet.PageSetup.FitToPagesWide = 1
        $worksheet.PageSetup.FitToPagesTall = 1
        $workbook.ExportAsFixedFormat(0, $OutputPath)
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
        foreach ($comObject in @($worksheet, $workbook, $excel)) {
            try {
                if ($null -ne $comObject -and [System.Runtime.InteropServices.Marshal]::IsComObject($comObject)) {
                    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject)
                }
            } catch {
            }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Get-WorkbookSnapshot {
    param([Parameter(Mandatory = $true)][string]$WorkbookPath)

    $excel = $null
    $workbook = $null
    $resultSheet = $null
    $errorsSheet = $null
    $controlSheet = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($WorkbookPath)
        $controlSheet = $workbook.Worksheets.Item('Control')
        $resultSheet = $workbook.Worksheets.Item('Result')
        $errorsSheet = $workbook.Worksheets.Item('Errors')

        $resultRows = [int]$resultSheet.UsedRange.Rows.Count
        $resultColumns = [int]$resultSheet.UsedRange.Columns.Count
        $errorsRows = [int]$errorsSheet.UsedRange.Rows.Count
        $errorsColumns = [int]$errorsSheet.UsedRange.Columns.Count

        $sampleRange = $resultSheet.Range('A1:F6').Value2
        $sample = @()
        if ($null -ne $sampleRange) {
            if ($sampleRange -is [System.Array]) {
                for ($row = 1; $row -le $sampleRange.GetLength(0); $row += 1) {
                    $values = @()
                    for ($column = 1; $column -le $sampleRange.GetLength(1); $column += 1) {
                        $values += [string]$sampleRange[$row, $column]
                    }
                    $sample += ,$values
                }
            } else {
                $sample += ,@([string]$sampleRange)
            }
        }

        return [pscustomobject]@{
            ResultRows          = $resultRows
            ResultColumns       = $resultColumns
            ErrorsRows          = $errorsRows
            ErrorsColumns       = $errorsColumns
            ControlStatus       = [string]$controlSheet.Range('B6').Value2
            ControlSourceCount  = [string]$controlSheet.Range('B7').Value2
            ControlResultCount  = [string]$controlSheet.Range('B8').Value2
            ControlErrorCount   = [string]$controlSheet.Range('B9').Value2
            Sample              = $sample
        }
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
        foreach ($comObject in @($controlSheet, $errorsSheet, $resultSheet, $workbook, $excel)) {
            try {
                if ($null -ne $comObject -and [System.Runtime.InteropServices.Marshal]::IsComObject($comObject)) {
                    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject)
                }
            } catch {
            }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Invoke-TestCase {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Scenario,
        [Parameter(Mandatory = $true)][scriptblock]$Body
    )

    $baselineExcel = @(Get-ExcelProcessIds)
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $details = & $Body
        $stopwatch.Stop()
        $leaked = @(Wait-For-ExcelBaseline -BaselineIds $baselineExcel -TimeoutSeconds 10)
        Assert-True -Condition ($leaked.Count -eq 0) -Message ("Excel process leak detected: " + ($leaked -join ', '))

        return [pscustomobject]@{
            Name         = $Name
            Scenario     = $Scenario
            Status       = 'PASS'
            DurationMs   = [int]$stopwatch.ElapsedMilliseconds
            Details      = $details
            ErrorMessage = $null
        }
    } catch {
        $stopwatch.Stop()
        return [pscustomobject]@{
            Name         = $Name
            Scenario     = $Scenario
            Status       = 'FAIL'
            DurationMs   = [int]$stopwatch.ElapsedMilliseconds
            Details      = $null
            ErrorMessage = $_.Exception.Message
        }
    }
}

function New-ReportMarkdown {
    param(
        [Parameter(Mandatory = $true)]$TestResults,
        [Parameter(Mandatory = $true)][datetime]$StartedAt,
        [Parameter(Mandatory = $true)][datetime]$FinishedAt
    )

    $duration = [int]($FinishedAt - $StartedAt).TotalSeconds
    $passCount = @($TestResults | Where-Object Status -eq 'PASS').Count
    $failCount = @($TestResults | Where-Object Status -eq 'FAIL').Count
    $testEnv = "Windows / PowerShell $($PSVersionTable.PSVersion) / Excel(M365) COM"

    $lines = @(
        '# PDF2Excel テストレポート',
        '',
        "- 作成日: $($StartedAt.ToString('yyyy-MM-dd HH:mm')) JST",
        '- 作成者: Codex (GPT-5)',
        "- 更新日: $($FinishedAt.ToString('yyyy-MM-dd'))",
        '',
        '## 概要',
        '',
        "- 実施日時: $($StartedAt.ToString('yyyy-MM-dd HH:mm:ss')) JST - $($FinishedAt.ToString('yyyy-MM-dd HH:mm:ss')) JST",
        "- 対象環境: $testEnv",
        '- 対象URLまたは対象機能: ローカル PowerShell / BAT / Excel(M365) による PDF2Excel 一括変換',
        "- 結果概要: ${passCount}件成功 / ${failCount}件失敗",
        "- 所要時間または主要な応答時間: 全体 ${duration} 秒",
        ("- エラー有無: {0}" -f $(if ($failCount -eq 0) { 'なし' } else { 'あり' })),
        ("- 失敗時の原因推定: {0}" -f $(if ($failCount -eq 0) { '該当なし' } else { '各シナリオ欄を参照' })),
        '',
        '## 実行シナリオ',
        ''
    )

    $lines += '| No | シナリオ | 結果 | 所要時間 | 補足 |'
    $lines += '| --- | --- | --- | --- | --- |'

    $index = 1
    foreach ($result in $TestResults) {
        $detailText = if ($result.ErrorMessage) { $result.ErrorMessage } elseif ($result.Details) { $result.Details } else { '' }
        $lines += "| $index | $($result.Name) | $($result.Status) | $($result.DurationMs) ms | $detailText |"
        $index += 1
    }

    $lines += ''
    $lines += '## 判定'
    $lines += ''
    if ($failCount -eq 0) {
        $lines += '- 主要な正常系、異常系、運用系のシナリオはすべて通過しました。'
        $lines += '- 実行後に余分な Excel プロセスが残らないことを確認しました。'
        $lines += '- `output/runtime` に一時ファイルが残らず、自動清掃が機能しています。'
    } else {
        $lines += '- 一部のテストが失敗しました。上記のシナリオ一覧を確認してください。'
    }

    $lines += ''
    $lines += '## 備考'
    $lines += ''
    $lines += '- 抽出精度そのものは `Pdf.Tables` に依存するため、実業務PDFでは列ズレ確認を推奨します。'
    $lines += '- テスト用PDFは Excel から生成したテキストPDFを使用しました。'

    if ($failCount -gt 0) {
        $lines += ''
        $lines += '## 失敗詳細'
        $lines += ''
        foreach ($result in $TestResults | Where-Object Status -eq 'FAIL') {
            $lines += "- $($result.Name): $($result.ErrorMessage)"
        }
    }

    return ($lines -join [Environment]::NewLine)
}

Ensure-Directory -Path $testsRoot
Reset-Directory -Path $workRoot
Ensure-Directory -Path $validPdfDir
Ensure-Directory -Path $mixedPdfDir
Ensure-Directory -Path $duplicateA
Ensure-Directory -Path $duplicateB
Ensure-Directory -Path $resultsRoot
Ensure-Directory -Path $reportsRoot

New-ExcelPdfFixture -OutputPath (Join-Path $validPdfDir 'valid_a.pdf') -Prefix 'VALIDA'
New-ExcelPdfFixture -OutputPath (Join-Path $validPdfDir 'valid_b.pdf') -Prefix 'VALIDB'
Copy-Item -LiteralPath (Join-Path $validPdfDir 'valid_a.pdf') -Destination (Join-Path $mixedPdfDir 'valid_a.pdf') -Force
Set-Content -LiteralPath (Join-Path $mixedPdfDir 'broken.pdf') -Value 'not-a-real-pdf' -Encoding ASCII
New-ExcelPdfFixture -OutputPath (Join-Path $duplicateA 'duplicate.pdf') -Prefix 'DUPA'
New-ExcelPdfFixture -OutputPath (Join-Path $duplicateB 'duplicate.pdf') -Prefix 'DUPB'

$templateTimestampBefore = if (Test-Path -LiteralPath $templatePath) { (Get-Item -LiteralPath $templatePath).LastWriteTimeUtc } else { $null }

$testResults = @()
$testResults += Invoke-TestCase -Name 'Template build' -Scenario 'Template rebuild succeeds and updates the xlsm file' -Body {
    & powershell -NoProfile -ExecutionPolicy Bypass -File $buildTemplateScript
    Assert-True -Condition (Test-Path -LiteralPath $templatePath) -Message 'Template file was not created.'
    $newTimestamp = (Get-Item -LiteralPath $templatePath).LastWriteTimeUtc
    Assert-True -Condition ($null -eq $templateTimestampBefore -or $newTimestamp -ge $templateTimestampBefore) -Message 'Template timestamp was not updated.'
    "Template updated at $newTimestamp"
}

$testResults += Invoke-TestCase -Name 'PowerShell conversion' -Scenario 'Convert two valid PDFs via PowerShell and verify Result row count' -Body {
    $outputPath = Join-Path $resultsRoot 'powershell_success.xlsx'
    & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFolder $validPdfDir -OutputFile $outputPath
    Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message 'PowerShell output workbook was not created.'
    $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
    Assert-True -Condition ($snapshot.ResultColumns -eq 31) -Message "Expected 31 columns but got $($snapshot.ResultColumns)."
    Assert-True -Condition ($snapshot.ResultRows -eq 9) -Message "Expected 9 rows including header but got $($snapshot.ResultRows)."
    Assert-True -Condition ($snapshot.ErrorsRows -eq 1) -Message "Expected only Errors header row but got $($snapshot.ErrorsRows)."
    Assert-True -Condition ($snapshot.Sample[1][0] -eq 'valid_a.pdf') -Message 'SourceFile column did not contain the expected file name.'
    Assert-True -Condition ($snapshot.ControlSourceCount -eq '2') -Message "Control sheet source count is unexpected: $($snapshot.ControlSourceCount)"
    Assert-True -Condition ($snapshot.ControlResultCount -eq '8') -Message "Control sheet result count is unexpected: $($snapshot.ControlResultCount)"
    Assert-True -Condition ($snapshot.ControlErrorCount -eq '0') -Message "Control sheet error count is unexpected: $($snapshot.ControlErrorCount)"
    "ResultRows=$($snapshot.ResultRows), ErrorsRows=$($snapshot.ErrorsRows)"
}

$testResults += Invoke-TestCase -Name 'Single PDF conversion' -Scenario 'A folder that contains only one PDF should still convert successfully' -Body {
    $singlePdfDir = Join-Path $fixturesRoot 'single'
    Reset-Directory -Path $singlePdfDir
    Copy-Item -LiteralPath (Join-Path $validPdfDir 'valid_a.pdf') -Destination (Join-Path $singlePdfDir 'valid_a.pdf') -Force
    $outputPath = Join-Path $resultsRoot 'single_pdf_success.xlsx'
    & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFolder $singlePdfDir -OutputFile $outputPath
    Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message 'Single PDF output workbook was not created.'
    $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
    Assert-True -Condition ($snapshot.ResultRows -eq 5) -Message "Expected 5 rows including header for one PDF but got $($snapshot.ResultRows)."
    Assert-True -Condition ($snapshot.ControlSourceCount -eq '1') -Message "Control sheet source count is unexpected: $($snapshot.ControlSourceCount)"
    'Single PDF conversion succeeded'
}

$testResults += Invoke-TestCase -Name 'InputFiles conversion' -Scenario 'Documented -InputFiles usage should convert multiple PDFs correctly' -Body {
    $outputPath = Join-Path $resultsRoot 'inputfiles_success.xlsx'
    $inputFilesArg = @(
        (Join-Path $validPdfDir 'valid_a.pdf'),
        (Join-Path $validPdfDir 'valid_b.pdf')
    ) -join ','
    & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFiles $inputFilesArg -OutputFile $outputPath
    Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message 'InputFiles output workbook was not created.'
    $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
    Assert-True -Condition ($snapshot.ResultRows -eq 9) -Message "InputFiles conversion returned unexpected Result row count: $($snapshot.ResultRows)."
    Assert-True -Condition ($snapshot.ControlSourceCount -eq '2') -Message "Control sheet source count is unexpected: $($snapshot.ControlSourceCount)"
    'InputFiles conversion succeeded'
}

$testResults += Invoke-TestCase -Name 'BAT conversion' -Scenario 'Convert the same valid PDFs through BAT' -Body {
    $outputPath = Join-Path $resultsRoot 'bat_success.xlsx'
    & cmd /c $batScript -InputFolder $validPdfDir -OutputFile $outputPath
    Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message 'BAT output workbook was not created.'
    $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
    Assert-True -Condition ($snapshot.ResultRows -eq 9) -Message "BAT conversion returned unexpected Result row count: $($snapshot.ResultRows)."
    "ResultRows=$($snapshot.ResultRows)"
}

$testResults += Invoke-TestCase -Name 'BAT direct no-pause' -Scenario 'BAT direct execution with arguments should exit without waiting for key input' -Body {
    $outputPath = Join-Path $resultsRoot 'bat_direct_no_pause.xlsx'
    $process = Start-Process -FilePath 'cmd.exe' -ArgumentList '/c', $batScript, '-InputFolder', $validPdfDir, '-OutputFile', $outputPath -NoNewWindow -Wait -PassThru
    Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message 'BAT direct output workbook was not created.'
    Assert-True -Condition ($process.ExitCode -eq 0) -Message "BAT direct execution failed with exit code $($process.ExitCode)."
    'BAT direct execution finished without pause'
}

$testResults += Invoke-TestCase -Name 'Input self-reference' -Scenario 'Using input/ itself as InputFolder should succeed without self-deleting files' -Body {
    Get-ChildItem -LiteralPath (Join-Path $projectRoot 'input') -Filter '*.pdf' -File -ErrorAction SilentlyContinue | Remove-Item -Force
    Copy-Item -LiteralPath (Join-Path $validPdfDir 'valid_a.pdf') -Destination (Join-Path $projectRoot 'input\valid_a.pdf') -Force
    Copy-Item -LiteralPath (Join-Path $validPdfDir 'valid_b.pdf') -Destination (Join-Path $projectRoot 'input\valid_b.pdf') -Force
    $outputPath = Join-Path $resultsRoot 'input_self_reference.xlsx'
    & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFolder (Join-Path $projectRoot 'input') -OutputFile $outputPath
    Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message 'Self-reference output workbook was not created.'
    $remaining = @(Get-ChildItem -LiteralPath (Join-Path $projectRoot 'input') -Filter '*.pdf' -File)
    $remainingNames = @($remaining | Select-Object -ExpandProperty Name)
    Assert-True -Condition ($remainingNames -contains 'valid_a.pdf') -Message 'Input self-reference run removed valid_a.pdf unexpectedly.'
    Assert-True -Condition ($remainingNames -contains 'valid_b.pdf') -Message 'Input self-reference run removed valid_b.pdf unexpectedly.'
    "RemainingInputFiles=$($remainingNames -join ',')"
}

$testResults += Invoke-TestCase -Name 'Duplicate filename rejection' -Scenario 'Different folders with the same PDF file name should be rejected explicitly' -Body {
    $outputPath = Join-Path $resultsRoot 'duplicate_should_fail.xlsx'
    if (Test-Path -LiteralPath $outputPath) {
        Remove-Item -LiteralPath $outputPath -Force
    }
    $duplicatePaths = @(
        (Join-Path $duplicateA 'duplicate.pdf'),
        (Join-Path $duplicateB 'duplicate.pdf')
    ) -join ','
    $output = ''
    $duplicateFailed = $false
    try {
        $output = & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFiles $duplicatePaths -OutputFile $outputPath 2>&1 | Out-String
    } catch {
        $duplicateFailed = $true
        $output = $_ | Out-String
    }
    if (-not $duplicateFailed -and $LASTEXITCODE -ne 0) {
        $duplicateFailed = $true
    }
    Assert-True -Condition $duplicateFailed -Message 'Duplicate filename run unexpectedly succeeded.'
    Assert-True -Condition (-not (Test-Path -LiteralPath $outputPath)) -Message 'Duplicate filename run should not create an output workbook.'
    Assert-True -Condition ($output.Contains('同名の PDF は同時に処理できません')) -Message 'Duplicate filename rejection message was not found.'
    'Duplicate PDF names are rejected'
}

$testResults += Invoke-TestCase -Name 'Broken PDF handling' -Scenario 'A broken PDF should not crash the run and should be reported in Errors' -Body {
    $outputPath = Join-Path $resultsRoot 'mixed_broken.xlsx'
    & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFolder $mixedPdfDir -OutputFile $outputPath
    Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message 'Mixed output workbook was not created.'
    $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
    Assert-True -Condition ($snapshot.ResultRows -eq 5) -Message "Expected 5 Result rows for one valid PDF but got $($snapshot.ResultRows)."
    Assert-True -Condition ($snapshot.ErrorsRows -ge 2) -Message 'Broken PDF was not reported in Errors sheet.'
    Assert-True -Condition ($snapshot.ControlErrorCount -eq '1') -Message "Control sheet error count is unexpected: $($snapshot.ControlErrorCount)"
    "ResultRows=$($snapshot.ResultRows), ErrorsRows=$($snapshot.ErrorsRows)"
}

$testResults += Invoke-TestCase -Name 'Nested output path' -Scenario 'Workbook can be saved into a new nested output directory' -Body {
    $nestedDir = Join-Path $resultsRoot 'nested\child\output'
    $outputPath = Join-Path $nestedDir 'nested_output.xlsx'
    if (Test-Path -LiteralPath $nestedDir) {
        Remove-Item -LiteralPath $nestedDir -Recurse -Force
    }
    & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFolder $validPdfDir -OutputFile $outputPath
    Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message 'Nested output workbook was not created.'
    "Nested output created"
}

$testResults += Invoke-TestCase -Name 'Runtime cleanup' -Scenario 'No temporary xlsm remains in output/runtime after execution' -Body {
    $runtimeFiles = @(Get-ChildItem -LiteralPath (Join-Path $projectRoot 'output\runtime') -File -Force | Where-Object { $_.Name -ne '.gitkeep' })
    $runtimeFileNames = @($runtimeFiles | Select-Object -ExpandProperty Name)
    Assert-True -Condition ($runtimeFiles.Count -eq 0) -Message ('Runtime files remained: ' + ($runtimeFileNames -join ', '))
    'Runtime directory is clean'
}

$testResults += Invoke-TestCase -Name 'No Excel leak' -Scenario 'No EXCEL.exe process remains after the full suite' -Body {
    $excelIds = @(Wait-For-ExcelBaseline -BaselineIds @() -TimeoutSeconds 10)
    Assert-True -Condition ($excelIds.Count -eq 0) -Message ('Excel processes still running: ' + ($excelIds -join ', '))
    'No Excel process remains'
}

$timeFinished = Get-Date

$summary = [pscustomobject]@{
    StartedAt = $timeStarted
    FinishedAt = $timeFinished
    Results = $testResults
}

$summary | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $jsonReportPath -Encoding UTF8
$reportMarkdown = New-ReportMarkdown -TestResults $testResults -StartedAt $timeStarted -FinishedAt $timeFinished
$reportMarkdown | Set-Content -LiteralPath $markdownReportPath -Encoding UTF8

$failed = @($testResults | Where-Object Status -eq 'FAIL')
if ($failed.Count -gt 0) {
    Write-Error ("Integration tests failed: " + ($failed.Name -join ', '))
}

Write-Host "Integration tests passed: $(@($testResults | Where-Object Status -eq 'PASS').Count) / $($testResults.Count)"
Write-Host "Markdown report: $markdownReportPath"
Write-Host "JSON report: $jsonReportPath"



