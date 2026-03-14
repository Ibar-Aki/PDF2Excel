param(
    [string]$CaseName,
    [string]$SingleResultPath
)

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
$profileFixtureDir = Join-Path $fixturesRoot 'profile10'
$japanesePdfDir = Join-Path $fixturesRoot 'japanese'
$bulkPdfDir = Join-Path $fixturesRoot 'bulk50'
$resultsRoot = Join-Path $testsRoot 'results'
$reportsRoot = Join-Path $projectRoot 'reports'
$runScript = Join-Path $projectRoot 'scripts\run_pdf2excel.ps1'
$buildTemplateScript = Join-Path $projectRoot 'scripts\build_excel_template.ps1'
$batScript = Join-Path $projectRoot 'run_pdf2excel.bat'
$commonScript = Join-Path $projectRoot 'scripts\pdf2excel.common.ps1'
$templatePath = Join-Path $projectRoot 'template\PDF2Excel_Converter.xlsm'
$jsonReportPath = Join-Path $resultsRoot 'integration-test-results.json'
$markdownReportPath = Join-Path $reportsRoot 'test-report.md'
$timeStarted = Get-Date
$script:customProfilePath = Join-Path $workRoot 'profile10.json'
$script:selfPath = $MyInvocation.MyCommand.Path

. $commonScript

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
    $summarySheet = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($WorkbookPath)
        $controlSheet = $workbook.Worksheets.Item('Control')
        $summarySheet = $workbook.Worksheets.Item('Summary')
        $resultSheet = $workbook.Worksheets.Item('Result')
        $errorsSheet = $workbook.Worksheets.Item('Errors')

        $resultRows = [int]$resultSheet.UsedRange.Rows.Count
        $resultColumns = [int]$resultSheet.UsedRange.Columns.Count
        $errorsRows = [int]$errorsSheet.UsedRange.Rows.Count
        $errorsColumns = [int]$errorsSheet.UsedRange.Columns.Count

        $sampleRange = $resultSheet.Range('A1:F12').Value2
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
            ControlSuccessCount = [string]$controlSheet.Range('B10').Value2
            ControlFailedCount  = [string]$controlSheet.Range('B11').Value2
            ControlElapsed      = [string]$controlSheet.Range('B12').Value2
            ControlProfile      = [string]$controlSheet.Range('B13').Value2
            SummaryTitle        = [string]$summarySheet.Range('A1').Value2
            SummaryFileCount    = [string]$summarySheet.Range('B5').Value2
            SummarySuccessCount = [string]$summarySheet.Range('B6').Value2
            SummaryFailedCount  = [string]$summarySheet.Range('B7').Value2
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

        foreach ($comObject in @($summarySheet, $controlSheet, $errorsSheet, $resultSheet, $workbook, $excel)) {
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
        '# PDF2Excel 統合テストレポート',
        '',
        ('- 作成日: {0} JST' -f $StartedAt.ToString("yyyy-MM-dd HH:mm")),
        '- 作成者: Codex (GPT-5)',
        ('- 更新日: {0}' -f $FinishedAt.ToString("yyyy-MM-dd")),
        '',
        '## サマリー',
        '',
        ('- 実施日時: {0} JST - {1} JST' -f $StartedAt.ToString("yyyy-MM-dd HH:mm:ss"), $FinishedAt.ToString("yyyy-MM-dd HH:mm:ss")),
        ('- 対象環境: {0}' -f $testEnv),
        '- 対象機能: PowerShell / BAT / Excel(M365) による PDF2Excel 一括変換',
        ('- 結果概要: {0} 件成功 / {1} 件失敗' -f $passCount, $failCount),
        ('- 所要時間: {0} 秒' -f $duration),
        ("- エラー有無: {0}" -f $(if ($failCount -eq 0) { 'なし' } else { 'あり' })),
        ("- 失敗概要: {0}" -f $(if ($failCount -eq 0) { 'なし' } else { '失敗シナリオ一覧を参照' })),
        '',
        '## シナリオ別結果',
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
    $lines += '## 総評'
    $lines += ''
    if ($failCount -eq 0) {
        $lines += '- 主要な正常系、異常系、運用系シナリオはすべて成功しました。'
        $lines += '- 各シナリオは子プロセス隔離とタイムアウト監視付きで実行されました。'
        $lines += '- 実行後に余分な Excel プロセスは残りませんでした。'
        $lines += '- `output/runtime/runs` 配下に一時ワークスペースは残りませんでした。'
    } else {
        $lines += '- 一部のテストが失敗しました。上記のシナリオ一覧を確認してください。'
    }

    $lines += ''
    $lines += '## 補足'
    $lines += ''
    $lines += '- 抽出精度は引き続き `Pdf.Tables` に依存するため、実業務PDFでの確認が必要です。'
    $lines += '- テスト用 PDF は Excel から生成したテキスト PDF を使用しています。'

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

function Initialize-TestFixtures {
    Ensure-Directory -Path $testsRoot
    Reset-Directory -Path $workRoot
    Ensure-Directory -Path $validPdfDir
    Ensure-Directory -Path $mixedPdfDir
    Ensure-Directory -Path $duplicateA
    Ensure-Directory -Path $duplicateB
    Ensure-Directory -Path $profileFixtureDir
    Ensure-Directory -Path $japanesePdfDir
    Ensure-Directory -Path $bulkPdfDir
    Ensure-Directory -Path $resultsRoot
    Ensure-Directory -Path $reportsRoot

    New-ExcelPdfFixture -OutputPath (Join-Path $validPdfDir 'valid_a.pdf') -Prefix 'VALIDA'
    New-ExcelPdfFixture -OutputPath (Join-Path $validPdfDir 'valid_b.pdf') -Prefix 'VALIDB'
    Copy-Item -LiteralPath (Join-Path $validPdfDir 'valid_a.pdf') -Destination (Join-Path $mixedPdfDir 'valid_a.pdf') -Force
    Set-Content -LiteralPath (Join-Path $mixedPdfDir 'broken.pdf') -Value 'not-a-real-pdf' -Encoding ASCII
    New-ExcelPdfFixture -OutputPath (Join-Path $duplicateA 'duplicate.pdf') -Prefix 'DUPA'
    New-ExcelPdfFixture -OutputPath (Join-Path $duplicateB 'duplicate.pdf') -Prefix 'DUPB'
    New-ExcelPdfFixture -OutputPath (Join-Path $profileFixtureDir 'profile10.pdf') -Prefix 'P10' -Columns 10
    Copy-Item -LiteralPath (Join-Path $validPdfDir 'valid_a.pdf') -Destination (Join-Path $japanesePdfDir '日本語_帳票A.pdf') -Force
    Copy-Item -LiteralPath (Join-Path $validPdfDir 'valid_b.pdf') -Destination (Join-Path $japanesePdfDir '請求書_テストB.pdf') -Force
    for ($index = 1; $index -le 50; $index += 1) {
        $sourceName = if (($index % 2) -eq 0) { 'valid_b.pdf' } else { 'valid_a.pdf' }
        $bulkName = 'bulk_{0:D2}.pdf' -f $index
        Copy-Item -LiteralPath (Join-Path $validPdfDir $sourceName) -Destination (Join-Path $bulkPdfDir $bulkName) -Force
    }

    $customProfileJson = @'
{
  "name": "profile10",
  "displayName": "Profile10 Test",
  "description": "Integration test profile for 10 columns.",
  "expectedColumns": 10,
  "headerRowsToSkip": 1,
  "targetRowCount": 4,
  "allowMoreColumns": false,
  "preferredTableKinds": [
    "Table"
  ],
  "preferredTableNameContains": [],
  "preferredTableIdContains": [],
  "sourceFileColumnName": "SourceFile",
  "dataColumnPrefix": "Column"
}
'@
    $customProfileJson | Set-Content -LiteralPath $script:customProfilePath -Encoding UTF8
}

function Invoke-TestProcess {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [string[]]$ArgumentList,
        [int]$TimeoutSeconds = 180
    )

    $stdoutPath = Join-Path $workRoot ([guid]::NewGuid().ToString() + '.stdout.log')
    $stderrPath = Join-Path $workRoot ([guid]::NewGuid().ToString() + '.stderr.log')

    try {
        $process = Start-Process -FilePath $FilePath -ArgumentList $ArgumentList -PassThru -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath -WindowStyle Hidden
        if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
            Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
            throw "Process timeout after ${TimeoutSeconds}s: $FilePath $($ArgumentList -join ' ')"
        }

        return [pscustomobject]@{
            ExitCode = $process.ExitCode
            StdOut   = if (Test-Path -LiteralPath $stdoutPath) { Get-Content -LiteralPath $stdoutPath -Raw -Encoding UTF8 } else { '' }
            StdErr   = if (Test-Path -LiteralPath $stderrPath) { Get-Content -LiteralPath $stderrPath -Raw -Encoding UTF8 } else { '' }
        }
    } finally {
        foreach ($path in @($stdoutPath, $stderrPath)) {
            if (Test-Path -LiteralPath $path) {
                Remove-Item -LiteralPath $path -Force
            }
        }
    }
}

function Get-TestCases {
    return @(
        [pscustomobject]@{ Name = 'Template build'; Scenario = 'Template rebuild succeeds and updates the xlsm file'; TimeoutSeconds = 120 },
        [pscustomobject]@{ Name = 'PowerShell conversion'; Scenario = 'Convert two valid PDFs via PowerShell and verify Result row count'; TimeoutSeconds = 180 },
        [pscustomobject]@{ Name = 'Single PDF conversion'; Scenario = 'A folder that contains only one PDF should still convert successfully'; TimeoutSeconds = 180 },
        [pscustomobject]@{ Name = 'InputFiles conversion'; Scenario = 'Documented -InputFiles usage should convert multiple PDFs correctly'; TimeoutSeconds = 180 },
        [pscustomobject]@{ Name = 'Japanese filename conversion'; Scenario = 'Japanese PDF file names should remain intact in Result and Summary'; TimeoutSeconds = 180 },
        [pscustomobject]@{ Name = 'Bulk 50 PDF performance'; Scenario = 'A 50-file batch should complete within the agreed timeout and preserve row counts'; TimeoutSeconds = 480 },
        [pscustomobject]@{ Name = 'KeepInput isolation'; Scenario = 'KeepInput keeps archived PDFs without re-importing them into the current run'; TimeoutSeconds = 180 },
        [pscustomobject]@{ Name = 'Concurrent run lock'; Scenario = 'A second run should fail fast while another run is already in progress'; TimeoutSeconds = 180 },
        [pscustomobject]@{ Name = 'BAT conversion'; Scenario = 'Convert the same valid PDFs through BAT'; TimeoutSeconds = 180 },
        [pscustomobject]@{ Name = 'BAT direct no-pause'; Scenario = 'BAT direct execution with arguments should exit without waiting for key input'; TimeoutSeconds = 180 },
        [pscustomobject]@{ Name = 'Input self-reference'; Scenario = 'Using input/ itself as InputFolder should succeed without self-deleting files'; TimeoutSeconds = 180 },
        [pscustomobject]@{ Name = 'Duplicate filename rejection'; Scenario = 'Different folders with the same PDF file name should be rejected explicitly'; TimeoutSeconds = 120 },
        [pscustomobject]@{ Name = 'Broken PDF handling'; Scenario = 'A broken PDF should not crash the run and should be reported in Errors'; TimeoutSeconds = 180 },
        [pscustomobject]@{ Name = 'Nested output path'; Scenario = 'Workbook can be saved into a new nested output directory'; TimeoutSeconds = 180 },
        [pscustomobject]@{ Name = 'Profile-based conversion'; Scenario = 'A custom profile should change the expected output columns and summary metadata'; TimeoutSeconds = 180 },
        [pscustomobject]@{ Name = 'Runtime cleanup'; Scenario = 'No temporary run workspace remains in output/runtime/runs after execution'; TimeoutSeconds = 60 },
        [pscustomobject]@{ Name = 'No Excel leak'; Scenario = 'No EXCEL.exe process remains after the full suite'; TimeoutSeconds = 60 }
    )
}

function Invoke-NamedScenario {
    param([Parameter(Mandatory = $true)][string]$Name)

    switch ($Name) {
        'Template build' {
            $templateTimestampBefore = if (Test-Path -LiteralPath $templatePath) { (Get-Item -LiteralPath $templatePath).LastWriteTimeUtc } else { $null }
            & powershell -NoProfile -ExecutionPolicy Bypass -File $buildTemplateScript
            Assert-True -Condition (Test-Path -LiteralPath $templatePath) -Message 'Template file was not created.'
            $newTimestamp = (Get-Item -LiteralPath $templatePath).LastWriteTimeUtc
            Assert-True -Condition ($null -eq $templateTimestampBefore -or $newTimestamp -ge $templateTimestampBefore) -Message 'Template timestamp was not updated.'
            return "Template updated at $newTimestamp"
        }
        'PowerShell conversion' {
            $outputPath = Join-Path $resultsRoot 'powershell_success.xlsx'
            & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFolder $validPdfDir -OutputFile $outputPath -NoConfirm
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message 'PowerShell output workbook was not created.'
            $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
            Assert-True -Condition ($snapshot.ResultColumns -eq 31) -Message "Expected 31 columns but got $($snapshot.ResultColumns)."
            Assert-True -Condition ($snapshot.ResultRows -eq 9) -Message "Expected 9 rows including header but got $($snapshot.ResultRows)."
            Assert-True -Condition ($snapshot.ErrorsRows -eq 1) -Message "Expected only Errors header row but got $($snapshot.ErrorsRows)."
            Assert-True -Condition ($snapshot.Sample[1][0] -eq 'valid_a.pdf') -Message 'SourceFile column did not contain the expected file name.'
            Assert-True -Condition ($snapshot.ControlSourceCount -eq '2') -Message "Control sheet source count is unexpected: $($snapshot.ControlSourceCount)"
            Assert-True -Condition ($snapshot.ControlResultCount -eq '8') -Message "Control sheet result count is unexpected: $($snapshot.ControlResultCount)"
            Assert-True -Condition ($snapshot.ControlErrorCount -eq '0') -Message "Control sheet error count is unexpected: $($snapshot.ControlErrorCount)"
            Assert-True -Condition ($snapshot.ControlSuccessCount -eq '2') -Message "Control sheet success count is unexpected: $($snapshot.ControlSuccessCount)"
            Assert-True -Condition ($snapshot.ControlFailedCount -eq '0') -Message "Control sheet failed count is unexpected: $($snapshot.ControlFailedCount)"
            Assert-True -Condition (-not [string]::IsNullOrWhiteSpace($snapshot.ControlProfile)) -Message 'Control sheet profile is unexpectedly blank.'
            Assert-True -Condition ($snapshot.SummaryTitle -eq 'Summary') -Message 'Summary sheet title is missing.'
            Assert-True -Condition ($snapshot.SummaryFileCount -eq '2') -Message "Summary file count is unexpected: $($snapshot.SummaryFileCount)"
            return "ResultRows=$($snapshot.ResultRows), ErrorsRows=$($snapshot.ErrorsRows)"
        }
        'Single PDF conversion' {
            $singlePdfDir = Join-Path $fixturesRoot 'single'
            Reset-Directory -Path $singlePdfDir
            Copy-Item -LiteralPath (Join-Path $validPdfDir 'valid_a.pdf') -Destination (Join-Path $singlePdfDir 'valid_a.pdf') -Force
            $outputPath = Join-Path $resultsRoot 'single_pdf_success.xlsx'
            & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFolder $singlePdfDir -OutputFile $outputPath -NoConfirm
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message 'Single PDF output workbook was not created.'
            $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
            Assert-True -Condition ($snapshot.ResultRows -eq 5) -Message "Expected 5 rows including header for one PDF but got $($snapshot.ResultRows)."
            Assert-True -Condition ($snapshot.ControlSourceCount -eq '1') -Message "Control sheet source count is unexpected: $($snapshot.ControlSourceCount)"
            return 'Single PDF conversion succeeded'
        }
        'InputFiles conversion' {
            $outputPath = Join-Path $resultsRoot 'inputfiles_success.xlsx'
            $inputFilesArg = @((Join-Path $validPdfDir 'valid_a.pdf'), (Join-Path $validPdfDir 'valid_b.pdf')) -join ','
            & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFiles $inputFilesArg -OutputFile $outputPath -NoConfirm
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message 'InputFiles output workbook was not created.'
            $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
            Assert-True -Condition ($snapshot.ResultRows -eq 9) -Message "InputFiles conversion returned unexpected Result row count: $($snapshot.ResultRows)."
            Assert-True -Condition ($snapshot.ControlSourceCount -eq '2') -Message "Control sheet source count is unexpected: $($snapshot.ControlSourceCount)"
            return 'InputFiles conversion succeeded'
        }
        'Japanese filename conversion' {
            $outputPath = Join-Path $resultsRoot 'japanese_success.xlsx'
            & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFolder $japanesePdfDir -OutputFile $outputPath -NoConfirm
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message 'Japanese filename output workbook was not created.'
            $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
            $sourceNames = @($snapshot.Sample | Select-Object -Skip 1 | ForEach-Object { $_[0] } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
            Assert-True -Condition ($snapshot.ControlSourceCount -eq '2') -Message "Control sheet source count is unexpected: $($snapshot.ControlSourceCount)"
            Assert-True -Condition ($sourceNames -contains '日本語_帳票A.pdf') -Message 'Japanese file name A was not preserved in Result.'
            Assert-True -Condition ($sourceNames -contains '請求書_テストB.pdf') -Message 'Japanese file name B was not preserved in Result.'
            return "SourceNames=$($sourceNames -join ',')"
        }
        'Bulk 50 PDF performance' {
            $outputPath = Join-Path $resultsRoot 'bulk50_success.xlsx'
            $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFolder $bulkPdfDir -OutputFile $outputPath -NoConfirm
            $stopwatch.Stop()
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message 'Bulk 50 PDF workbook was not created.'
            $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
            Assert-True -Condition ($snapshot.ControlSourceCount -eq '50') -Message "Expected 50 source PDFs but got $($snapshot.ControlSourceCount)."
            Assert-True -Condition ($snapshot.ResultRows -eq 201) -Message "Expected 201 rows including header for 50 PDFs but got $($snapshot.ResultRows)."
            Assert-True -Condition ($stopwatch.Elapsed.TotalSeconds -lt 300) -Message ("50 PDF batch took too long: {0:N1} seconds." -f $stopwatch.Elapsed.TotalSeconds)
            return ("ElapsedSeconds={0:N1}, ResultRows={1}" -f $stopwatch.Elapsed.TotalSeconds, $snapshot.ResultRows)
        }
        'KeepInput isolation' {
            $inputStore = Join-Path $projectRoot 'input'
            Get-ChildItem -LiteralPath $inputStore -Filter '*.pdf' -File -ErrorAction SilentlyContinue | Remove-Item -Force
            Copy-Item -LiteralPath (Join-Path $validPdfDir 'valid_b.pdf') -Destination (Join-Path $inputStore 'archived_valid_b.pdf') -Force

            $singleKeepDir = Join-Path $fixturesRoot 'single_keep'
            Reset-Directory -Path $singleKeepDir
            Copy-Item -LiteralPath (Join-Path $validPdfDir 'valid_a.pdf') -Destination (Join-Path $singleKeepDir 'valid_a.pdf') -Force

            $outputPath = Join-Path $resultsRoot 'keepinput_isolation.xlsx'
            & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFolder $singleKeepDir -OutputFile $outputPath -KeepInput -NoConfirm
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message 'KeepInput output workbook was not created.'
            $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
            Assert-True -Condition ($snapshot.ResultRows -eq 5) -Message "KeepInput should import only the selected PDF, but ResultRows=$($snapshot.ResultRows)."

            $storedNames = @(Get-ChildItem -LiteralPath $inputStore -Filter '*.pdf' -File | Select-Object -ExpandProperty Name)
            Assert-True -Condition ($storedNames -contains 'archived_valid_b.pdf') -Message 'KeepInput did not preserve the archived PDF.'
            Assert-True -Condition ($storedNames -contains 'valid_a.pdf') -Message 'KeepInput did not store the selected PDF.'
            return "StoredInputFiles=$($storedNames -join ',')"
        }
        'Concurrent run lock' {
            $secondOutput = Join-Path $resultsRoot 'lock_second.xlsx'
            $runtimeRoot = Join-Path $projectRoot 'output\runtime'
            $lockPath = Join-Path $runtimeRoot 'run.lock'
            $mutex = $null
            $lockAcquired = $false
            $output = ''

            try {
                Ensure-Directory -Path $runtimeRoot
                $createdNew = $false
                $mutex = New-Object System.Threading.Mutex($false, 'Global\PDF2Excel_RunMutex', [ref]$createdNew)
                $lockAcquired = $mutex.WaitOne(0, $false)
                Assert-True -Condition $lockAcquired -Message 'Test mutex could not be acquired.'

                $lockPayload = [ordered]@{
                    runInstanceId = 'test-lock'
                    pid           = $PID
                    startedAt     = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                    machineName   = $env:COMPUTERNAME
                    userName      = $env:USERNAME
                } | ConvertTo-Json
                Set-Content -LiteralPath $lockPath -Value $lockPayload -Encoding UTF8

                $secondFailed = $false
                try {
                    $output = & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFolder $validPdfDir -OutputFile $secondOutput -NoConfirm 2>&1 | Out-String
                } catch {
                    $secondFailed = $true
                    $output = $_ | Out-String
                }

                if (-not $secondFailed -and $LASTEXITCODE -ne 0) {
                    $secondFailed = $true
                }

                Assert-True -Condition $secondFailed -Message 'Second run unexpectedly succeeded while the mutex was held.'
                Assert-True -Condition (-not (Test-Path -LiteralPath $secondOutput)) -Message 'Second run created an output workbook unexpectedly.'
                return 'Concurrent lock rejected the second run'
            } finally {
                if ($lockAcquired -and $mutex) {
                    try {
                        $mutex.ReleaseMutex() | Out-Null
                    } catch {
                    }
                }
                if ($mutex) {
                    try {
                        $mutex.Dispose()
                    } catch {
                    }
                }
                if (Test-Path -LiteralPath $lockPath) {
                    Remove-Item -LiteralPath $lockPath -Force
                }
            }
        }
        'BAT conversion' {
            $outputPath = Join-Path $resultsRoot 'bat_success.xlsx'
            & cmd /c $batScript -InputFolder $validPdfDir -OutputFile $outputPath -NoConfirm
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message 'BAT output workbook was not created.'
            $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
            Assert-True -Condition ($snapshot.ResultRows -eq 9) -Message "BAT conversion returned unexpected Result row count: $($snapshot.ResultRows)."
            return "ResultRows=$($snapshot.ResultRows)"
        }
        'BAT direct no-pause' {
            $outputPath = Join-Path $resultsRoot 'bat_direct_no_pause.xlsx'
            & cmd /c $batScript -InputFolder $validPdfDir -OutputFile $outputPath -NoConfirm
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message 'BAT direct output workbook was not created.'
            return 'BAT direct execution finished without pause'
        }
        'Input self-reference' {
            Get-ChildItem -LiteralPath (Join-Path $projectRoot 'input') -Filter '*.pdf' -File -ErrorAction SilentlyContinue | Remove-Item -Force
            Copy-Item -LiteralPath (Join-Path $validPdfDir 'valid_a.pdf') -Destination (Join-Path $projectRoot 'input\valid_a.pdf') -Force
            Copy-Item -LiteralPath (Join-Path $validPdfDir 'valid_b.pdf') -Destination (Join-Path $projectRoot 'input\valid_b.pdf') -Force
            $outputPath = Join-Path $resultsRoot 'input_self_reference.xlsx'
            & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFolder (Join-Path $projectRoot 'input') -OutputFile $outputPath -NoConfirm
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message 'Self-reference output workbook was not created.'
            $remainingNames = @(Get-ChildItem -LiteralPath (Join-Path $projectRoot 'input') -Filter '*.pdf' -File | Select-Object -ExpandProperty Name)
            Assert-True -Condition ($remainingNames -contains 'valid_a.pdf') -Message 'Input self-reference run removed valid_a.pdf unexpectedly.'
            Assert-True -Condition ($remainingNames -contains 'valid_b.pdf') -Message 'Input self-reference run removed valid_b.pdf unexpectedly.'
            return "RemainingInputFiles=$($remainingNames -join ',')"
        }
        'Duplicate filename rejection' {
            $outputPath = Join-Path $resultsRoot 'duplicate_should_fail.xlsx'
            if (Test-Path -LiteralPath $outputPath) {
                Remove-Item -LiteralPath $outputPath -Force
            }

            $duplicatePaths = @((Join-Path $duplicateA 'duplicate.pdf'), (Join-Path $duplicateB 'duplicate.pdf')) -join ','
            $output = ''
            $duplicateFailed = $false
            try {
                $output = & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFiles $duplicatePaths -OutputFile $outputPath -NoConfirm 2>&1 | Out-String
            } catch {
                $duplicateFailed = $true
                $output = $_ | Out-String
            }

            if (-not $duplicateFailed -and $LASTEXITCODE -ne 0) {
                $duplicateFailed = $true
            }

            Assert-True -Condition $duplicateFailed -Message 'Duplicate filename run unexpectedly succeeded.'
            Assert-True -Condition (-not (Test-Path -LiteralPath $outputPath)) -Message 'Duplicate filename run should not create an output workbook.'
            return 'Duplicate PDF names are rejected'
        }
        'Broken PDF handling' {
            $outputPath = Join-Path $resultsRoot 'mixed_broken.xlsx'
            & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFolder $mixedPdfDir -OutputFile $outputPath -NoConfirm
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message 'Mixed output workbook was not created.'
            $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
            Assert-True -Condition ($snapshot.ResultRows -eq 5) -Message "Expected 5 Result rows for one valid PDF but got $($snapshot.ResultRows)."
            Assert-True -Condition ($snapshot.ErrorsRows -ge 2) -Message 'Broken PDF was not reported in Errors sheet.'
            Assert-True -Condition ($snapshot.ControlErrorCount -eq '1') -Message "Control sheet error count is unexpected: $($snapshot.ControlErrorCount)"
            Assert-True -Condition ($snapshot.ControlFailedCount -eq '1') -Message "Control sheet failed count is unexpected: $($snapshot.ControlFailedCount)"
            return "ResultRows=$($snapshot.ResultRows), ErrorsRows=$($snapshot.ErrorsRows)"
        }
        'Nested output path' {
            $nestedDir = Join-Path $resultsRoot 'nested\child\output'
            $outputPath = Join-Path $nestedDir 'nested_output.xlsx'
            if (Test-Path -LiteralPath $nestedDir) {
                Remove-Item -LiteralPath $nestedDir -Recurse -Force
            }
            & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFolder $validPdfDir -OutputFile $outputPath -NoConfirm
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message 'Nested output workbook was not created.'
            return 'Nested output created'
        }
        'Profile-based conversion' {
            $outputPath = Join-Path $resultsRoot 'profile10_success.xlsx'
            & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFolder $profileFixtureDir -ProfilePath $script:customProfilePath -OutputFile $outputPath -NoConfirm
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message 'Profile-based output workbook was not created.'
            $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
            Assert-True -Condition ($snapshot.ResultColumns -eq 11) -Message "Expected 11 columns for the 10-column profile but got $($snapshot.ResultColumns)."
            Assert-True -Condition ($snapshot.ControlProfile -eq 'Profile10 Test') -Message "Profile name was not written to Control: $($snapshot.ControlProfile)"
            Assert-True -Condition ($snapshot.SummarySuccessCount -eq '1') -Message "Summary success count is unexpected: $($snapshot.SummarySuccessCount)"
            return 'Custom profile conversion succeeded'
        }
        'Runtime cleanup' {
            $runtimeRuns = @(Get-ChildItem -LiteralPath (Join-Path $projectRoot 'output\runtime\runs') -Directory -ErrorAction SilentlyContinue)
            $runtimeFileNames = @($runtimeRuns | Select-Object -ExpandProperty Name)
            Assert-True -Condition ($runtimeRuns.Count -eq 0) -Message ('Runtime run workspaces remained: ' + ($runtimeFileNames -join ', '))
            return 'Runtime directory is clean'
        }
        'No Excel leak' {
            $excelIds = @(Wait-For-ExcelBaseline -BaselineIds @() -TimeoutSeconds 10)
            Assert-True -Condition ($excelIds.Count -eq 0) -Message ('Excel processes still running: ' + ($excelIds -join ', '))
            return 'No Excel process remains'
        }
        default {
            throw "Unknown test case: $Name"
        }
    }
}

function Invoke-IsolatedTestCase {
    param([Parameter(Mandatory = $true)]$Definition)

    $resultPath = Join-Path $resultsRoot ("case_" + (($Definition.Name -replace '[^A-Za-z0-9]+', '_').Trim('_')) + '.json')
    if (Test-Path -LiteralPath $resultPath) {
        Remove-Item -LiteralPath $resultPath -Force
    }

    $job = Start-Job -ScriptBlock {
        param($scriptPath, $caseName, $singleResultPath)
        & $scriptPath -CaseName $caseName -SingleResultPath $singleResultPath
    } -ArgumentList $script:selfPath, $Definition.Name, $resultPath

    if (-not ($job | Wait-Job -Timeout $Definition.TimeoutSeconds -ErrorAction SilentlyContinue)) {
        Stop-Job -Job $job -Force -ErrorAction SilentlyContinue
        return [pscustomobject]@{
            Name         = $Definition.Name
            Scenario     = $Definition.Scenario
            Status       = 'FAIL'
            DurationMs   = $Definition.TimeoutSeconds * 1000
            Details      = $null
            ErrorMessage = "Timed out after $($Definition.TimeoutSeconds) seconds."
        }
    }

    if (-not (Test-Path -LiteralPath $resultPath)) {
        return [pscustomobject]@{
            Name         = $Definition.Name
            Scenario     = $Definition.Scenario
            Status       = 'FAIL'
            DurationMs   = 0
            Details      = $null
            ErrorMessage = "Test result file was not created."
        }
    }

    return Get-Content -LiteralPath $resultPath -Raw -Encoding UTF8 | ConvertFrom-Json
}

if (-not [string]::IsNullOrWhiteSpace($CaseName)) {
    Initialize-TestFixtures
    $definition = Get-TestCases | Where-Object Name -eq $CaseName | Select-Object -First 1
    if ($null -eq $definition) {
        throw "Unknown test case: $CaseName"
    }

    $result = Invoke-TestCase -Name $definition.Name -Scenario $definition.Scenario -Body { Invoke-NamedScenario -Name $CaseName }
    if ($SingleResultPath) {
        $result | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $SingleResultPath -Encoding UTF8
    }

    if ($result.Status -eq 'FAIL') {
        Write-Error $result.ErrorMessage
    }

    exit $(if ($result.Status -eq 'PASS') { 0 } else { 1 })
}

Initialize-TestFixtures
$testResults = @()
foreach ($definition in Get-TestCases) {
    $testResults += Invoke-IsolatedTestCase -Definition $definition
}

$timeFinished = Get-Date

$summary = [pscustomobject]@{
    StartedAt  = $timeStarted
    FinishedAt = $timeFinished
    Results    = $testResults
}

$summary | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $jsonReportPath -Encoding UTF8
$reportMarkdown = New-ReportMarkdown -TestResults $testResults -StartedAt $timeStarted -FinishedAt $timeFinished
$reportMarkdown | Set-Content -LiteralPath $markdownReportPath -Encoding UTF8

$failed = @($testResults | Where-Object Status -eq 'FAIL')
if ($failed.Count -gt 0) {
    Write-Error ("Integration tests failed: " + ($failed.Name -join ', '))
}

Write-Host "統合テスト成功: $(@($testResults | Where-Object Status -eq 'PASS').Count) / $($testResults.Count)"
Write-Host "Markdown レポート: $markdownReportPath"
Write-Host "JSON レポート: $jsonReportPath"
