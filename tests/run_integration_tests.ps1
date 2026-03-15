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
$attendancePdfDir = Join-Path $fixturesRoot 'attendance_jp'
$sampleAttendancePdfDir = Join-Path $projectRoot 'samples\v1\pdf\attendance_jp'
$sampleSalesPdfDir = Join-Path $projectRoot 'samples\v1\pdf\sales_daily_jp'
$sampleInventoryPdfDir = Join-Path $projectRoot 'samples\v1\pdf\inventory_jp'
$sampleInquiryPdfDir = Join-Path $projectRoot 'samples\v1\pdf\inquiry_jp'
$sampleConstructionPocPdfDir = Join-Path $projectRoot 'samples\v2\pdf\construction_transfer_poc'
$resultsRoot = Join-Path $testsRoot 'results'
$reportsRoot = Join-Path $projectRoot 'reports'
$runScript = Join-Path $projectRoot 'scripts\run_pdf2excel.ps1'
$runScriptV1 = Join-Path $projectRoot 'scripts\run_pdf2excel_v1.ps1'
$runScriptV2 = Join-Path $projectRoot 'scripts\run_pdf2excel_v2.ps1'
$buildTemplateScript = Join-Path $projectRoot 'scripts\build_excel_template.ps1'
$buildSamplesScript = Join-Path $projectRoot 'scripts\build_sample_pdfs.ps1'
$batScript = Join-Path $projectRoot 'run_pdf2excel.bat'
$batScriptV1 = Join-Path $projectRoot 'run_pdf2excel_v1.bat'
$batScriptV2 = Join-Path $projectRoot 'run_pdf2excel_v2.bat'
$commonScript = Join-Path $projectRoot 'scripts\pdf2excel.common.ps1'
$templatePathV1 = Join-Path $projectRoot 'template\PDF2Excel_V1_Converter.xlsm'
$templatePathV2 = Join-Path $projectRoot 'template\PDF2Excel_V2_Converter.xlsm'
$jsonReportPath = Join-Path $resultsRoot 'integration-test-results.json'
$markdownReportPath = Join-Path $reportsRoot 'test-report.md'
$timeStarted = Get-Date

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

$suiteBaselineExcel = @(Get-ExcelProcessIds)
$script:customProfilePath = Join-Path $workRoot 'profile10.json'
$script:attendanceProfilePath = Join-Path $projectRoot 'config\profiles\v1\attendance_monthly_jp.json'
$script:salesProfilePath = Join-Path $projectRoot 'config\profiles\v1\sales_daily_jp.json'
$script:inventoryProfilePath = Join-Path $projectRoot 'config\profiles\v1\inventory_list_jp.json'
$script:inquiryProfilePath = Join-Path $projectRoot 'config\profiles\v1\inquiry_weekly_jp.json'
$script:constructionPocProfilePath = Join-Path $projectRoot 'config\profiles\v2\construction_transfer_poc.json'
$script:selfPath = $MyInvocation.MyCommand.Path

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

function New-JapaneseAttendancePdfFixture {
    param(
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)][string]$MonthLabel
    )

    $excel = $null
    $workbook = $null
    $worksheet = $null

    $memberRows = @(
        @('A001', '佐藤花子', '営業部', '通常'),
        @('A002', '鈴木一郎', '営業部', '在宅'),
        @('A003', '田中美咲', '管理部', '通常'),
        @('A004', '高橋健太', '管理部', '通常'),
        @('A005', '伊藤直子', '開発部', '在宅'),
        @('A006', '渡辺大輔', '開発部', '通常')
    )
    $statusCycle = @('出勤', '在宅', '休暇', '半休', '遅刻', '出勤', '出勤', '在宅')

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Name = '勤怠表'

        $headers = @('社員番号', '氏名', '所属', '勤務区分')
        $headers += @(1..30 | ForEach-Object { '{0}日' -f $_ })
        $headers += @('備考')

        for ($column = 1; $column -le $headers.Count; $column += 1) {
            $worksheet.Cells.Item(1, $column).Value2 = $headers[$column - 1]
        }

        for ($rowIndex = 0; $rowIndex -lt $memberRows.Count; $rowIndex += 1) {
            $excelRow = $rowIndex + 2
            $member = $memberRows[$rowIndex]
            for ($column = 1; $column -le 4; $column += 1) {
                $worksheet.Cells.Item($excelRow, $column).Value2 = $member[$column - 1]
            }

            for ($day = 1; $day -le 30; $day += 1) {
                $status = $statusCycle[($day + $rowIndex) % $statusCycle.Count]
                if ((($day + $rowIndex) % 5) -ne 0) {
                    $worksheet.Cells.Item($excelRow, $day + 4).Value2 = $status
                }
            }

            $worksheet.Cells.Item($excelRow, 35).Value2 = "$MonthLabel 月次確認済み"
        }

        $worksheet.Range('A1:AI1').Font.Bold = $true
        $worksheet.Columns.AutoFit() | Out-Null
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
    $reviewSheet = $null

    try {
        function Convert-SheetToTableSnapshot {
            param(
                [Parameter(Mandatory = $true)]$Worksheet,
                [int]$MaxRows = 120
            )

            $usedRange = $Worksheet.UsedRange
            $rowCount = [int]$usedRange.Rows.Count
            $columnCount = [int]$usedRange.Columns.Count
            $captureRows = [Math]::Min($rowCount, $MaxRows)
            $range = $Worksheet.Range($Worksheet.Cells.Item(1, 1), $Worksheet.Cells.Item($captureRows, $columnCount)).Value2
            $rows = @()
            $rawRows = @()

            if ($null -ne $range) {
                if ($range -is [System.Array]) {
                    for ($row = 1; $row -le $range.GetLength(0); $row += 1) {
                        $values = @()
                        $rawValues = @()
                        for ($column = 1; $column -le $range.GetLength(1); $column += 1) {
                            $rawValues += $range[$row, $column]
                            $values += [string]$range[$row, $column]
                        }
                        $rawRows += ,$rawValues
                        $rows += ,$values
                    }
                } else {
                    $rawRows += ,@($range)
                    $rows += ,@([string]$range)
                }
            }

            $headers = @()
            if ($rows.Count -gt 0) {
                $headers = @($rows[0])
            }

            $records = @()
            $typedRecords = @()
            for ($rowIndex = 1; $rowIndex -lt $rows.Count; $rowIndex += 1) {
                $record = [ordered]@{}
                $typedRecord = [ordered]@{}
                for ($columnIndex = 0; $columnIndex -lt $headers.Count; $columnIndex += 1) {
                    $header = $headers[$columnIndex]
                    if (-not [string]::IsNullOrWhiteSpace($header)) {
                        $record[$header] = if ($columnIndex -lt $rows[$rowIndex].Count) { $rows[$rowIndex][$columnIndex] } else { '' }
                        $typedRecord[$header] = if ($columnIndex -lt $rawRows[$rowIndex].Count) { $rawRows[$rowIndex][$columnIndex] } else { $null }
                    }
                }

                if ($record.Count -gt 0) {
                    $records += [pscustomobject]$record
                    $typedRecords += [pscustomobject]$typedRecord
                }
            }

            return [pscustomobject]@{
                Headers      = $headers
                Rows         = $rows
                Records      = $records
                TypedRecords = $typedRecords
            }
        }

        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($WorkbookPath)
        $controlSheet = $workbook.Worksheets.Item('Control')
        $summarySheet = $workbook.Worksheets.Item('Summary')
        $resultSheet = $workbook.Worksheets.Item('Result')
        $errorsSheet = $workbook.Worksheets.Item('Errors')
        try {
            $reviewSheet = $workbook.Worksheets.Item('Review')
        } catch {
            $reviewSheet = $null
        }

        $resultRows = [int]$resultSheet.UsedRange.Rows.Count
        $resultColumns = [int]$resultSheet.UsedRange.Columns.Count
        $errorsRows = [int]$errorsSheet.UsedRange.Rows.Count
        $errorsColumns = [int]$errorsSheet.UsedRange.Columns.Count
        $reviewRows = if ($reviewSheet) { [int]$reviewSheet.UsedRange.Rows.Count } else { 0 }
        $resultSnapshot = Convert-SheetToTableSnapshot -Worksheet $resultSheet
        $reviewSnapshot = if ($reviewSheet) { Convert-SheetToTableSnapshot -Worksheet $reviewSheet } else { [pscustomobject]@{ Headers = @(); Rows = @(); Records = @(); TypedRecords = @() } }

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
            ControlVersion      = [string]$controlSheet.Range('B15').Value2
            SummaryTitle        = [string]$summarySheet.Range('A1').Value2
            SummaryFileCount    = [string]$summarySheet.Range('B5').Value2
            SummarySuccessCount = [string]$summarySheet.Range('B6').Value2
            SummaryFailedCount  = [string]$summarySheet.Range('B7').Value2
            ReviewRows          = $reviewRows
            Sample              = $resultSnapshot.Rows
            ResultHeaders       = $resultSnapshot.Headers
            ResultRecords       = $resultSnapshot.Records
            ResultTypedRecords  = $resultSnapshot.TypedRecords
            ReviewHeaders       = $reviewSnapshot.Headers
            ReviewRecords       = $reviewSnapshot.Records
            ReviewTypedRecords  = $reviewSnapshot.TypedRecords
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

        foreach ($comObject in @($reviewSheet, $summarySheet, $controlSheet, $errorsSheet, $resultSheet, $workbook, $excel)) {
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
            RetryCount   = 0
            RetriedBy    = ''
        }
    } catch {
        $stopwatch.Stop()
        $leaked = @(Wait-For-ExcelBaseline -BaselineIds $baselineExcel -TimeoutSeconds 10)
        $errorMessage = $_.Exception.Message
        if ($leaked.Count -gt 0) {
            $errorMessage += " / Excel process leak detected: $($leaked -join ', ')"
        }
        return [pscustomobject]@{
            Name         = $Name
            Scenario     = $Scenario
            Status       = 'FAIL'
            DurationMs   = [int]$stopwatch.ElapsedMilliseconds
            Details      = $null
            ErrorMessage = $errorMessage
            RetryCount   = 0
            RetriedBy    = ''
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
    $retriedCount = @($TestResults | Where-Object { $_.RetryCount -gt 0 }).Count
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
        ('- 再試行発生: {0} 件' -f $retriedCount),
        ("- エラー有無: {0}" -f $(if ($failCount -eq 0) { 'なし' } else { 'あり' })),
        ("- 失敗概要: {0}" -f $(if ($failCount -eq 0) { 'なし' } else { '失敗シナリオ一覧を参照' })),
        '',
        '## シナリオ別結果',
        ''
    )

    $lines += '| No | シナリオ | 結果 | 所要時間 | 再試行 | 補足 |'
    $lines += '| --- | --- | --- | --- | --- | --- |'

    $index = 1
    foreach ($result in $TestResults) {
        $detailText = if ($result.ErrorMessage) { $result.ErrorMessage } elseif ($result.Details) { $result.Details } else { '' }
        $statusLabel = if ($result.Status -eq 'PASS') { '成功' } else { '失敗' }
        $retryLabel = if ([int]$result.RetryCount -gt 0) { "$($result.RetryCount)回 ($($result.RetriedBy))" } else { 'なし' }
        $lines += "| $index | $($result.Name) | $statusLabel | $($result.DurationMs) ms | $retryLabel | $detailText |"
        $index += 1
    }

    $lines += ''
    $lines += '## 総評'
    $lines += ''
    if ($failCount -eq 0) {
        $lines += '- 主要な正常系、異常系、運用系シナリオはすべて成功しました。'
        $lines += '- 各シナリオは子プロセス隔離とタイムアウト監視付きで実行されました。'
        $lines += '- 既知の Excel 一時失敗だけを 1 回まで再試行し、それ以外は即時失敗として扱いました。'
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
    Ensure-Directory -Path $resultsRoot
    Ensure-Directory -Path $reportsRoot
    foreach ($fixtureDir in @($validPdfDir, $mixedPdfDir, $duplicateA, $duplicateB, $profileFixtureDir, $japanesePdfDir, $bulkPdfDir, $attendancePdfDir)) {
        Ensure-Directory -Path $fixtureDir
    }

    Get-ChildItem -LiteralPath $resultsRoot -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like 'case_*.json' -or $_.Extension -eq '.xlsx' } |
        Remove-Item -Force -ErrorAction SilentlyContinue

    $requiredSamplePaths = @(
        (Join-Path $sampleAttendancePdfDir '2026年03月_勤怠管理表.pdf')
        (Join-Path $sampleSalesPdfDir '2026-03-15_売上日報_東京店.pdf')
        (Join-Path $sampleInventoryPdfDir '春季_在庫一覧_倉庫A.pdf')
        (Join-Path $sampleInquiryPdfDir '2026年03月_問い合わせ管理表_第1週.pdf')
        (Join-Path $sampleConstructionPocPdfDir '2026年02月_作業員勤怠一覧_PoC.pdf')
        (Join-Path $sampleConstructionPocPdfDir '2026年02月_作業員勤怠一覧_PoC_ヘッダー不一致負例.pdf')
    )
    $sampleFallbackMap = [ordered]@{
        (Join-Path $sampleAttendancePdfDir '2026年03月_勤怠管理表.pdf') = Join-Path $projectRoot 'samples\pdf\attendance_jp\2026年03月_勤怠管理表.pdf'
        (Join-Path $sampleSalesPdfDir '2026-03-15_売上日報_東京店.pdf') = Join-Path $projectRoot 'samples\pdf\sales_daily_jp\2026-03-15_売上日報_東京店.pdf'
        (Join-Path $sampleInventoryPdfDir '春季_在庫一覧_倉庫A.pdf') = Join-Path $projectRoot 'samples\pdf\inventory_jp\春季_在庫一覧_倉庫A.pdf'
        (Join-Path $sampleInquiryPdfDir '2026年03月_問い合わせ管理表_第1週.pdf') = Join-Path $projectRoot 'samples\pdf\inquiry_jp\2026年03月_問い合わせ管理表_第1週.pdf'
        (Join-Path $sampleConstructionPocPdfDir '2026年02月_作業員勤怠一覧_PoC.pdf') = Join-Path $projectRoot 'samples\pdf\construction_transfer_poc\2026年02月_作業員勤怠一覧_PoC.pdf'
        (Join-Path $sampleConstructionPocPdfDir '2026年02月_作業員勤怠一覧_PoC_2ページ同一列.pdf') = Join-Path $projectRoot 'samples\pdf\construction_transfer_poc\2026年02月_作業員勤怠一覧_PoC_2ページ同一列.pdf'
        (Join-Path $sampleConstructionPocPdfDir '2026年02月_作業員勤怠一覧_PoC_ヘッダー不一致負例.pdf') = Join-Path $projectRoot 'samples\pdf\construction_transfer_poc\2026年02月_作業員勤怠一覧_PoC_ヘッダー不一致負例.pdf'
        (Join-Path $sampleConstructionPocPdfDir '2026年02月_作業員勤怠一覧_PoC_時刻確認負例.pdf') = Join-Path $projectRoot 'samples\pdf\construction_transfer_poc\2026年02月_作業員勤怠一覧_PoC_時刻確認負例.pdf'
        (Join-Path $sampleConstructionPocPdfDir '2026年04月-06月_作業員勤怠一覧_PoC_6ページ同一列.pdf') = Join-Path $projectRoot 'samples\pdf\construction_transfer_poc\2026年04月-06月_作業員勤怠一覧_PoC_6ページ同一列.pdf'
    }
    foreach ($destinationPath in $sampleFallbackMap.Keys) {
        $sourcePath = $sampleFallbackMap[$destinationPath]
        if ((-not (Test-Path -LiteralPath $destinationPath)) -and (Test-Path -LiteralPath $sourcePath)) {
            Ensure-Directory -Path (Split-Path -Parent $destinationPath)
            Copy-Item -LiteralPath $sourcePath -Destination $destinationPath -Force
        }
    }
    $needsSampleRebuild = (@($requiredSamplePaths | Where-Object { -not (Test-Path -LiteralPath $_) }).Count) -gt 0
    if ($needsSampleRebuild) {
        [void](Wait-For-ExcelBaseline -BaselineIds $suiteBaselineExcel -TimeoutSeconds 15)
        for ($attempt = 0; $attempt -lt 2; $attempt += 1) {
            $buildResult = Invoke-TestProcess -FilePath 'powershell' -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $buildSamplesScript) -TimeoutSeconds 240
            if ($buildResult.ExitCode -eq 0) {
                break
            }

            $combinedOutput = ($buildResult.StdOut + ' ' + $buildResult.StdErr)
            $isRetryableBuildFailure =
                $attempt -eq 0 -and (
                    $combinedOutput -match 'being used by another process' -or
                    $combinedOutput -match '別のプロセス' -or
                    $combinedOutput -match 'process cannot access the file'
                )
            if ($isRetryableBuildFailure) {
                [void](Wait-For-ExcelBaseline -BaselineIds $suiteBaselineExcel -TimeoutSeconds 15)
                Start-Sleep -Seconds 3
                continue
            }

            throw "サンプル PDF 再生成に失敗しました: ExitCode=$($buildResult.ExitCode) / $combinedOutput"
        }
    }

    New-ExcelPdfFixture -OutputPath (Join-Path $validPdfDir 'valid_a.pdf') -Prefix 'VALIDA'
    New-ExcelPdfFixture -OutputPath (Join-Path $validPdfDir 'valid_b.pdf') -Prefix 'VALIDB'
    Copy-Item -LiteralPath (Join-Path $validPdfDir 'valid_a.pdf') -Destination (Join-Path $mixedPdfDir 'valid_a.pdf') -Force
    Set-Content -LiteralPath (Join-Path $mixedPdfDir 'broken.pdf') -Value 'not-a-real-pdf' -Encoding ASCII
    New-ExcelPdfFixture -OutputPath (Join-Path $duplicateA 'duplicate.pdf') -Prefix 'DUPA'
    New-ExcelPdfFixture -OutputPath (Join-Path $duplicateB 'duplicate.pdf') -Prefix 'DUPB'
    New-ExcelPdfFixture -OutputPath (Join-Path $profileFixtureDir 'profile10.pdf') -Prefix 'P10' -Columns 10
    Copy-Item -LiteralPath (Join-Path $validPdfDir 'valid_a.pdf') -Destination (Join-Path $japanesePdfDir '日本語_帳票A.pdf') -Force
    Copy-Item -LiteralPath (Join-Path $validPdfDir 'valid_b.pdf') -Destination (Join-Path $japanesePdfDir '請求書_テストB.pdf') -Force
    New-JapaneseAttendancePdfFixture -OutputPath (Join-Path $attendancePdfDir '2026年03月_勤怠管理表.pdf') -MonthLabel '2026年03月'
    New-JapaneseAttendancePdfFixture -OutputPath (Join-Path $attendancePdfDir '2026年04月_勤怠管理表.pdf') -MonthLabel '2026年04月'
    for ($index = 1; $index -le 50; $index += 1) {
        $sourceName = if (($index % 2) -eq 0) { 'valid_b.pdf' } else { 'valid_a.pdf' }
        $bulkName = 'bulk_{0:D2}.pdf' -f $index
        Copy-Item -LiteralPath (Join-Path $validPdfDir $sourceName) -Destination (Join-Path $bulkPdfDir $bulkName) -Force
    }

    $customProfileJson = @'
{
  "name": "profile10",
  "displayName": "10列テストプロファイル",
  "description": "10 列帳票の統合テスト用プロファイルです。",
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
        [pscustomobject]@{ Name = 'テンプレート再生成'; Scenario = 'テンプレート再生成が成功し、xlsm の更新日時が進むこと'; TimeoutSeconds = 120 },
        [pscustomobject]@{ Name = 'PowerShell 経由の正常変換'; Scenario = '有効な PDF 2 件を PowerShell から変換し、Result 行数を確認すること'; TimeoutSeconds = 180 },
        [pscustomobject]@{ Name = '単票 PDF の変換'; Scenario = '1 件だけの PDF フォルダでも正常に変換できること'; TimeoutSeconds = 180 },
        [pscustomobject]@{ Name = 'InputFiles 指定の変換'; Scenario = '公開インターフェースの -InputFiles で複数 PDF を正しく処理できること'; TimeoutSeconds = 180 },
        [pscustomobject]@{ Name = '日本語ファイル名の変換'; Scenario = '日本語ファイル名が Result と Summary にそのまま残ること'; TimeoutSeconds = 180 },
        [pscustomobject]@{ Name = '50件一括変換性能'; Scenario = '50 件の PDF を許容時間内に変換し、行数が崩れないこと'; TimeoutSeconds = 540 },
        [pscustomobject]@{ Name = 'KeepInput の隔離動作'; Scenario = 'KeepInput を使っても今回分だけが専用 staging で処理されること'; TimeoutSeconds = 180 },
        [pscustomobject]@{ Name = '同時実行ロック'; Scenario = '別実行中は 2 本目が即時失敗すること'; TimeoutSeconds = 180 },
        [pscustomobject]@{ Name = 'BAT 経由の変換'; Scenario = '同じ PDF 群を BAT から正常に変換できること'; TimeoutSeconds = 180 },
        [pscustomobject]@{ Name = 'BAT 直実行で待機しない'; Scenario = '引数付き BAT 実行で pause せず終了すること'; TimeoutSeconds = 180 },
        [pscustomobject]@{ Name = 'input 自己参照'; Scenario = 'input 自体を入力フォルダにしても自己削除せず処理できること'; TimeoutSeconds = 180 },
        [pscustomobject]@{ Name = '同名ファイル拒否'; Scenario = '別フォルダの同名 PDF を明示的に拒否すること'; TimeoutSeconds = 120 },
        [pscustomobject]@{ Name = '壊れた PDF の処理'; Scenario = '壊れた PDF が全体を止めず Errors に出ること'; TimeoutSeconds = 180 },
        [pscustomobject]@{ Name = '深い出力先パス'; Scenario = '深いフォルダ階層の保存先でも Excel を出力できること'; TimeoutSeconds = 180 },
        [pscustomobject]@{ Name = 'プロファイル切替変換'; Scenario = 'カスタムプロファイルで列数と Summary が切り替わること'; TimeoutSeconds = 180 },
        [pscustomobject]@{ Name = '日本語勤怠管理表の変換'; Scenario = '6 人分の月次勤怠管理表を日本語プロファイルで正しく変換できること'; TimeoutSeconds = 240 },
        [pscustomobject]@{ Name = '日本語売上日報の変換'; Scenario = '店舗別の売上日報を日本語プロファイルで正しく変換できること'; TimeoutSeconds = 240 },
        [pscustomobject]@{ Name = '日本語在庫一覧の変換'; Scenario = '倉庫別の在庫一覧を日本語プロファイルで正しく変換できること'; TimeoutSeconds = 240 },
        [pscustomobject]@{ Name = '日本語問い合わせ管理表の変換'; Scenario = '週次の問い合わせ管理表を日本語プロファイルで正しく変換できること'; TimeoutSeconds = 240 },
        [pscustomobject]@{ Name = '建設現場転記PoCの変換'; Scenario = '改行セルや時刻ゆれを含む建設現場向けPoC帳票を raw 転記できること'; TimeoutSeconds = 300 },
        [pscustomobject]@{ Name = 'V2 2ページ同一列の変換'; Scenario = '同一列ヘッダーの2ページ建設帳票を1つの Result に連結できること'; TimeoutSeconds = 360 },
        [pscustomobject]@{ Name = 'V2 6ページ同一列の変換'; Scenario = '同一列ヘッダーの6ページ建設帳票を1つの Result に連結できること'; TimeoutSeconds = 420 },
        [pscustomobject]@{ Name = 'V2 ヘッダー不一致負例の分離'; Scenario = 'sameHeader に乗らない multi-page 帳票を Errors 側へ分離できること'; TimeoutSeconds = 360 },
        [pscustomobject]@{ Name = 'V2 時刻確認負例の分離'; Scenario = '第2時刻ペアの invalid / 片側空を Review で拾えること'; TimeoutSeconds = 240 },
        [pscustomobject]@{ Name = '一時領域の後片付け'; Scenario = '実行後に output/runtime/runs 配下へ残骸が残らないこと'; TimeoutSeconds = 60 },
        [pscustomobject]@{ Name = 'Excel プロセス残留なし'; Scenario = 'スイート完了後に余分な EXCEL.exe が残らないこと'; TimeoutSeconds = 60 }
    )
}

function Invoke-NamedScenario {
    param([Parameter(Mandatory = $true)][string]$Name)

    switch ($Name) {
        'テンプレート再生成' {
            $beforeV1 = if (Test-Path -LiteralPath $templatePathV1) { (Get-Item -LiteralPath $templatePathV1).LastWriteTimeUtc } else { $null }
            $beforeV2 = if (Test-Path -LiteralPath $templatePathV2) { (Get-Item -LiteralPath $templatePathV2).LastWriteTimeUtc } else { $null }
            & powershell -NoProfile -ExecutionPolicy Bypass -File $buildTemplateScript -TemplatePath $templatePathV1 -TemplateVariant v1
            & powershell -NoProfile -ExecutionPolicy Bypass -File $buildTemplateScript -TemplatePath $templatePathV2 -TemplateVariant v2
            Assert-True -Condition (Test-Path -LiteralPath $templatePathV1) -Message 'V1 テンプレートファイルが作成されていません。'
            Assert-True -Condition (Test-Path -LiteralPath $templatePathV2) -Message 'V2 テンプレートファイルが作成されていません。'
            $afterV1 = (Get-Item -LiteralPath $templatePathV1).LastWriteTimeUtc
            $afterV2 = (Get-Item -LiteralPath $templatePathV2).LastWriteTimeUtc
            Assert-True -Condition ($null -eq $beforeV1 -or $afterV1 -ge $beforeV1) -Message 'V1 テンプレートの更新日時が進んでいません。'
            Assert-True -Condition ($null -eq $beforeV2 -or $afterV2 -ge $beforeV2) -Message 'V2 テンプレートの更新日時が進んでいません。'
            return "V1=$afterV1 / V2=$afterV2"
        }
        'PowerShell 経由の正常変換' {
            $outputPath = Join-Path $resultsRoot 'powershell_success.xlsx'
            & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFolder $validPdfDir -OutputFile $outputPath -NoConfirm
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message 'PowerShell 実行の出力ブックが作成されていません。'
            $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
            Assert-True -Condition ($snapshot.ResultColumns -eq 31) -Message "列数が想定と異なります: $($snapshot.ResultColumns)"
            Assert-True -Condition ($snapshot.ResultRows -eq 9) -Message "Result 行数が想定と異なります: $($snapshot.ResultRows)"
            Assert-True -Condition ($snapshot.ErrorsRows -eq 1) -Message "Errors 行数が想定と異なります: $($snapshot.ErrorsRows)"
            Assert-True -Condition ($snapshot.Sample[1][0] -eq 'valid_a.pdf') -Message 'A列の元ファイル名が想定と異なります。'
            Assert-True -Condition ($snapshot.ControlSourceCount -eq '2') -Message "Control の対象 PDF 数が想定と異なります: $($snapshot.ControlSourceCount)"
            Assert-True -Condition ($snapshot.ControlResultCount -eq '8') -Message "Control の取込データ行数が想定と異なります: $($snapshot.ControlResultCount)"
            Assert-True -Condition ($snapshot.ControlErrorCount -eq '0') -Message "Control のエラー件数が想定と異なります: $($snapshot.ControlErrorCount)"
            Assert-True -Condition ($snapshot.ControlSuccessCount -eq '2') -Message "Control の成功 PDF 数が想定と異なります: $($snapshot.ControlSuccessCount)"
            Assert-True -Condition ($snapshot.ControlFailedCount -eq '0') -Message "Control の失敗 PDF 数が想定と異なります: $($snapshot.ControlFailedCount)"
            Assert-True -Condition (-not [string]::IsNullOrWhiteSpace($snapshot.ControlProfile)) -Message 'Control のプロファイル表示が空です。'
            Assert-True -Condition ($snapshot.SummaryTitle -eq 'Summary') -Message 'Summary シートのタイトルが不足しています。'
            Assert-True -Condition ($snapshot.SummaryFileCount -eq '2') -Message "Summary の対象 PDF 数が想定と異なります: $($snapshot.SummaryFileCount)"
            return "Result 行数=$($snapshot.ResultRows), Errors 行数=$($snapshot.ErrorsRows)"
        }
        '単票 PDF の変換' {
            $singlePdfDir = Join-Path $fixturesRoot 'single'
            Reset-Directory -Path $singlePdfDir
            Copy-Item -LiteralPath (Join-Path $validPdfDir 'valid_a.pdf') -Destination (Join-Path $singlePdfDir 'valid_a.pdf') -Force
            $outputPath = Join-Path $resultsRoot 'single_pdf_success.xlsx'
            & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFolder $singlePdfDir -OutputFile $outputPath -NoConfirm
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message '単票 PDF の出力ブックが作成されていません。'
            $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
            Assert-True -Condition ($snapshot.ResultRows -eq 5) -Message "単票 PDF の Result 行数が想定と異なります: $($snapshot.ResultRows)"
            Assert-True -Condition ($snapshot.ControlSourceCount -eq '1') -Message "Control の対象 PDF 数が想定と異なります: $($snapshot.ControlSourceCount)"
            return '単票 PDF の変換に成功'
        }
        'InputFiles 指定の変換' {
            $outputPath = Join-Path $resultsRoot 'inputfiles_success.xlsx'
            $inputFilesArg = @((Join-Path $validPdfDir 'valid_a.pdf'), (Join-Path $validPdfDir 'valid_b.pdf')) -join ','
            & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFiles $inputFilesArg -OutputFile $outputPath -NoConfirm
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message 'InputFiles の出力ブックが作成されていません。'
            $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
            Assert-True -Condition ($snapshot.ResultRows -eq 9) -Message "InputFiles 指定の Result 行数が想定と異なります: $($snapshot.ResultRows)"
            Assert-True -Condition ($snapshot.ControlSourceCount -eq '2') -Message "Control の対象 PDF 数が想定と異なります: $($snapshot.ControlSourceCount)"
            return 'InputFiles 指定の変換に成功'
        }
        '日本語ファイル名の変換' {
            $outputPath = Join-Path $resultsRoot 'japanese_success.xlsx'
            & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFolder $japanesePdfDir -OutputFile $outputPath -NoConfirm
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message '日本語ファイル名テストの出力ブックが作成されていません。'
            $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
            $sourceNames = @($snapshot.Sample | Select-Object -Skip 1 | ForEach-Object { $_[0] } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
            Assert-True -Condition ($snapshot.ControlSourceCount -eq '2') -Message "Control の対象 PDF 数が想定と異なります: $($snapshot.ControlSourceCount)"
            Assert-True -Condition ($sourceNames -contains '日本語_帳票A.pdf') -Message '日本語ファイル名 A が Result に保持されていません。'
            Assert-True -Condition ($sourceNames -contains '請求書_テストB.pdf') -Message '日本語ファイル名 B が Result に保持されていません。'
            return "元ファイル名=$($sourceNames -join ',')"
        }
        '50件一括変換性能' {
            $outputPath = Join-Path $resultsRoot 'bulk50_success.xlsx'
            $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFolder $bulkPdfDir -OutputFile $outputPath -NoConfirm
            $stopwatch.Stop()
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message '50件一括変換の出力ブックが作成されていません。'
            $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
            Assert-True -Condition ($snapshot.ControlSourceCount -eq '50') -Message "対象 PDF 数が 50 件になっていません: $($snapshot.ControlSourceCount)"
            Assert-True -Condition ($snapshot.ResultRows -eq 201) -Message "50 件一括変換の Result 行数が想定と異なります: $($snapshot.ResultRows)"
            Assert-True -Condition ($stopwatch.Elapsed.TotalSeconds -lt 300) -Message ("50 件一括変換の処理時間が長すぎます: {0:N1} 秒" -f $stopwatch.Elapsed.TotalSeconds)
            return ("処理秒数={0:N1}, Result 行数={1}" -f $stopwatch.Elapsed.TotalSeconds, $snapshot.ResultRows)
        }
        'KeepInput の隔離動作' {
            $inputStore = Join-Path $projectRoot 'input'
            Get-ChildItem -LiteralPath $inputStore -Filter '*.pdf' -File -ErrorAction SilentlyContinue | Remove-Item -Force
            Copy-Item -LiteralPath (Join-Path $validPdfDir 'valid_b.pdf') -Destination (Join-Path $inputStore 'archived_valid_b.pdf') -Force

            $singleKeepDir = Join-Path $fixturesRoot 'single_keep'
            Reset-Directory -Path $singleKeepDir
            Copy-Item -LiteralPath (Join-Path $validPdfDir 'valid_a.pdf') -Destination (Join-Path $singleKeepDir 'valid_a.pdf') -Force

            $outputPath = Join-Path $resultsRoot 'keepinput_isolation.xlsx'
            & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFolder $singleKeepDir -OutputFile $outputPath -KeepInput -NoConfirm
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message 'KeepInput テストの出力ブックが作成されていません。'
            $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
            Assert-True -Condition ($snapshot.ResultRows -eq 5) -Message "KeepInput 実行で今回対象外の PDF が混ざっています: $($snapshot.ResultRows)"

            $storedNames = @(Get-ChildItem -LiteralPath $inputStore -Filter '*.pdf' -File | Select-Object -ExpandProperty Name)
            Assert-True -Condition ($storedNames -contains 'archived_valid_b.pdf') -Message 'KeepInput で保管 PDF を保持できていません。'
            Assert-True -Condition ($storedNames -contains 'valid_a.pdf') -Message 'KeepInput で今回対象 PDF を input に保管できていません。'
            return "保管中ファイル=$($storedNames -join ',')"
        }
        '同時実行ロック' {
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
                Assert-True -Condition $lockAcquired -Message 'テスト用 mutex を取得できませんでした。'

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

                Assert-True -Condition $secondFailed -Message 'ロック中なのに 2 本目の実行が成功してしまいました。'
                Assert-True -Condition (-not (Test-Path -LiteralPath $secondOutput)) -Message 'ロック中なのに 2 本目の出力ブックが作成されました。'
                return '同時実行ロックにより 2 本目を拒否'
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
        'BAT 経由の変換' {
            $outputPath = Join-Path $resultsRoot 'bat_success.xlsx'
            & cmd /c $batScriptV1 -InputFolder $validPdfDir -OutputFile $outputPath -NoConfirm
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message 'BAT 実行の出力ブックが作成されていません。'
            $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
            Assert-True -Condition ($snapshot.ResultRows -eq 9) -Message "BAT 実行の Result 行数が想定と異なります: $($snapshot.ResultRows)"
            return "Result 行数=$($snapshot.ResultRows)"
        }
        'BAT 直実行で待機しない' {
            $outputPath = Join-Path $resultsRoot 'bat_direct_no_pause.xlsx'
            & cmd /c $batScriptV1 -InputFolder $validPdfDir -OutputFile $outputPath -NoConfirm
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message 'BAT 直実行の出力ブックが作成されていません。'
            return 'BAT 直実行が待機せず終了'
        }
        'input 自己参照' {
            Get-ChildItem -LiteralPath (Join-Path $projectRoot 'input') -Filter '*.pdf' -File -ErrorAction SilentlyContinue | Remove-Item -Force
            Copy-Item -LiteralPath (Join-Path $validPdfDir 'valid_a.pdf') -Destination (Join-Path $projectRoot 'input\valid_a.pdf') -Force
            Copy-Item -LiteralPath (Join-Path $validPdfDir 'valid_b.pdf') -Destination (Join-Path $projectRoot 'input\valid_b.pdf') -Force
            $outputPath = Join-Path $resultsRoot 'input_self_reference.xlsx'
            & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFolder (Join-Path $projectRoot 'input') -OutputFile $outputPath -NoConfirm
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message 'input 自己参照の出力ブックが作成されていません。'
            $remainingNames = @(Get-ChildItem -LiteralPath (Join-Path $projectRoot 'input') -Filter '*.pdf' -File | Select-Object -ExpandProperty Name)
            Assert-True -Condition ($remainingNames -contains 'valid_a.pdf') -Message 'input 自己参照の実行で valid_a.pdf が消えました。'
            Assert-True -Condition ($remainingNames -contains 'valid_b.pdf') -Message 'input 自己参照の実行で valid_b.pdf が消えました。'
            return "残存入力ファイル=$($remainingNames -join ',')"
        }
        '同名ファイル拒否' {
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

            Assert-True -Condition $duplicateFailed -Message '同名ファイルの実行が成功してしまいました。'
            Assert-True -Condition (-not (Test-Path -LiteralPath $outputPath)) -Message '同名ファイル実行で出力ブックが作成されてはいけません。'
            return '同名 PDF を拒否できた'
        }
        '壊れた PDF の処理' {
            $outputPath = Join-Path $resultsRoot 'mixed_broken.xlsx'
            & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFolder $mixedPdfDir -OutputFile $outputPath -NoConfirm
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message '壊れた PDF 混在テストの出力ブックが作成されていません。'
            $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
            Assert-True -Condition ($snapshot.ResultRows -eq 5) -Message "壊れた PDF 混在時の Result 行数が想定と異なります: $($snapshot.ResultRows)"
            Assert-True -Condition ($snapshot.ErrorsRows -ge 2) -Message '壊れた PDF が Errors シートに出ていません。'
            Assert-True -Condition ($snapshot.ControlErrorCount -eq '1') -Message "Control のエラー件数が想定と異なります: $($snapshot.ControlErrorCount)"
            Assert-True -Condition ($snapshot.ControlFailedCount -eq '1') -Message "Control の失敗 PDF 数が想定と異なります: $($snapshot.ControlFailedCount)"
            return "Result 行数=$($snapshot.ResultRows), Errors 行数=$($snapshot.ErrorsRows)"
        }
        '深い出力先パス' {
            $nestedDir = Join-Path $resultsRoot 'nested\child\output'
            $outputPath = Join-Path $nestedDir 'nested_output.xlsx'
            if (Test-Path -LiteralPath $nestedDir) {
                Remove-Item -LiteralPath $nestedDir -Recurse -Force
            }
            & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFolder $validPdfDir -OutputFile $outputPath -NoConfirm
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message '深い保存先の出力ブックが作成されていません。'
            return '深い保存先への出力に成功'
        }
        'プロファイル切替変換' {
            $outputPath = Join-Path $resultsRoot 'profile10_success.xlsx'
            & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFolder $profileFixtureDir -ProfilePath $script:customProfilePath -OutputFile $outputPath -NoConfirm
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message 'プロファイル切替の出力ブックが作成されていません。'
            $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
            Assert-True -Condition ($snapshot.ResultColumns -eq 11) -Message "10 列プロファイルの列数が想定と異なります: $($snapshot.ResultColumns)"
            Assert-True -Condition ($snapshot.ControlProfile -eq '10列テストプロファイル') -Message "Control のプロファイル表示が想定と異なります: $($snapshot.ControlProfile)"
            Assert-True -Condition ($snapshot.SummarySuccessCount -eq '1') -Message "Summary の成功 PDF 数が想定と異なります: $($snapshot.SummarySuccessCount)"
            return 'カスタムプロファイルでの変換に成功'
        }
        '日本語勤怠管理表の変換' {
            $outputPath = Join-Path $resultsRoot 'attendance_monthly_jp.xlsx'
            & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFolder $sampleAttendancePdfDir -ProfilePath $script:attendanceProfilePath -OutputFile $outputPath -NoConfirm
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message '日本語勤怠管理表テストの出力ブックが作成されていません。'
            $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
            $sourceNames = @($snapshot.Sample | Select-Object -Skip 1 | ForEach-Object { $_[0] } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
            $memberNames = @($snapshot.Sample | Select-Object -Skip 1 | ForEach-Object { $_[2] } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
            Assert-True -Condition ($snapshot.ResultColumns -eq 36) -Message "日本語勤怠管理表の列数が想定と異なります: $($snapshot.ResultColumns)"
            Assert-True -Condition ($snapshot.ResultRows -eq 13) -Message "日本語勤怠管理表の行数が想定と異なります: $($snapshot.ResultRows)"
            Assert-True -Condition ($snapshot.ControlProfile -eq '日本語勤怠管理表プロファイル') -Message "Control のプロファイル表示が想定と異なります: $($snapshot.ControlProfile)"
            Assert-True -Condition ($sourceNames -contains '2026年03月_勤怠管理表.pdf') -Message '3 月の勤怠 PDF 名が保持されていません。'
            Assert-True -Condition ($sourceNames -contains '2026年04月_勤怠管理表.pdf') -Message '4 月の勤怠 PDF 名が保持されていません。'
            Assert-True -Condition ($memberNames -contains '佐藤花子') -Message '勤怠管理表の氏名が保持されていません。'
            Assert-True -Condition ($memberNames -contains '渡辺大輔') -Message '勤怠管理表の氏名が十分に読み込めていません。'
            return "勤怠PDF=$($sourceNames -join ','), 氏名=$($memberNames -join ',')"
        }
        '日本語売上日報の変換' {
            $outputPath = Join-Path $resultsRoot 'sales_daily_jp.xlsx'
            & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFolder $sampleSalesPdfDir -ProfilePath $script:salesProfilePath -OutputFile $outputPath -NoConfirm
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message '日本語売上日報テストの出力ブックが作成されていません。'
            $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
            $sourceNames = @($snapshot.Sample | Select-Object -Skip 1 | ForEach-Object { $_[0] } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
            $storeNames = @($snapshot.Sample | Select-Object -Skip 1 | ForEach-Object { $_[2] } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
            Assert-True -Condition ($snapshot.ResultColumns -eq 13) -Message "日本語売上日報の列数が想定と異なります: $($snapshot.ResultColumns)"
            Assert-True -Condition ($snapshot.ResultRows -eq 11) -Message "日本語売上日報の行数が想定と異なります: $($snapshot.ResultRows)"
            Assert-True -Condition ($snapshot.ControlProfile -eq '日本語売上日報プロファイル') -Message "Control のプロファイル表示が想定と異なります: $($snapshot.ControlProfile)"
            Assert-True -Condition ($sourceNames -contains '2026-03-15_売上日報_東京店.pdf') -Message '東京店の売上日報 PDF 名が保持されていません。'
            Assert-True -Condition ($sourceNames -contains '2026-03-16_売上日報_横浜店.pdf') -Message '横浜店の売上日報 PDF 名が保持されていません。'
            Assert-True -Condition ($storeNames -contains '東京店') -Message '売上日報の店舗名が保持されていません。'
            Assert-True -Condition ($storeNames -contains '横浜店') -Message '売上日報の複数店舗が十分に読み込めていません。'
            return "売上PDF=$($sourceNames -join ','), 店舗=$($storeNames -join ',')"
        }
        '日本語在庫一覧の変換' {
            $outputPath = Join-Path $resultsRoot 'inventory_list_jp.xlsx'
            & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFolder $sampleInventoryPdfDir -ProfilePath $script:inventoryProfilePath -OutputFile $outputPath -NoConfirm
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message '日本語在庫一覧テストの出力ブックが作成されていません。'
            $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
            $sourceNames = @($snapshot.Sample | Select-Object -Skip 1 | ForEach-Object { $_[0] } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
            $warehouseNames = @($snapshot.Sample | Select-Object -Skip 1 | ForEach-Object { $_[4] } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
            Assert-True -Condition ($snapshot.ResultColumns -eq 11) -Message "日本語在庫一覧の列数が想定と異なります: $($snapshot.ResultColumns)"
            Assert-True -Condition ($snapshot.ResultRows -eq 13) -Message "日本語在庫一覧の行数が想定と異なります: $($snapshot.ResultRows)"
            Assert-True -Condition ($snapshot.ControlProfile -eq '日本語在庫一覧プロファイル') -Message "Control のプロファイル表示が想定と異なります: $($snapshot.ControlProfile)"
            Assert-True -Condition ($sourceNames -contains '春季_在庫一覧_倉庫A.pdf') -Message '倉庫Aの在庫一覧 PDF 名が保持されていません。'
            Assert-True -Condition ($sourceNames -contains '春季_在庫一覧_倉庫B.pdf') -Message '倉庫Bの在庫一覧 PDF 名が保持されていません。'
            Assert-True -Condition ($warehouseNames -contains '倉庫A') -Message '在庫一覧の倉庫名が保持されていません。'
            Assert-True -Condition ($warehouseNames -contains '倉庫B') -Message '在庫一覧の複数倉庫が十分に読み込めていません。'
            return "在庫PDF=$($sourceNames -join ','), 倉庫=$($warehouseNames -join ',')"
        }
        '日本語問い合わせ管理表の変換' {
            $outputPath = Join-Path $resultsRoot 'inquiry_weekly_jp.xlsx'
            & powershell -NoProfile -ExecutionPolicy Bypass -File $runScript -InputFolder $sampleInquiryPdfDir -ProfilePath $script:inquiryProfilePath -OutputFile $outputPath -NoConfirm
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message '日本語問い合わせ管理表テストの出力ブックが作成されていません。'
            $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
            $sourceNames = @($snapshot.Sample | Select-Object -Skip 1 | ForEach-Object { $_[0] } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
            $customerNames = @($snapshot.Sample | Select-Object -Skip 1 | ForEach-Object { $_[3] } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
            Assert-True -Condition ($snapshot.ResultColumns -eq 10) -Message "日本語問い合わせ管理表の列数が想定と異なります: $($snapshot.ResultColumns)"
            Assert-True -Condition ($snapshot.ResultRows -eq 11) -Message "日本語問い合わせ管理表の行数が想定と異なります: $($snapshot.ResultRows)"
            Assert-True -Condition ($snapshot.ControlProfile -eq '日本語問い合わせ管理表プロファイル') -Message "Control のプロファイル表示が想定と異なります: $($snapshot.ControlProfile)"
            Assert-True -Condition ($sourceNames -contains '2026年03月_問い合わせ管理表_第1週.pdf') -Message '第1週の問い合わせ管理表 PDF 名が保持されていません。'
            Assert-True -Condition ($sourceNames -contains '2026年03月_問い合わせ管理表_第2週.pdf') -Message '第2週の問い合わせ管理表 PDF 名が保持されていません。'
            Assert-True -Condition ($customerNames -contains '株式会社青葉') -Message '問い合わせ管理表の顧客名が保持されていません。'
            Assert-True -Condition ($customerNames -contains '北辰物流') -Message '問い合わせ管理表の複数顧客が十分に読み込めていません。'
            return "問い合わせPDF=$($sourceNames -join ','), 顧客=$($customerNames -join ',')"
        }
        '建設現場転記PoCの変換' {
            $outputPath = Join-Path $resultsRoot 'construction_transfer_poc.xlsx'
            $singlePdfDir = Join-Path $fixturesRoot 'construction_single_v2'
            Reset-Directory -Path $singlePdfDir
            Copy-Item -LiteralPath (Join-Path $sampleConstructionPocPdfDir '2026年02月_作業員勤怠一覧_PoC.pdf') -Destination (Join-Path $singlePdfDir '2026年02月_作業員勤怠一覧_PoC.pdf') -Force
            & powershell -NoProfile -ExecutionPolicy Bypass -File $runScriptV2 -InputFolder $singlePdfDir -ProfilePath $script:constructionPocProfilePath -OutputFile $outputPath -NoConfirm
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message '建設現場転記PoCテストの出力ブックが作成されていません。'
            $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
            $sourceNames = @($snapshot.Sample | Select-Object -Skip 1 | ForEach-Object { $_[0] } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
            $names = @($snapshot.Sample | Select-Object -Skip 1 | ForEach-Object { $_[3] } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
            $sites = @($snapshot.Sample | Select-Object -Skip 1 | ForEach-Object { $_[5] } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
            $normalizedInMinutes = @($snapshot.ResultRecords | ForEach-Object { $_.'正規化入場1_分' } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
            $normalizedOutMinutes = @($snapshot.ResultRecords | ForEach-Object { $_.'正規化退場1_分' } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
            $normalizedSecondOutMinutes = @($snapshot.ResultRecords | ForEach-Object { $_.'正規化退場2_分' } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
            $normalizedSecondOutTexts = @($snapshot.ResultRecords | ForEach-Object { $_.'正規化退場2' } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
            $normalizationStatuses = @($snapshot.ResultRecords | ForEach-Object { $_.'時刻正規化状態' } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
            $normalizedSecondOutTypes = @($snapshot.ResultTypedRecords | ForEach-Object { if ($null -ne $_.'正規化退場2_分') { $_.'正規化退場2_分'.GetType().Name } } | Select-Object -Unique)
            Assert-True -Condition ($snapshot.ResultRows -eq 6) -Message "建設現場転記PoCの行数が想定と異なります: $($snapshot.ResultRows)"
            Assert-True -Condition ($snapshot.ControlVersion -eq 'VER2') -Message "Control の版表示が想定と異なります: $($snapshot.ControlVersion)"
            Assert-True -Condition ($snapshot.ControlProfile -eq '建設現場転記PoCプロファイル') -Message "Control のプロファイル表示が想定と異なります: $($snapshot.ControlProfile)"
            Assert-True -Condition ($sourceNames -contains '2026年02月_作業員勤怠一覧_PoC.pdf') -Message 'PoC PDF 名が保持されていません。'
            Assert-True -Condition ($names -contains '佐藤 花子') -Message 'PoC 帳票の氏名が保持されていません。'
            Assert-True -Condition ((@($sites | Where-Object { $_ -like '*東京駅前再開発*' }).Count) -ge 1) -Message 'PoC 帳票の現場名が保持されていません。'
            Assert-True -Condition ($snapshot.ResultHeaders -contains '正規化入場1') -Message 'Result に正規化入場1 列がありません。'
            Assert-True -Condition ($snapshot.ResultHeaders -contains '正規化退場1_分') -Message 'Result に正規化退場1_分 列がありません。'
            Assert-True -Condition ($snapshot.ResultHeaders -contains '時刻正規化状態') -Message 'Result に時刻正規化状態 列がありません。'
            Assert-True -Condition ($snapshot.ReviewHeaders -contains '正規化入場1_raw') -Message 'Review に正規化入場1_raw 列がありません。'
            Assert-True -Condition ($snapshot.ReviewHeaders -contains 'ReasonCategory') -Message 'Review に ReasonCategory 列がありません。'
            Assert-True -Condition ($snapshot.ReviewHeaders -contains '正規化退場2') -Message 'Review に正規化退場2 列がありません。'
            Assert-True -Condition ($normalizedInMinutes -contains '555') -Message ('9時15分 の分換算結果が見つかりません: ' + ($normalizedInMinutes -join ','))
            Assert-True -Condition ($normalizedOutMinutes -contains '1080') -Message ('18:00 の分換算結果が見つかりません: ' + ($normalizedOutMinutes -join ','))
            Assert-True -Condition ($normalizedSecondOutMinutes -contains '1440') -Message ('24:00 の分換算結果が見つかりません: ' + ($normalizedSecondOutMinutes -join ','))
            Assert-True -Condition ($normalizedSecondOutTexts -contains '24:00') -Message ('24:00 の表示結果が見つかりません: ' + ($normalizedSecondOutTexts -join ','))
            Assert-True -Condition ($normalizedSecondOutTypes -contains 'Double') -Message ('正規化退場2_分 が数値型ではありません: ' + ($normalizedSecondOutTypes -join ','))
            Assert-True -Condition ($normalizationStatuses -contains 'OK') -Message ('時刻正規化状態に OK がありません: ' + ($normalizationStatuses -join ','))
            Assert-True -Condition ($snapshot.ReviewRows -ge 6) -Message "Review シートに行監査結果が十分に出ていません: $($snapshot.ReviewRows)"
            return "PoCPDF=$($sourceNames -join ','), 氏名=$($names -join ','), 現場=$($sites -join ','), Review=$($snapshot.ReviewRows), 正規化退場2分=$($normalizedSecondOutMinutes -join ',')"
        }
        'V2 2ページ同一列の変換' {
            $twoPageDir = Join-Path $fixturesRoot 'construction_2page_v2'
            Reset-Directory -Path $twoPageDir
            Copy-Item -LiteralPath (Join-Path $sampleConstructionPocPdfDir '2026年02月_作業員勤怠一覧_PoC_2ページ同一列.pdf') -Destination (Join-Path $twoPageDir '2026年02月_作業員勤怠一覧_PoC_2ページ同一列.pdf') -Force
            $outputPath = Join-Path $resultsRoot 'construction_transfer_poc_2page.xlsx'
            & powershell -NoProfile -ExecutionPolicy Bypass -File $runScriptV2 -InputFolder $twoPageDir -ProfilePath $script:constructionPocProfilePath -OutputFile $outputPath -NoConfirm
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message '2ページ同一列テストの出力ブックが作成されていません。'
            $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
            $twoPageNames = @($snapshot.ResultRecords | ForEach-Object { $_.'項目3' } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            $twoPageDistinctNames = @($twoPageNames | Select-Object -Unique)
            Assert-True -Condition ($snapshot.ResultRows -eq 11) -Message "2ページ同一列の Result 行数が想定と異なります: $($snapshot.ResultRows)"
            Assert-True -Condition ($snapshot.ControlVersion -eq 'VER2') -Message "2ページ同一列の版表示が想定と異なります: $($snapshot.ControlVersion)"
            Assert-True -Condition ($twoPageDistinctNames.Count -eq 5) -Message ('2ページ同一列で氏名重複が解消されていません: ' + ($twoPageDistinctNames -join ','))
            Assert-True -Condition ($snapshot.ReviewRows -ge 11) -Message '2ページ同一列で Review が十分に出ていません。'
            return "2ページ Result=$($snapshot.ResultRows), 氏名=$($twoPageDistinctNames -join ','), Review=$($snapshot.ReviewRows)"
        }
        'V2 6ページ同一列の変換' {
            $sixPageDir = Join-Path $fixturesRoot 'construction_6page_v2'
            Reset-Directory -Path $sixPageDir
            Copy-Item -LiteralPath (Join-Path $sampleConstructionPocPdfDir '2026年04月-06月_作業員勤怠一覧_PoC_6ページ同一列.pdf') -Destination (Join-Path $sixPageDir '2026年04月-06月_作業員勤怠一覧_PoC_6ページ同一列.pdf') -Force
            $outputPath = Join-Path $resultsRoot 'construction_transfer_poc_6page.xlsx'
            & powershell -NoProfile -ExecutionPolicy Bypass -File $runScriptV2 -InputFolder $sixPageDir -ProfilePath $script:constructionPocProfilePath -OutputFile $outputPath -NoConfirm
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message '6ページ同一列テストの出力ブックが作成されていません。'
            $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
            $fileNames = @($snapshot.ResultRecords | ForEach-Object { $_.'元ファイル名' } | Select-Object -Unique)
            $sixPageNames = @($snapshot.ResultRecords | ForEach-Object { $_.'項目3' } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
            Assert-True -Condition ($snapshot.ResultRows -eq 31) -Message "6ページ同一列の Result 行数が想定と異なります: $($snapshot.ResultRows)"
            Assert-True -Condition ($snapshot.ControlVersion -eq 'VER2') -Message "6ページ同一列の版表示が想定と異なります: $($snapshot.ControlVersion)"
            Assert-True -Condition ($fileNames.Count -eq 1) -Message '6ページ同一列で元ファイル名が分断されています。'
            Assert-True -Condition ($sixPageNames.Count -eq 5) -Message ('6ページ同一列で氏名の抽出が不安定です: ' + ($sixPageNames -join ','))
            Assert-True -Condition ($snapshot.ReviewRows -ge 31) -Message '6ページ同一列で Review が十分に出ていません。'
            return "6ページ Result=$($snapshot.ResultRows), 氏名=$($sixPageNames -join ','), Review=$($snapshot.ReviewRows)"
        }
        'V2 ヘッダー不一致負例の分離' {
            $headerMismatchDir = Join-Path $fixturesRoot 'construction_header_mismatch_v2'
            Reset-Directory -Path $headerMismatchDir
            Copy-Item -LiteralPath (Join-Path $sampleConstructionPocPdfDir '2026年02月_作業員勤怠一覧_PoC_ヘッダー不一致負例.pdf') -Destination (Join-Path $headerMismatchDir '2026年02月_作業員勤怠一覧_PoC_ヘッダー不一致負例.pdf') -Force
            $outputPath = Join-Path $resultsRoot 'construction_transfer_poc_header_mismatch.xlsx'
            & powershell -NoProfile -ExecutionPolicy Bypass -File $runScriptV2 -InputFolder $headerMismatchDir -ProfilePath $script:constructionPocProfilePath -OutputFile $outputPath -NoConfirm
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message 'ヘッダー不一致負例テストの出力ブックが作成されていません。'
            $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
            $reviewReasons = @($snapshot.ReviewRecords | ForEach-Object { $_.'Reason' } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            $reviewCategories = @($snapshot.ReviewRecords | ForEach-Object { $_.'ReasonCategory' } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            Assert-True -Condition ($snapshot.ResultRows -eq 1) -Message "ヘッダー不一致負例で Result 行が出ています: $($snapshot.ResultRows)"
            Assert-True -Condition ($snapshot.ErrorsRows -ge 2) -Message "ヘッダー不一致負例で Errors が不足しています: $($snapshot.ErrorsRows)"
            Assert-True -Condition ($snapshot.ReviewRows -ge 2) -Message "ヘッダー不一致負例で Review が不足しています: $($snapshot.ReviewRows)"
            Assert-True -Condition ((@($reviewReasons | Where-Object { $_ -like '*安全に結合できませんでした*' }).Count) -ge 1) -Message ('ヘッダー不一致負例で Review 理由が見つかりません: ' + ($reviewReasons -join ' | '))
            Assert-True -Condition ((@($reviewCategories | Where-Object { $_ -like '*HEADER_MISMATCH*' }).Count) -ge 1) -Message ('ヘッダー不一致負例で ReasonCategory が見つかりません: ' + ($reviewCategories -join ' | '))
            return "header-mismatch Result=$($snapshot.ResultRows), Errors=$($snapshot.ErrorsRows), Review=$($snapshot.ReviewRows), Categories=$($reviewCategories -join ',')"
        }
        'V2 時刻確認負例の分離' {
            $reviewNegativeDir = Join-Path $fixturesRoot 'construction_review_negative_v2'
            Reset-Directory -Path $reviewNegativeDir
            Copy-Item -LiteralPath (Join-Path $sampleConstructionPocPdfDir '2026年02月_作業員勤怠一覧_PoC_時刻確認負例.pdf') -Destination (Join-Path $reviewNegativeDir '2026年02月_作業員勤怠一覧_PoC_時刻確認負例.pdf') -Force
            $outputPath = Join-Path $resultsRoot 'construction_transfer_poc_review_negative.xlsx'
            & powershell -NoProfile -ExecutionPolicy Bypass -File $runScriptV2 -InputFolder $reviewNegativeDir -ProfilePath $script:constructionPocProfilePath -OutputFile $outputPath -NoConfirm
            Assert-True -Condition (Test-Path -LiteralPath $outputPath) -Message '時刻確認負例テストの出力ブックが作成されていません。'
            $snapshot = Get-WorkbookSnapshot -WorkbookPath $outputPath
            $reviewStatuses = @($snapshot.ReviewRecords | ForEach-Object { $_.'時刻正規化状態' } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
            $reviewNotes = @($snapshot.ReviewRecords | ForEach-Object { $_.'時刻確認メモ' } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            $reviewCategories = @($snapshot.ReviewRecords | ForEach-Object { $_.'ReasonCategory' } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
            Assert-True -Condition ($snapshot.ResultRows -eq 6) -Message "時刻確認負例で Result 行数が崩れています: $($snapshot.ResultRows)"
            Assert-True -Condition ($reviewStatuses -contains '要確認') -Message ('時刻確認負例で Review に 要確認 がありません: ' + ($reviewStatuses -join ','))
            Assert-True -Condition ((@($reviewNotes | Where-Object { $_ -like '*24:30*' -or $_ -like '*範囲外*' }).Count) -ge 1) -Message ('時刻確認負例で invalid 理由が見つかりません: ' + ($reviewNotes -join ' | '))
            Assert-True -Condition ((@($reviewNotes | Where-Object { $_ -like '*片側の時刻だけ*' }).Count) -ge 1) -Message ('時刻確認負例で片側空の理由が見つかりません: ' + ($reviewNotes -join ' | '))
            Assert-True -Condition ($reviewCategories -contains 'TIME_INVALID') -Message ('時刻確認負例で TIME_INVALID がありません: ' + ($reviewCategories -join ','))
            Assert-True -Condition ($reviewCategories -contains 'TIME_MISSING') -Message ('時刻確認負例で TIME_MISSING がありません: ' + ($reviewCategories -join ','))
            return "review-negative Result=$($snapshot.ResultRows), ReviewStatuses=$($reviewStatuses -join ','), Categories=$($reviewCategories -join ',')"
        }
        '一時領域の後片付け' {
            $runtimeRuns = @(Get-ChildItem -LiteralPath (Join-Path $projectRoot 'output\runtime\runs') -Directory -ErrorAction SilentlyContinue)
            $runtimeFileNames = @($runtimeRuns | Select-Object -ExpandProperty Name)
            Assert-True -Condition ($runtimeRuns.Count -eq 0) -Message ('一時ワークスペースが残っています: ' + ($runtimeFileNames -join ', '))
            return '一時領域は空'
        }
        'Excel プロセス残留なし' {
            $excelIds = @(Wait-For-ExcelBaseline -BaselineIds $suiteBaselineExcel -TimeoutSeconds 10)
            Assert-True -Condition ($excelIds.Count -eq 0) -Message ('Excel プロセスが残っています: ' + ($excelIds -join ', '))
            return 'Excel プロセス残留なし'
        }
        default {
            throw "未知のテストケースです: $Name"
        }
    }
}

function Invoke-IsolatedTestCase {
    param([Parameter(Mandatory = $true)]$Definition)

    function New-IsolatedFailureResult {
        param(
            [string]$ErrorMessage,
            [int]$DurationMs = 0
        )

        return [pscustomobject]@{
            Name         = $Definition.Name
            Scenario     = $Definition.Scenario
            Status       = 'FAIL'
            DurationMs   = $DurationMs
            Details      = $null
            ErrorMessage = $ErrorMessage
            RetryCount   = 0
            RetriedBy    = ''
        }
    }

    function Test-IsTransientExcelFailure {
        param([string]$ErrorMessage)

        if ([string]::IsNullOrWhiteSpace($ErrorMessage)) {
            return $false
        }

        foreach ($pattern in @(
            '0x800A03EC',
            'RPC_E_CALL_REJECTED',
            '呼び出し先が呼び出しを拒否しました',
            'message filter indicated that the application is busy',
            'Excel.*COM',
            'RefreshAll.*失敗',
            'Workbooks\.Open.*失敗'
        )) {
            if ($ErrorMessage -match $pattern) {
                return $true
            }
        }

        return $false
    }

    function Invoke-IsolatedOnce {
        param([int]$AttemptNumber)

        $resultPath = Join-Path $resultsRoot ("case_" + (($Definition.Name -replace '[^A-Za-z0-9]+', '_').Trim('_')) + "_$AttemptNumber.json")
        if (Test-Path -LiteralPath $resultPath) {
            Remove-Item -LiteralPath $resultPath -Force
        }

        $job = $null
        try {
            $job = Start-Job -ScriptBlock {
                param($scriptPath, $caseName, $singleResultPath)
                & $scriptPath -CaseName $caseName -SingleResultPath $singleResultPath
            } -ArgumentList $script:selfPath, $Definition.Name, $resultPath

            if (-not ($job | Wait-Job -Timeout $Definition.TimeoutSeconds -ErrorAction SilentlyContinue)) {
                Stop-Job -Job $job -ErrorAction SilentlyContinue
                return New-IsolatedFailureResult -ErrorMessage "$($Definition.TimeoutSeconds) 秒でタイムアウトしました。" -DurationMs ($Definition.TimeoutSeconds * 1000)
            }

            if (-not (Test-Path -LiteralPath $resultPath)) {
                return New-IsolatedFailureResult -ErrorMessage 'テスト結果ファイルが作成されませんでした。'
            }

            return Get-Content -LiteralPath $resultPath -Raw -Encoding UTF8 | ConvertFrom-Json
        } finally {
            if ($job) {
                Remove-Job -Job $job -ErrorAction SilentlyContinue
            }
            if (Test-Path -LiteralPath $resultPath) {
                Remove-Item -LiteralPath $resultPath -Force -ErrorAction SilentlyContinue
            }
        }
    }

    $firstResult = Invoke-IsolatedOnce -AttemptNumber 1
    if ($firstResult.Status -eq 'PASS' -or -not (Test-IsTransientExcelFailure -ErrorMessage $firstResult.ErrorMessage)) {
        return $firstResult
    }

    $retriedResult = Invoke-IsolatedOnce -AttemptNumber 2
    $retriedResult | Add-Member -NotePropertyName RetryCount -NotePropertyValue 1 -Force
    $retriedResult | Add-Member -NotePropertyName RetriedBy -NotePropertyValue 'EXCEL_TRANSIENT_COM' -Force
    return $retriedResult
}

if (-not [string]::IsNullOrWhiteSpace($CaseName)) {
    Initialize-TestFixtures
    $definition = Get-TestCases | Where-Object Name -eq $CaseName | Select-Object -First 1
    if ($null -eq $definition) {
        throw "未知のテストケースです: $CaseName"
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
    Write-Error ("統合テストに失敗しました: " + ($failed.Name -join ', '))
}

Write-Host "統合テスト成功: $(@($testResults | Where-Object Status -eq 'PASS').Count) / $($testResults.Count)"
Write-Host "Markdown レポート: $markdownReportPath"
Write-Host "JSON レポート: $jsonReportPath"
