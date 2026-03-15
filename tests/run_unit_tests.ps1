Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$resultsRoot = Join-Path $projectRoot 'tests\results'
$reportsRoot = Join-Path $projectRoot 'reports'
$jsonReportPath = Join-Path $resultsRoot 'unit-test-results.json'
$markdownReportPath = Join-Path $reportsRoot 'unit-test-report.md'
$commonScript = Join-Path $projectRoot 'scripts\pdf2excel.common.ps1'
$runScript = Join-Path $projectRoot 'scripts\run_pdf2excel.ps1'
$timeStarted = Get-Date

. $commonScript
. $runScript -SkipMain

Ensure-Directory -Path $resultsRoot
Ensure-Directory -Path $reportsRoot

function Assert-True {
    param(
        [Parameter(Mandatory = $true)][bool]$Condition,
        [Parameter(Mandatory = $true)][string]$Message
    )

    if (-not $Condition) {
        throw $Message
    }
}

function Invoke-UnitTest {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][scriptblock]$Body
    )

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $detail = & $Body
        $stopwatch.Stop()
        return [pscustomobject]@{
            Name         = $Name
            Status       = 'PASS'
            DurationMs   = [int]$stopwatch.ElapsedMilliseconds
            Details      = [string]$detail
            ErrorMessage = $null
        }
    } catch {
        $stopwatch.Stop()
        return [pscustomobject]@{
            Name         = $Name
            Status       = 'FAIL'
            DurationMs   = [int]$stopwatch.ElapsedMilliseconds
            Details      = $null
            ErrorMessage = $_.Exception.Message
        }
    }
}

function New-UnitReportMarkdown {
    param(
        [Parameter(Mandatory = $true)]$TestResults,
        [Parameter(Mandatory = $true)][datetime]$StartedAt,
        [Parameter(Mandatory = $true)][datetime]$FinishedAt
    )

    $duration = [int]($FinishedAt - $StartedAt).TotalSeconds
    $passCount = @($TestResults | Where-Object Status -eq 'PASS').Count
    $failCount = @($TestResults | Where-Object Status -eq 'FAIL').Count

    $lines = @(
        '# PDF2Excel ユニットテストレポート',
        '',
        ('- 作成日: {0} JST' -f $StartedAt.ToString('yyyy-MM-dd HH:mm')),
        '- 作成者: Codex (GPT-5)',
        ('- 更新日: {0}' -f $FinishedAt.ToString('yyyy-MM-dd')),
        '',
        '## サマリー',
        '',
        ('- 実施日時: {0} JST - {1} JST' -f $StartedAt.ToString('yyyy-MM-dd HH:mm:ss'), $FinishedAt.ToString('yyyy-MM-dd HH:mm:ss')),
        ('- 対象環境: Windows / PowerShell {0}' -f $PSVersionTable.PSVersion),
        '- 対象機能: 共通関数、プロファイル解決、Power Query 文字列生成',
        ('- 結果概要: {0} 件成功 / {1} 件失敗' -f $passCount, $failCount),
        ('- 所要時間: {0} 秒' -f $duration),
        ("- エラー有無: {0}" -f $(if ($failCount -eq 0) { 'なし' } else { 'あり' })),
        '',
        '## テスト結果',
        '',
        '| No | テスト | 結果 | 所要時間 | 補足 |',
        '| --- | --- | --- | --- | --- |'
    )

    $index = 1
    foreach ($result in $TestResults) {
        $note = if ($result.ErrorMessage) { $result.ErrorMessage } else { $result.Details }
        $statusLabel = if ($result.Status -eq 'PASS') { '成功' } else { '失敗' }
        $lines += "| $index | $($result.Name) | $statusLabel | $($result.DurationMs) ms | $note |"
        $index += 1
    }

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

$testResults = @()

$testResults += Invoke-UnitTest -Name 'Escape-MString は二重引用符をエスケープする' -Body {
    $actual = Escape-MString -Value 'A"B'
    Assert-True -Condition ($actual -eq 'A""B') -Message "エスケープ結果が想定と異なります: $actual"
    return $actual
}

$testResults += Invoke-UnitTest -Name 'ConvertTo-MTextListLiteral は空配列を処理できる' -Body {
    $actual = ConvertTo-MTextListLiteral -Values @()
    Assert-True -Condition ($actual -eq '{}') -Message "リテラル結果が想定と異なります: $actual"
    return $actual
}

$testResults += Invoke-UnitTest -Name 'ConvertTo-MLogicalLiteral は小文字の真偽値を返す' -Body {
    $trueValue = ConvertTo-MLogicalLiteral -Value $true
    $falseValue = ConvertTo-MLogicalLiteral -Value $false
    Assert-True -Condition ($trueValue -eq 'true') -Message "true の結果が想定と異なります: $trueValue"
    Assert-True -Condition ($falseValue -eq 'false') -Message "false の結果が想定と異なります: $falseValue"
    return "$trueValue / $falseValue"
}

$testResults += Invoke-UnitTest -Name 'Get-ProfileOutputColumnNames は接頭辞と元ファイル列を並べる' -Body {
    $profile = [pscustomobject]@{
        SourceFileColumnName = 'SourceFile'
        ExpectedColumns      = 3
        DataColumnPrefix     = 'Column'
    }
    $actual = @(Get-ProfileOutputColumnNames -Profile $profile)
    Assert-True -Condition ($actual.Count -eq 4) -Message "列数が想定と異なります: $($actual.Count)"
    Assert-True -Condition ($actual[0] -eq 'SourceFile') -Message "先頭列名が想定と異なります: $($actual[0])"
    Assert-True -Condition ($actual[3] -eq 'Column3') -Message "末尾列名が想定と異なります: $($actual[3])"
    return ($actual -join ',')
}

$testResults += Invoke-UnitTest -Name 'Get-ProfileConfiguration は既定プロファイルを読み込む' -Body {
    $profile = Get-ProfileConfiguration -RequestedProfileName 'default'
    Assert-True -Condition ($profile.ExpectedColumns -eq 30) -Message "expectedColumns が想定と異なります: $($profile.ExpectedColumns)"
    Assert-True -Condition ($profile.SourceFileColumnName -eq 'SourceFile') -Message "sourceFileColumnName が想定と異なります: $($profile.SourceFileColumnName)"
    return "$($profile.DisplayName) / $($profile.ExpectedColumns)"
}

$testResults += Invoke-UnitTest -Name 'Get-StagingQueryFormula は入力フォルダと判定ロジックを埋め込む' -Body {
    $profile = Get-ProfileConfiguration -RequestedProfileName 'default'
    $formula = Get-StagingQueryFormula -InputPath 'C:\Temp\Input' -Profile $profile
    Assert-True -Condition ($formula.Contains('Folder.Files("C:\Temp\Input")')) -Message 'Folder.Files の入力パスが埋め込まれていません。'
    Assert-True -Condition ($formula.Contains('COLUMN_OVERFLOW')) -Message 'COLUMN_OVERFLOW の判定ブロックが見つかりません。'
    return 'Folder.Files と COLUMN_OVERFLOW を確認'
}

$testResults += Invoke-UnitTest -Name 'Get-ResultQueryFormula は staging クエリを参照する' -Body {
    $profile = Get-ProfileConfiguration -RequestedProfileName 'default'
    $formula = Get-ResultQueryFormula -InputPath 'C:\Temp\Input' -Profile $profile
    Assert-True -Condition ($formula.Contains('Source = PDF2Excel_Staging')) -Message 'Result クエリが staging クエリを参照していません。'
    Assert-True -Condition (-not $formula.Contains('Folder.Files(')) -Message 'Result クエリに Folder.Files の重複ロジックが残っています。'
    return 'Staging 参照を確認'
}

$testResults += Invoke-UnitTest -Name 'Resolve-RunErrorInfo はロックとキャンセルを分類する' -Body {
    $locked = Resolve-RunErrorInfo -Message '別の PDF2Excel 実行が進行中です。'
    $cancelled = Resolve-RunErrorInfo -Message '実行前チェックでキャンセルしました。'
    Assert-True -Condition ($locked.ErrorCode -eq 'RUN_LOCKED') -Message "ロック時のコードが想定と異なります: $($locked.ErrorCode)"
    Assert-True -Condition ($cancelled.ErrorCode -eq 'RUN_CANCELLED') -Message "キャンセル時のコードが想定と異なります: $($cancelled.ErrorCode)"
    return "$($locked.ErrorCode) / $($cancelled.ErrorCode)"
}

$testResults += Invoke-UnitTest -Name 'Get-ProfileConfiguration は日本語勤怠プロファイルを読み込む' -Body {
    $profile = Get-ProfileConfiguration -RequestedProfileName 'attendance_monthly_jp'
    Assert-True -Condition ($profile.ExpectedColumns -eq 35) -Message "expectedColumns が想定と異なります: $($profile.ExpectedColumns)"
    Assert-True -Condition ($profile.TargetRowCount -eq 6) -Message "targetRowCount が想定と異なります: $($profile.TargetRowCount)"
    Assert-True -Condition ($profile.SourceFileColumnName -eq '元ファイル名') -Message "sourceFileColumnName が想定と異なります: $($profile.SourceFileColumnName)"
    return "$($profile.DisplayName) / $($profile.ExpectedColumns)"
}

$testResults += Invoke-UnitTest -Name 'Get-ProfileConfiguration は日本語売上日報プロファイルを読み込む' -Body {
    $profile = Get-ProfileConfiguration -RequestedProfileName 'sales_daily_jp'
    Assert-True -Condition ($profile.ExpectedColumns -eq 12) -Message "expectedColumns が想定と異なります: $($profile.ExpectedColumns)"
    Assert-True -Condition ($profile.TargetRowCount -eq 5) -Message "targetRowCount が想定と異なります: $($profile.TargetRowCount)"
    Assert-True -Condition ($profile.DisplayName -eq '日本語売上日報プロファイル') -Message "displayName が想定と異なります: $($profile.DisplayName)"
    return "$($profile.DisplayName) / $($profile.ExpectedColumns)"
}

$testResults += Invoke-UnitTest -Name 'Get-ProfileConfiguration は日本語在庫一覧プロファイルを読み込む' -Body {
    $profile = Get-ProfileConfiguration -RequestedProfileName 'inventory_list_jp'
    Assert-True -Condition ($profile.ExpectedColumns -eq 10) -Message "expectedColumns が想定と異なります: $($profile.ExpectedColumns)"
    Assert-True -Condition ($profile.TargetRowCount -eq 6) -Message "targetRowCount が想定と異なります: $($profile.TargetRowCount)"
    Assert-True -Condition ($profile.DisplayName -eq '日本語在庫一覧プロファイル') -Message "displayName が想定と異なります: $($profile.DisplayName)"
    return "$($profile.DisplayName) / $($profile.ExpectedColumns)"
}

$testResults += Invoke-UnitTest -Name 'Get-ProfileConfiguration は日本語問い合わせ管理表プロファイルを読み込む' -Body {
    $profile = Get-ProfileConfiguration -RequestedProfileName 'inquiry_weekly_jp'
    Assert-True -Condition ($profile.ExpectedColumns -eq 9) -Message "expectedColumns が想定と異なります: $($profile.ExpectedColumns)"
    Assert-True -Condition ($profile.TargetRowCount -eq 5) -Message "targetRowCount が想定と異なります: $($profile.TargetRowCount)"
    Assert-True -Condition ($profile.DisplayName -eq '日本語問い合わせ管理表プロファイル') -Message "displayName が想定と異なります: $($profile.DisplayName)"
    return "$($profile.DisplayName) / $($profile.ExpectedColumns)"
}

$testResults += Invoke-UnitTest -Name 'Get-ProfileConfiguration は建設現場転記PoCプロファイルを読み込む' -Body {
    $profile = Get-ProfileConfiguration -RequestedProfileName 'construction_transfer_poc'
    Assert-True -Condition ($profile.ExpectedColumns -eq 30) -Message "expectedColumns が想定と異なります: $($profile.ExpectedColumns)"
    Assert-True -Condition ($profile.TargetRowCount -eq 5) -Message "targetRowCount が想定と異なります: $($profile.TargetRowCount)"
    Assert-True -Condition ($profile.DisplayName -eq '建設現場転記PoCプロファイル') -Message "displayName が想定と異なります: $($profile.DisplayName)"
    return "$($profile.DisplayName) / $($profile.ExpectedColumns)"
}

$testResults += Invoke-UnitTest -Name 'run_pdf2excel.bat は ASCII のみで構成される' -Body {
    $batPath = Join-Path $projectRoot 'run_pdf2excel.bat'
    $bytes = [System.IO.File]::ReadAllBytes($batPath)
    $nonAscii = @($bytes | Where-Object { $_ -gt 127 })
    Assert-True -Condition ($nonAscii.Count -eq 0) -Message 'run_pdf2excel.bat に非 ASCII バイトが含まれています。'
    return 'ASCII のみを確認'
}

$testResults += Invoke-UnitTest -Name 'run_pdf2excel_menu.ps1 は UTF-8 BOM で保存される' -Body {
    $menuPath = Join-Path $projectRoot 'scripts\run_pdf2excel_menu.ps1'
    $bytes = [System.IO.File]::ReadAllBytes($menuPath)
    Assert-True -Condition ($bytes.Length -ge 3) -Message 'run_pdf2excel_menu.ps1 が空です。'
    Assert-True -Condition ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) -Message 'run_pdf2excel_menu.ps1 が UTF-8 BOM ではありません。'
    return 'UTF-8 BOM を確認'
}

$timeFinished = Get-Date
$summary = [pscustomobject]@{
    StartedAt  = $timeStarted
    FinishedAt = $timeFinished
    Results    = $testResults
}

$summary | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $jsonReportPath -Encoding UTF8
$reportMarkdown = New-UnitReportMarkdown -TestResults $testResults -StartedAt $timeStarted -FinishedAt $timeFinished
$reportMarkdown | Set-Content -LiteralPath $markdownReportPath -Encoding UTF8

$failed = @($testResults | Where-Object Status -eq 'FAIL')
if ($failed.Count -gt 0) {
    Write-Error ('ユニットテストに失敗しました: ' + ($failed.Name -join ', '))
}

Write-Host "ユニットテスト成功: $(@($testResults | Where-Object Status -eq 'PASS').Count) / $($testResults.Count)"
Write-Host "Markdown レポート: $markdownReportPath"
Write-Host "JSON レポート: $jsonReportPath"
