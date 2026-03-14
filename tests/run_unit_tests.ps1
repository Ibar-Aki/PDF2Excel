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
        $lines += "| $index | $($result.Name) | $($result.Status) | $($result.DurationMs) ms | $note |"
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

$testResults += Invoke-UnitTest -Name 'Escape-MString escapes quotes' -Body {
    $actual = Escape-MString -Value 'A"B'
    Assert-True -Condition ($actual -eq 'A""B') -Message "Unexpected escaped text: $actual"
    return $actual
}

$testResults += Invoke-UnitTest -Name 'ConvertTo-MTextListLiteral handles empty list' -Body {
    $actual = ConvertTo-MTextListLiteral -Values @()
    Assert-True -Condition ($actual -eq '{}') -Message "Unexpected literal: $actual"
    return $actual
}

$testResults += Invoke-UnitTest -Name 'ConvertTo-MLogicalLiteral returns lowercase literal' -Body {
    $trueValue = ConvertTo-MLogicalLiteral -Value $true
    $falseValue = ConvertTo-MLogicalLiteral -Value $false
    Assert-True -Condition ($trueValue -eq 'true') -Message "True literal mismatch: $trueValue"
    Assert-True -Condition ($falseValue -eq 'false') -Message "False literal mismatch: $falseValue"
    return "$trueValue / $falseValue"
}

$testResults += Invoke-UnitTest -Name 'Get-ProfileOutputColumnNames uses prefix and source file' -Body {
    $profile = [pscustomobject]@{
        SourceFileColumnName = 'SourceFile'
        ExpectedColumns      = 3
        DataColumnPrefix     = 'Column'
    }
    $actual = @(Get-ProfileOutputColumnNames -Profile $profile)
    Assert-True -Condition ($actual.Count -eq 4) -Message "Unexpected column count: $($actual.Count)"
    Assert-True -Condition ($actual[0] -eq 'SourceFile') -Message "Unexpected first column: $($actual[0])"
    Assert-True -Condition ($actual[3] -eq 'Column3') -Message "Unexpected last column: $($actual[3])"
    return ($actual -join ',')
}

$testResults += Invoke-UnitTest -Name 'Get-ProfileConfiguration loads default profile' -Body {
    $profile = Get-ProfileConfiguration -RequestedProfileName 'default'
    Assert-True -Condition ($profile.ExpectedColumns -eq 30) -Message "ExpectedColumns mismatch: $($profile.ExpectedColumns)"
    Assert-True -Condition ($profile.SourceFileColumnName -eq 'SourceFile') -Message "SourceFile column mismatch: $($profile.SourceFileColumnName)"
    return "$($profile.DisplayName) / $($profile.ExpectedColumns)"
}

$testResults += Invoke-UnitTest -Name 'Get-StagingQueryFormula embeds folder path and staging source' -Body {
    $profile = Get-ProfileConfiguration -RequestedProfileName 'default'
    $formula = Get-StagingQueryFormula -InputPath 'C:\Temp\Input' -Profile $profile
    Assert-True -Condition ($formula.Contains('Folder.Files("C:\Temp\Input")')) -Message 'Folder.Files path was not embedded.'
    Assert-True -Condition ($formula.Contains('COLUMN_OVERFLOW')) -Message 'Expected error code block was not found.'
    return 'Folder.Files + COLUMN_OVERFLOW confirmed'
}

$testResults += Invoke-UnitTest -Name 'Get-ResultQueryFormula references staging query' -Body {
    $profile = Get-ProfileConfiguration -RequestedProfileName 'default'
    $formula = Get-ResultQueryFormula -InputPath 'C:\Temp\Input' -Profile $profile
    Assert-True -Condition ($formula.Contains('Source = PDF2Excel_Staging')) -Message 'Result query does not reference staging query.'
    Assert-True -Condition (-not $formula.Contains('Folder.Files(')) -Message 'Result query still contains duplicated Folder.Files logic.'
    return 'Staging reference confirmed'
}

$testResults += Invoke-UnitTest -Name 'Resolve-RunErrorInfo maps lock and cancel states' -Body {
    $locked = Resolve-RunErrorInfo -Message '別の PDF2Excel 実行が進行中です。'
    $cancelled = Resolve-RunErrorInfo -Message '実行前チェックでキャンセルしました。'
    Assert-True -Condition ($locked.ErrorCode -eq 'RUN_LOCKED') -Message "Unexpected lock code: $($locked.ErrorCode)"
    Assert-True -Condition ($cancelled.ErrorCode -eq 'RUN_CANCELLED') -Message "Unexpected cancel code: $($cancelled.ErrorCode)"
    return "$($locked.ErrorCode) / $($cancelled.ErrorCode)"
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
    Write-Error ('Unit tests failed: ' + ($failed.Name -join ', '))
}

Write-Host "Unit tests passed: $(@($testResults | Where-Object Status -eq 'PASS').Count) / $($testResults.Count)"
Write-Host "Markdown report: $markdownReportPath"
Write-Host "JSON report: $jsonReportPath"
