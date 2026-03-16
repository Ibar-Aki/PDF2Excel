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

$testResults += Invoke-UnitTest -Name 'Get-ProfileConfiguration は生データ転記PoCプロファイルを読み込む' -Body {
    $profile = Get-ProfileConfiguration -RequestedProfilePath (Join-Path $projectRoot 'config\profiles\v2\construction_transfer_poc.json')
    Assert-True -Condition ($profile.ExpectedColumns -eq 30) -Message "expectedColumns が想定と異なります: $($profile.ExpectedColumns)"
    Assert-True -Condition ($profile.TargetRowCount -eq 5) -Message "targetRowCount が想定と異なります: $($profile.TargetRowCount)"
    Assert-True -Condition ($profile.DisplayName -eq '生データ転記PoCプロファイル') -Message "displayName が想定と異なります: $($profile.DisplayName)"
    Assert-True -Condition ($profile.MultiPageMergeMode -eq 'sameHeader') -Message "multiPageMergeMode が想定と異なります: $($profile.MultiPageMergeMode)"
    Assert-True -Condition ($profile.NormalizedTimeColumns.Count -eq 4) -Message "normalizedTimeColumns 数が想定と異なります: $($profile.NormalizedTimeColumns.Count)"
    return "$($profile.DisplayName) / $($profile.ExpectedColumns)"
}

$testResults += Invoke-UnitTest -Name 'TemplateBuilder VBA は V2 テンプレート生成マクロを持つ' -Body {
    $builderPath = Join-Path $projectRoot 'template\vba\PDF2ExcelTemplateBuilder.bas'
    $builderText = Get-Content -LiteralPath $builderPath -Raw -Encoding UTF8
    Assert-True -Condition ($builderText.Contains('Private Const REVIEW_SHEET As String = "Review"')) -Message 'Review シート定数が見つかりません。'
    Assert-True -Condition ($builderText.Contains('Public Sub BuildPDF2ExcelV2TemplateInActiveWorkbook()')) -Message 'V2 テンプレート生成マクロが見つかりません。'
    Assert-True -Condition ($builderText.Contains('GetRequiredSheetCount = 5')) -Message 'V2 用の 5 シート構成が見つかりません。'
    Assert-True -Condition (-not $builderText.Contains('Public Sub RefreshAndBuildWorkbook()')) -Message 'TemplateBuilder に重複する Refresh マクロが残っています。'
    Assert-True -Condition (-not $builderText.Contains('Public Sub ExportResultAsXlsx()')) -Message 'TemplateBuilder に重複する Export マクロが残っています。'
    return 'Review / BuildPDF2ExcelV2 / 5 シート / 重複除去を確認'
}

$testResults += Invoke-UnitTest -Name 'テンプレート再作成手順は配置先を明記する' -Body {
    $readmePath = Join-Path $projectRoot 'README.md'
    $manualPath = Join-Path $projectRoot 'docs\user-manual.md'
    $handoffReadmePath = Join-Path $projectRoot 'handoff\HANDOFF_README_V2_SOURCE.md'
    $readmeText = Get-Content -LiteralPath $readmePath -Raw -Encoding UTF8
    $manualText = Get-Content -LiteralPath $manualPath -Raw -Encoding UTF8
    $handoffReadmeText = Get-Content -LiteralPath $handoffReadmePath -Raw -Encoding UTF8
    Assert-True -Condition ($readmeText.Contains('template/PDF2Excel_V1_Converter.xlsm')) -Message 'README に V1 テンプレート配置先がありません。'
    Assert-True -Condition ($readmeText.Contains('template/PDF2Excel_V2_Converter.xlsm')) -Message 'README に V2 テンプレート配置先がありません。'
    Assert-True -Condition ($manualText.Contains('template/PDF2Excel_V1_Converter.xlsm')) -Message 'ユーザーマニュアルに V1 テンプレート配置先がありません。'
    Assert-True -Condition ($manualText.Contains('template/PDF2Excel_V2_Converter.xlsm')) -Message 'ユーザーマニュアルに V2 テンプレート配置先がありません。'
    Assert-True -Condition ($handoffReadmeText.Contains('template/PDF2Excel_V2_Converter.xlsm')) -Message 'handoff README に V2 テンプレート配置先がありません。'
    return 'README / user-manual / handoff README の配置先明記を確認'
}

$testResults += Invoke-UnitTest -Name 'run_pdf2excel.ps1 は Secure モードでローカル runtime を使う' -Body {
    $scriptText = Get-Content -LiteralPath $runScript -Raw -Encoding UTF8
    Assert-True -Condition ($scriptText.Contains("[ValidateSet('Standard', 'Secure')]")) -Message 'SecurityMode の ValidateSet が見つかりません。'
    Assert-True -Condition ($scriptText.Contains('Join-Path $env:LOCALAPPDATA ''PDF2Excel\runtime''')) -Message 'LOCALAPPDATA 配下の runtime ルートが見つかりません。'
    Assert-True -Condition ($scriptText.Contains("VER2 Secure では -KeepInput は無効")) -Message 'KeepInput 無効化メッセージが見つかりません。'
    return 'SecurityMode / LOCALAPPDATA / KeepInput 無効化を確認'
}

$testResults += Invoke-UnitTest -Name 'V2 ラッパーは Secure モードを渡し RemoteSigned で起動する' -Body {
    $scriptPath = Join-Path $projectRoot 'scripts\run_pdf2excel_v2.ps1'
    $menuPath = Join-Path $projectRoot 'scripts\run_pdf2excel_menu_v2.ps1'
    $scriptText = Get-Content -LiteralPath $scriptPath -Raw -Encoding UTF8
    $menuText = Get-Content -LiteralPath $menuPath -Raw -Encoding UTF8
    Assert-True -Condition ($scriptText.Contains("'-SecurityMode', 'Secure'")) -Message 'V2 ラッパーが SecurityMode Secure を渡していません。'
    Assert-True -Condition (-not $scriptText.Contains('ExecutionPolicy Bypass')) -Message 'V2 ラッパーに ExecutionPolicy Bypass が残っています。'
    Assert-True -Condition ($scriptText.Contains('ExecutionPolicy RemoteSigned')) -Message 'V2 ラッパーに ExecutionPolicy RemoteSigned がありません。'
    Assert-True -Condition (-not $menuText.Contains('ExecutionPolicy Bypass')) -Message 'V2 メニューラッパーに ExecutionPolicy Bypass が残っています。'
    Assert-True -Condition ($menuText.Contains('ExecutionPolicy RemoteSigned')) -Message 'V2 メニューラッパーに ExecutionPolicy RemoteSigned がありません。'
    return 'SecurityMode Secure / RemoteSigned / Bypass 除去を確認'
}

$testResults += Invoke-UnitTest -Name 'V2 BAT は ASCII かつ RemoteSigned で起動する' -Body {
    $batPath = Join-Path $projectRoot 'run_pdf2excel_v2.bat'
    $batBytes = [System.IO.File]::ReadAllBytes($batPath)
    $hasNonAscii = $false
    foreach ($value in $batBytes) {
        if ($value -gt 127) {
            $hasNonAscii = $true
            break
        }
    }
    $batText = [System.Text.Encoding]::ASCII.GetString($batBytes)
    Assert-True -Condition (-not $hasNonAscii) -Message 'V2 BAT に非 ASCII 文字が含まれています。'
    Assert-True -Condition (-not $batText.Contains('ExecutionPolicy Bypass')) -Message 'V2 BAT に ExecutionPolicy Bypass が残っています。'
    Assert-True -Condition ($batText.Contains('ExecutionPolicy RemoteSigned')) -Message 'V2 BAT に ExecutionPolicy RemoteSigned がありません。'
    return 'ASCII / RemoteSigned / Bypass 除去を確認'
}

$testResults += Invoke-UnitTest -Name 'V2 BAT とメニューは ForceMenu 導線を持つ' -Body {
    $batPath = Join-Path $projectRoot 'run_pdf2excel_v2.bat'
    $menuPath = Join-Path $projectRoot 'scripts\run_pdf2excel_menu.ps1'
    $menuV2Path = Join-Path $projectRoot 'scripts\run_pdf2excel_menu_v2.ps1'
    $batText = Get-Content -LiteralPath $batPath -Raw -Encoding ASCII
    $menuText = Get-Content -LiteralPath $menuPath -Raw -Encoding UTF8
    $menuV2Text = Get-Content -LiteralPath $menuV2Path -Raw -Encoding UTF8
    Assert-True -Condition ($batText.Contains('-ForceMenu')) -Message 'V2 BAT に ForceMenu 分岐がありません。'
    Assert-True -Condition ($menuText.Contains('[switch]$ForceMenu')) -Message '共通メニューに ForceMenu スイッチがありません。'
    Assert-True -Condition ($menuText.Contains('-not $ForceMenu -and $remainingArgs.Count -gt 0')) -Message '共通メニューに ForceMenu 優先分岐がありません。'
    Assert-True -Condition ($menuV2Text.Contains('[switch]$ForceMenu')) -Message 'V2 メニューラッパーに ForceMenu スイッチがありません。'
    return 'ForceMenu 導線を確認'
}

$testResults += Invoke-UnitTest -Name '共通メニューは完了メッセージを版別に分ける' -Body {
    $menuPath = Join-Path $projectRoot 'scripts\run_pdf2excel_menu.ps1'
    $menuText = Get-Content -LiteralPath $menuPath -Raw -Encoding UTF8
    Assert-True -Condition ($menuText.Contains('$completionSheetMessage = if ($VersionMode -eq ''v2'')')) -Message '完了メッセージの版別分岐がありません。'
    Assert-True -Condition ($menuText.Contains('Result / Review / Errors / Summary を確認してください。')) -Message 'V2 用完了メッセージがありません。'
    Assert-True -Condition ($menuText.Contains('Result / Errors / Summary を確認してください。')) -Message 'V1 用完了メッセージがありません。'
    return 'V1/V2 完了メッセージ分岐を確認'
}

$testResults += Invoke-UnitTest -Name 'handoff ビルドは VBA モジュールを同梱する' -Body {
    $scriptPath = Join-Path $projectRoot 'scripts\build_handoff_package.ps1'
    $scriptText = Get-Content -LiteralPath $scriptPath -Raw -Encoding UTF8
    Assert-True -Condition ($scriptText.Contains("'template\vba'")) -Message 'handoff 作成先に template\vba がありません。'
    Assert-True -Condition ($scriptText.Contains("template\vba\PDF2ExcelTemplateBuilder.bas")) -Message 'TemplateBuilder.bas の同梱が見つかりません。'
    Assert-True -Condition ($scriptText.Contains("template\vba\PDF2ExcelMacros.bas")) -Message 'PDF2ExcelMacros.bas の同梱が見つかりません。'
    Assert-True -Condition ($scriptText.Contains('function Sync-VbaModuleEncodingMirror')) -Message 'Shift_JIS ミラー同期関数が見つかりません。'
    Assert-True -Condition ($scriptText.Contains('PDF2ExcelTemplateBuilder.sjis.bas')) -Message 'TemplateBuilder.sjis.bas の同期が見つかりません。'
    return 'template\vba / sjis ミラー同期 / TemplateBuilder.bas / PDF2ExcelMacros.bas を確認'
}

$testResults += Invoke-UnitTest -Name 'build_handoff_package は V2 限定再生成を受け付ける' -Body {
    $scriptPath = Join-Path $projectRoot 'scripts\build_handoff_package.ps1'
    $scriptText = Get-Content -LiteralPath $scriptPath -Raw -Encoding UTF8
    Assert-True -Condition ($scriptText.Contains("[ValidateSet('all', 'v1', 'v2')]")) -Message 'TargetVersion の ValidateSet が見つかりません。'
    Assert-True -Condition ($scriptText.Contains('$TargetVersion -ne ''all''')) -Message 'TargetVersion のフィルタ分岐が見つかりません。'
    return 'TargetVersion フィルタを確認'
}

$testResults += Invoke-UnitTest -Name 'run_pdf2excel.ps1 は待機メッセージを表示する' -Body {
    $scriptText = Get-Content -LiteralPath $runScript -Raw -Encoding UTF8
    Assert-True -Condition ($scriptText.Contains('PDF取り込みに時間がかかります。しばらくお待ちください....')) -Message '待機メッセージが見つかりません。'
    Assert-True -Condition ($scriptText.Contains('Show-ProcessingNotice')) -Message '待機メッセージ呼び出しが見つかりません。'
    return '待機メッセージを確認'
}

$testResults += Invoke-UnitTest -Name 'Normalize-TimeText は全角コロンを半角時刻へ正規化する' -Body {
    $actual = Normalize-TimeText -Value '08：00'
    Assert-True -Condition ($actual.NormalizedText -eq '08:00') -Message "正規化時刻が想定と異なります: $($actual.NormalizedText)"
    Assert-True -Condition ($actual.MinutesFromMidnight -eq 480) -Message "分換算が想定と異なります: $($actual.MinutesFromMidnight)"
    Assert-True -Condition ($actual.Status -eq 'OK') -Message "状態が想定と異なります: $($actual.Status)"
    return "$($actual.NormalizedText) / $($actual.MinutesFromMidnight)"
}

$testResults += Invoke-UnitTest -Name 'Normalize-TimeText は時分表記と空白込みを正規化する' -Body {
    $actual = Normalize-TimeText -Value '9時 15分'
    Assert-True -Condition ($actual.NormalizedText -eq '09:15') -Message "正規化時刻が想定と異なります: $($actual.NormalizedText)"
    Assert-True -Condition ($actual.MinutesFromMidnight -eq 555) -Message "分換算が想定と異なります: $($actual.MinutesFromMidnight)"
    return "$($actual.NormalizedText) / $($actual.MinutesFromMidnight)"
}

$testResults += Invoke-UnitTest -Name 'Normalize-TimeText は時のみ表記を 00 分補完する' -Body {
    $actual = Normalize-TimeText -Value '18時'
    Assert-True -Condition ($actual.NormalizedText -eq '18:00') -Message "正規化時刻が想定と異なります: $($actual.NormalizedText)"
    Assert-True -Condition ($actual.MinutesFromMidnight -eq 1080) -Message "分換算が想定と異なります: $($actual.MinutesFromMidnight)"
    return "$($actual.NormalizedText) / $($actual.MinutesFromMidnight)"
}

$testResults += Invoke-UnitTest -Name 'Normalize-TimeText は Excel 時刻比率の文字列も正規化する' -Body {
    $actual = Normalize-TimeText -Value '0.385416666666667'
    Assert-True -Condition ($actual.NormalizedText -eq '09:15') -Message "正規化時刻が想定と異なります: $($actual.NormalizedText)"
    Assert-True -Condition ($actual.MinutesFromMidnight -eq 555) -Message "分換算が想定と異なります: $($actual.MinutesFromMidnight)"
    return "$($actual.NormalizedText) / $($actual.MinutesFromMidnight)"
}

$testResults += Invoke-UnitTest -Name 'Normalize-TimeText は 24:00 を有効時刻として扱う' -Body {
    $actual = Normalize-TimeText -Value '24：00'
    Assert-True -Condition ($actual.NormalizedText -eq '24:00') -Message "正規化時刻が想定と異なります: $($actual.NormalizedText)"
    Assert-True -Condition ($actual.MinutesFromMidnight -eq 1440) -Message "分換算が想定と異なります: $($actual.MinutesFromMidnight)"
    Assert-True -Condition ($actual.Status -eq 'OK') -Message "状態が想定と異なります: $($actual.Status)"
    return "$($actual.NormalizedText) / $($actual.MinutesFromMidnight)"
}

$testResults += Invoke-UnitTest -Name 'Normalize-TimeText は 24:30 を範囲外扱いにする' -Body {
    $actual = Normalize-TimeText -Value '24:30'
    Assert-True -Condition ($actual.Status -eq 'INVALID') -Message "状態が想定と異なります: $($actual.Status)"
    Assert-True -Condition ($actual.Note -eq '時刻の範囲外です。') -Message "メモが想定と異なります: $($actual.Note)"
    return "$($actual.Status) / $($actual.Note)"
}

$testResults += Invoke-UnitTest -Name 'Normalize-TimeText は 整数文字列 1 を 01:00 として扱う' -Body {
    $actual = Normalize-TimeText -Value '1'
    Assert-True -Condition ($actual.NormalizedText -eq '01:00') -Message "正規化時刻が想定と異なります: $($actual.NormalizedText)"
    Assert-True -Condition ($actual.MinutesFromMidnight -eq 60) -Message "分換算が想定と異なります: $($actual.MinutesFromMidnight)"
    Assert-True -Condition ($actual.Status -eq 'OK') -Message "状態が想定と異なります: $($actual.Status)"
    return "$($actual.NormalizedText) / $($actual.MinutesFromMidnight)"
}

$testResults += Invoke-UnitTest -Name 'Normalize-TimeText は 24:00:00 を 24:00 として扱う' -Body {
    $actual = Normalize-TimeText -Value '24:00:00'
    Assert-True -Condition ($actual.NormalizedText -eq '24:00') -Message "正規化時刻が想定と異なります: $($actual.NormalizedText)"
    Assert-True -Condition ($actual.MinutesFromMidnight -eq 1440) -Message "分換算が想定と異なります: $($actual.MinutesFromMidnight)"
    Assert-True -Condition ($actual.Status -eq 'OK') -Message "状態が想定と異なります: $($actual.Status)"
    return "$($actual.NormalizedText) / $($actual.MinutesFromMidnight)"
}

$testResults += Invoke-UnitTest -Name 'Get-ResultOutputColumnNames は V2 で正規化列を追加する' -Body {
    $profile = Get-ProfileConfiguration -RequestedProfilePath (Join-Path $projectRoot 'config\profiles\v2\construction_transfer_poc.json')
    $actual = @(Get-ResultOutputColumnNames -Profile $profile -VersionMode v2)
    Assert-True -Condition ($actual -contains '正規化入場1') -Message '正規化入場1 列が含まれていません。'
    Assert-True -Condition ($actual -contains '正規化退場1_分') -Message '正規化退場1_分 列が含まれていません。'
    Assert-True -Condition ($actual -contains '時刻正規化状態') -Message '時刻正規化状態 列が含まれていません。'
    return ($actual[-5..-1] -join ',')
}

$testResults += Invoke-UnitTest -Name 'Get-NormalizedTimeColumnDefinitions は Review raw 列名を返す' -Body {
    $profile = Get-ProfileConfiguration -RequestedProfilePath (Join-Path $projectRoot 'config\profiles\v2\construction_transfer_poc.json')
    $definitions = @(Get-NormalizedTimeColumnDefinitions -Profile $profile -VersionMode v2)
    Assert-True -Condition ($definitions[0].ReviewRawColumnName -eq '正規化入場1_raw') -Message "ReviewRawColumnName が想定と異なります: $($definitions[0].ReviewRawColumnName)"
    return ($definitions | ForEach-Object { $_.ReviewRawColumnName } | Select-Object -First 2) -join ','
}

$testResults += Invoke-UnitTest -Name 'Get-ReviewReasonCategories は HEADER_MISMATCH を先頭に重複なく返す' -Body {
    $profile = Get-ProfileConfiguration -RequestedProfilePath (Join-Path $projectRoot 'config\profiles\v2\construction_transfer_poc.json')
    $definitions = @(Get-NormalizedTimeColumnDefinitions -Profile $profile -VersionMode v2)
    $rawValues = @{}
    foreach ($definition in $definitions) {
        $rawValues[$definition.DisplayName] = ''
    }
    $rawValues[$definitions[0].DisplayName] = '08:00 09:00'
    $rawValues[$definitions[1].DisplayName] = ''
    $rawValues[$definitions[2].DisplayName] = '24:30'
    $rawValues[$definitions[3].DisplayName] = '18:00'

    $categories = Get-ReviewReasonCategories -Definitions $definitions -RawValuesByDisplayName $rawValues -ExistingReason '同一 PDF 内にヘッダー不一致または連続しない候補表があり、安全に結合できませんでした。' -ExistingCategoryCsv 'TIME_MULTI'

    Assert-True -Condition ($categories -eq 'HEADER_MISMATCH,TIME_MULTI,TIME_MISSING,TIME_INVALID') -Message "ReasonCategory の並びまたは重複除去が想定と異なります: $categories"
    return $categories
}

$testResults += Invoke-UnitTest -Name 'Get-StagingQueryFormula は V2 で canonical sameHeader と曖昧分離を考慮する' -Body {
    $profile = Get-ProfileConfiguration -RequestedProfilePath (Join-Path $projectRoot 'config\profiles\v2\construction_transfer_poc.json')
    $formula = Get-StagingQueryFormula -InputPath 'C:\Temp\Input' -Profile $profile
    Assert-True -Condition ($formula.Contains('GetCanonicalHeaderSignature')) -Message 'canonical ヘッダー署名ロジックが見つかりません。'
    Assert-True -Condition ($formula.Contains('FindHorizontalMergeSequences')) -Message 'partial merge sequence ロジックが見つかりません。'
    Assert-True -Condition ($formula.Contains('TABLE_GROUP_AMBIGUOUS')) -Message 'sameHeader 分離エラーが見つかりません。'
    Assert-True -Condition ($formula.Contains('正規化入場1_raw')) -Message 'Review raw 列が見つかりません。'
    Assert-True -Condition ($formula.Contains('ReasonCategory')) -Message 'Review の ReasonCategory 列が見つかりません。'
    Assert-True -Condition ($formula.Contains('TIME_MULTI')) -Message 'Review の TIME_MULTI 分類が見つかりません。'
    Assert-True -Condition ($formula.Contains('TIME_MISSING')) -Message 'Review の TIME_MISSING 分類が見つかりません。'
    return 'canonical sameHeader / ambiguity split / review raw / reason categories'
}

$testResults += Invoke-UnitTest -Name 'Get-ReviewQueryFormulaV2 は ReasonCategory を展開対象に含める' -Body {
    $profile = Get-ProfileConfiguration -RequestedProfilePath (Join-Path $projectRoot 'config\profiles\v2\construction_transfer_poc.json')
    $formula = Get-ReviewQueryFormulaV2 -Profile $profile
    Assert-True -Condition ($formula.Contains('ReasonCategory')) -Message 'Review クエリの展開列に ReasonCategory がありません。'
    return 'ReasonCategory expansion ready'
}

$testResults += Invoke-UnitTest -Name 'run_pdf2excel.bat は ASCII のみで構成される' -Body {
    $batPath = Join-Path $projectRoot 'run_pdf2excel.bat'
    $bytes = [System.IO.File]::ReadAllBytes($batPath)
    $nonAscii = @($bytes | Where-Object { $_ -gt 127 })
    Assert-True -Condition ($nonAscii.Count -eq 0) -Message 'run_pdf2excel.bat に非 ASCII バイトが含まれています。'
    return 'ASCII のみを確認'
}

$testResults += Invoke-UnitTest -Name 'run_pdf2excel_v1.bat と run_pdf2excel_v2.bat は ASCII のみで構成される' -Body {
    foreach ($batName in @('run_pdf2excel_v1.bat', 'run_pdf2excel_v2.bat')) {
        $batPath = Join-Path $projectRoot $batName
        $bytes = [System.IO.File]::ReadAllBytes($batPath)
        $nonAscii = @($bytes | Where-Object { $_ -gt 127 })
        Assert-True -Condition ($nonAscii.Count -eq 0) -Message "$batName に非 ASCII バイトが含まれています。"
    }
    return '版別 BAT の ASCII を確認'
}

$testResults += Invoke-UnitTest -Name 'run_pdf2excel_menu.ps1 は UTF-8 BOM で保存される' -Body {
    $menuPath = Join-Path $projectRoot 'scripts\run_pdf2excel_menu.ps1'
    $bytes = [System.IO.File]::ReadAllBytes($menuPath)
    Assert-True -Condition ($bytes.Length -ge 3) -Message 'run_pdf2excel_menu.ps1 が空です。'
    Assert-True -Condition ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) -Message 'run_pdf2excel_menu.ps1 が UTF-8 BOM ではありません。'
    return 'UTF-8 BOM を確認'
}

$testResults += Invoke-UnitTest -Name '版別メニュー PowerShell は UTF-8 BOM で保存される' -Body {
    foreach ($menuName in @('run_pdf2excel_menu_v1.ps1', 'run_pdf2excel_menu_v2.ps1')) {
        $menuPath = Join-Path $projectRoot ("scripts\{0}" -f $menuName)
        $bytes = [System.IO.File]::ReadAllBytes($menuPath)
        Assert-True -Condition ($bytes.Length -ge 3) -Message "$menuName が空です。"
        Assert-True -Condition ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) -Message "$menuName が UTF-8 BOM ではありません。"
    }
    return '版別メニューの UTF-8 BOM を確認'
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
