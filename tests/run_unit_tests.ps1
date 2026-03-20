Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$legacyRoot = Join-Path $projectRoot 'legacy\v1'
$resultsRoot = Join-Path $projectRoot 'tests\results'
$reportsRoot = Join-Path $projectRoot 'reports'
$jsonReportPath = Join-Path $resultsRoot 'unit-test-results.json'
$markdownReportPath = Join-Path $reportsRoot 'unit-test-report.md'
$commonScript = Join-Path $projectRoot 'scripts\pdf2excel.common.ps1'
$runScript = Join-Path $projectRoot 'scripts\run_pdf2excel.ps1'
$scaffoldScript = Join-Path $projectRoot 'scripts\new_profile_scaffold.ps1'
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

$testResults += Invoke-UnitTest -Name 'Get-PathLocationInfo は UNC とローカルパスを判定できる' -Body {
    $unc = Get-PathLocationInfo -Path '\\server\share\PDF2Excel'
    $local = Get-PathLocationInfo -Path 'C:\Work\PDF2Excel'
    Assert-True -Condition $unc.IsShared -Message 'UNC パスが shared と判定されません。'
    Assert-True -Condition $unc.IsUnc -Message 'UNC パスが UNC と判定されません。'
    Assert-True -Condition (-not $local.IsUnc) -Message 'ローカルパスが UNC 扱いされています。'
    return 'UNC / local を確認'
}

$testResults += Invoke-UnitTest -Name 'Get-LocalAppDataPdf2ExcelPath は PDF2Excel 配下を返す' -Body {
    $original = $env:LOCALAPPDATA
    try {
        $env:LOCALAPPDATA = 'C:\Users\TestUser\AppData\Local'
        $actual = Get-LocalAppDataPdf2ExcelPath -ChildPath 'logs'
        Assert-True -Condition ($actual -eq 'C:\Users\TestUser\AppData\Local\PDF2Excel\logs') -Message "LOCALAPPDATA 配下のパスが想定と異なります: $actual"
        return $actual
    } finally {
        $env:LOCALAPPDATA = $original
    }
}

$testResults += Invoke-UnitTest -Name 'Protect-MessagePaths は登録済みパスをマスクできる' -Body {
    $pathMap = [ordered]@{
        'C:\Users\TestUser\AppData\Local\PDF2Excel\logs\run_001.log' = '...\logs\run_001.log'
        'C:\Users\TestUser\AppData\Local\PDF2Excel\runtime\runs\run_001\staging\a.pdf' = '...\staging\a.pdf'
    }
    $message = '出力先=C:\Users\TestUser\AppData\Local\PDF2Excel\logs\run_001.log / staging=C:\Users\TestUser\AppData\Local\PDF2Excel\runtime\runs\run_001\staging\a.pdf'
    $protected = Protect-MessagePaths -Message $message -PathMap $pathMap
    Assert-True -Condition (-not $protected.Contains('C:\Users\TestUser\AppData\Local\PDF2Excel')) -Message "フルパスが残っています: $protected"
    Assert-True -Condition ($protected.Contains('...\logs\run_001.log')) -Message "ログパスのマスクが見つかりません: $protected"
    Assert-True -Condition ($protected.Contains('...\staging\a.pdf')) -Message "staging パスのマスクが見つかりません: $protected"
    return $protected
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

$testResults += Invoke-UnitTest -Name 'Resolve-RunErrorInfo はプロファイルと LOCALAPPDATA を分類する' -Body {
    $profile = Resolve-RunErrorInfo -Message 'プロファイルが見つかりません。'
    $local = Resolve-RunErrorInfo -Message 'VER2 Secure の正式運用では LOCALAPPDATA が必要です。'
    Assert-True -Condition ($profile.ErrorCode -eq 'PROFILE_RESOLUTION_ERROR') -Message "プロファイル時のコードが想定と異なります: $($profile.ErrorCode)"
    Assert-True -Condition ($local.ErrorCode -eq 'LOCAL_RUNTIME_UNAVAILABLE') -Message "LOCALAPPDATA 時のコードが想定と異なります: $($local.ErrorCode)"
    return "$($profile.ErrorCode) / $($local.ErrorCode)"
}

$testResults += Invoke-UnitTest -Name 'Get-RunErrorGuidance は次アクションを返す' -Body {
    $guidance = Get-RunErrorGuidance -ErrorInfo ([pscustomobject]@{ ErrorCode = 'RUN_LOCKED'; ErrorCategory = '実行競合' })
    Assert-True -Condition ($guidance.Contains('別の実行が完了するまで待って')) -Message "案内文が想定と異なります: $guidance"
    return $guidance
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

$testResults += Invoke-UnitTest -Name 'Get-ProfileConfiguration は生データ転記サンプルプロファイルを読み込む' -Body {
    $profile = Get-ProfileConfiguration -RequestedProfilePath (Join-Path $projectRoot 'config\profiles\v2\construction_transfer_poc.json')
    Assert-True -Condition ($profile.ExpectedColumns -eq 30) -Message "expectedColumns が想定と異なります: $($profile.ExpectedColumns)"
    Assert-True -Condition ($profile.TargetRowCount -eq 5) -Message "targetRowCount が想定と異なります: $($profile.TargetRowCount)"
    Assert-True -Condition ($profile.DisplayName -eq '生データ転記サンプルプロファイル') -Message "displayName が想定と異なります: $($profile.DisplayName)"
    Assert-True -Condition ($profile.MultiPageMergeMode -eq 'sameHeader') -Message "multiPageMergeMode が想定と異なります: $($profile.MultiPageMergeMode)"
    Assert-True -Condition ($profile.NormalizedTimeColumns.Count -eq 4) -Message "normalizedTimeColumns 数が想定と異なります: $($profile.NormalizedTimeColumns.Count)"
    return "$($profile.DisplayName) / $($profile.ExpectedColumns)"
}

$testResults += Invoke-UnitTest -Name 'Get-ProfileConfiguration は危険な ProfileName を拒否する' -Body {
    $thrown = $false
    try {
        Get-ProfileConfiguration -RequestedProfileName '..\default' | Out-Null
    } catch {
        $thrown = $true
        Assert-True -Condition ($_.Exception.Message.Contains('使用できない文字')) -Message "例外メッセージが想定と異なります: $($_.Exception.Message)"
    }

    Assert-True -Condition $thrown -Message '危険な ProfileName が拒否されませんでした。'
    return '危険な ProfileName を拒否'
}

$testResults += Invoke-UnitTest -Name 'Get-ProfileConfiguration は外部 ProfilePath を既定拒否し明示許可で読み込む' -Body {
    $externalProfilePath = Join-Path $resultsRoot 'external-profile.json'
    $externalProfileJson = @'
{
  "name": "external_profile",
  "displayName": "External Profile",
  "description": "External test profile",
  "expectedColumns": 2,
  "headerRowsToSkip": 0,
  "targetRowCount": 1,
  "allowMoreColumns": false,
  "preferredTableKinds": [],
  "preferredTableNameContains": [],
  "preferredTableIdContains": [],
  "sourceFileColumnName": "SourceFile",
  "dataColumnPrefix": "Column"
}
'@

    try {
        [System.IO.File]::WriteAllText($externalProfilePath, $externalProfileJson, (New-Object System.Text.UTF8Encoding($false)))

        $thrown = $false
        try {
            Get-ProfileConfiguration -RequestedProfilePath $externalProfilePath | Out-Null
        } catch {
            $thrown = $true
            Assert-True -Condition ($_.Exception.Message.Contains('config/profiles 配下以外')) -Message "例外メッセージが想定と異なります: $($_.Exception.Message)"
        }

        Assert-True -Condition $thrown -Message '外部 ProfilePath が既定で拒否されませんでした。'

        $profile = Get-ProfileConfiguration -RequestedProfilePath $externalProfilePath -AllowExternalProfilePath
        Assert-True -Condition ($profile.Name -eq 'external_profile') -Message "外部 profile 名が想定と異なります: $($profile.Name)"
        Assert-True -Condition ($profile.ProfilePath -eq $externalProfilePath) -Message "外部 profile の解決パスが想定と異なります: $($profile.ProfilePath)"
        return $profile.Name
    } finally {
        if (Test-Path -LiteralPath $externalProfilePath) {
            Remove-Item -LiteralPath $externalProfilePath -Force -ErrorAction SilentlyContinue
        }
    }
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
    $manualPath = Join-Path $projectRoot 'docs\01-current\01-user-manual.md'
    $handoffReadmePath = Join-Path $projectRoot 'handoff\sources\HANDOFF_README_V2_SOURCE.md'
    $readmeText = Get-Content -LiteralPath $readmePath -Raw -Encoding UTF8
    $manualText = Get-Content -LiteralPath $manualPath -Raw -Encoding UTF8
    $handoffReadmeText = Get-Content -LiteralPath $handoffReadmePath -Raw -Encoding UTF8
    Assert-True -Condition ($readmeText.Contains('template/PDF2Excel_V2_Converter.xlsm')) -Message 'README に V2 テンプレート配置先がありません。'
    Assert-True -Condition ($manualText.Contains('template/PDF2Excel_V2_Converter.xlsm')) -Message 'ユーザーマニュアルに V2 テンプレート配置先がありません。'
    Assert-True -Condition ($handoffReadmeText.Contains('template/PDF2Excel_V2_Converter.xlsm')) -Message 'handoff README に V2 テンプレート配置先がありません。'
    return 'README / user-manual / handoff README の配置先明記を確認'
}

$testResults += Invoke-UnitTest -Name 'run_pdf2excel.ps1 は Secure モードでローカル runtime を使う' -Body {
    $scriptText = Get-Content -LiteralPath $runScript -Raw -Encoding UTF8
    Assert-True -Condition ($scriptText.Contains("[ValidateSet('Standard', 'Secure')]")) -Message 'SecurityMode の ValidateSet が見つかりません。'
    Assert-True -Condition ($scriptText.Contains("[ValidateSet('INFO', 'DEBUG')]")) -Message 'LogLevel の ValidateSet が見つかりません。'
    Assert-True -Condition ($scriptText.Contains("Get-LocalAppDataPdf2ExcelPath -ChildPath 'runtime'")) -Message 'LOCALAPPDATA 配下の runtime ルート解決が見つかりません。'
    Assert-True -Condition ($scriptText.Contains("Get-LocalAppDataPdf2ExcelPath -ChildPath 'logs'")) -Message 'LOCALAPPDATA 配下の logs ルート解決が見つかりません。'
    Assert-True -Condition ($scriptText.Contains('[switch]$CheckEnvironment')) -Message 'CheckEnvironment オプションが見つかりません。'
    Assert-True -Condition ($scriptText.Contains('[switch]$AllowExternalProfilePath')) -Message 'AllowExternalProfilePath オプションが見つかりません。'
    Assert-True -Condition ($scriptText.Contains('[string]$RunReportPath')) -Message 'RunReportPath オプションが見つかりません。'
    Assert-True -Condition ($scriptText.Contains('Join-Path $reportsDir ''run-history.csv''')) -Message 'run-history.csv の出力先が見つかりません。'
    Assert-True -Condition ($scriptText.Contains('Join-Path $reportsDir ''environment-check.md''')) -Message 'environment-check.md の出力先が見つかりません。'
    Assert-True -Condition ($scriptText.Contains('Protect-MessagePaths -Message $Message -PathMap $script:sensitivePathMap')) -Message 'Secure ログのパスマスクが見つかりません。'
    Assert-True -Condition ($scriptText.Contains('if ($Level -eq ''DEBUG'' -and $LogLevel -ne ''DEBUG'')')) -Message 'DEBUG ログ抑制の分岐が見つかりません。'
    Assert-True -Condition ($scriptText.Contains('Stop-Process -Id $ProcessId -Force -ErrorAction Stop')) -Message 'Excel 強制終了の実装が見つかりません。'
    Assert-True -Condition ($scriptText.Contains('function Get-PdfPreviewSummary')) -Message '簡易プレビュー関数が見つかりません。'
    Assert-True -Condition ($scriptText.Contains('function Invoke-EnvironmentCheck')) -Message '環境チェック関数が見つかりません。'
    Assert-True -Condition ($scriptText.Contains('テンプレート再生成導線')) -Message '環境チェックにテンプレート再生成導線がありません。'
    Assert-True -Condition ($scriptText.Contains('実行競合状態')) -Message '環境チェックに実行競合状態がありません。'
    Assert-True -Condition ($scriptText.Contains('Append-RunHistory -Status ''SUCCESS''')) -Message '実行履歴追記が見つかりません。'
    Assert-True -Condition (-not $scriptText.Contains('& powershell -NoProfile -ExecutionPolicy Bypass -File $buildTemplateScript')) -Message 'テンプレート生成で Bypass が残っています。'
    Assert-True -Condition ($scriptText.Contains("VER2 Secure では -KeepInput は無効")) -Message 'KeepInput 無効化メッセージが見つかりません。'
    Assert-True -Condition ($scriptText.Contains('Clear-ControlPathsForSecureOutput')) -Message 'Secure 出力向け Control パス空欄化が見つかりません。'
    Assert-True -Condition ($scriptText.Contains('$templateIntegrityPath = Join-Path $configDir ''template-integrity.json''')) -Message 'テンプレート整合性マニフェストの参照がありません。'
    Assert-True -Condition ($scriptText.Contains('Assert-TemplateIntegrity')) -Message 'テンプレート整合性チェックがありません。'
    return 'SecurityMode / LogLevel / LOCALAPPDATA runtime/logs / path mask / template integrity / KeepInput 無効化を確認'
}

$testResults += Invoke-UnitTest -Name 'run_pdf2excel_menu.ps1 はプロファイル選択と環境チェック導線を持つ' -Body {
    $menuPath = Join-Path $projectRoot 'scripts\run_pdf2excel_menu.ps1'
    $menuText = Get-Content -LiteralPath $menuPath -Raw -Encoding UTF8
    Assert-True -Condition ($menuText.Contains('[9] プロファイルを選ぶ')) -Message 'メニューにプロファイル選択がありません。'
    Assert-True -Condition ($menuText.Contains('[0] 環境チェック')) -Message 'メニューに環境チェックがありません。'
    Assert-True -Condition ($menuText.Contains('Get-RunReportTempPath')) -Message 'メニューの実行レポート連携がありません。'
    Assert-True -Condition ($menuText.Contains('現在のプロファイル')) -Message '現在のプロファイル表示がありません。'
    return 'profile select / environment check / run report / current profile を確認'
}

$testResults += Invoke-UnitTest -Name 'V2 ラッパーは Secure モードを渡し RemoteSigned で起動する' -Body {
    $scriptPath = Join-Path $projectRoot 'scripts\run_pdf2excel_v2.ps1'
    $menuPath = Join-Path $projectRoot 'scripts\run_pdf2excel_menu_v2.ps1'
    $scriptText = Get-Content -LiteralPath $scriptPath -Raw -Encoding UTF8
    $menuText = Get-Content -LiteralPath $menuPath -Raw -Encoding UTF8
    Assert-True -Condition ($scriptText.Contains("'-SecurityMode', 'Secure'")) -Message 'V2 ラッパーが SecurityMode Secure を渡していません。'
    Assert-True -Condition ($scriptText.Contains('ExecutionPolicy RemoteSigned')) -Message 'V2 ラッパーに ExecutionPolicy RemoteSigned がありません。'
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
    Assert-True -Condition ($batText.Contains('ExecutionPolicy RemoteSigned')) -Message 'V2 BAT に ExecutionPolicy RemoteSigned がありません。'
    return 'ASCII / RemoteSigned / Bypass 除去を確認'
}

$testResults += Invoke-UnitTest -Name '正式運用 BAT は V2 Secure を起動し RemoteSigned に統一されている' -Body {
    $batText = Get-Content -LiteralPath (Join-Path $projectRoot 'run_pdf2excel.bat') -Raw -Encoding ASCII
    $v1Bat = Get-Content -LiteralPath (Join-Path $legacyRoot 'run_pdf2excel_v1.bat') -Raw -Encoding ASCII
    $v1Script = Get-Content -LiteralPath (Join-Path $legacyRoot 'scripts\run_pdf2excel_v1.ps1') -Raw -Encoding UTF8
    $v1Menu = Get-Content -LiteralPath (Join-Path $legacyRoot 'scripts\run_pdf2excel_menu_v1.ps1') -Raw -Encoding UTF8
    Assert-True -Condition ($batText.Contains('ExecutionPolicy RemoteSigned')) -Message '正式運用 BAT に RemoteSigned がありません。'
    Assert-True -Condition ($batText.Contains('-VersionMode v2')) -Message '正式運用 BAT が V2 を起動していません。'
    Assert-True -Condition ($batText.Contains('-DefaultSecurityMode Secure')) -Message '正式運用 BAT が Secure を既定化していません。'
    Assert-True -Condition ($batText.Contains('-DefaultProfileName construction_transfer_poc')) -Message '正式運用 BAT の既定プロファイルが見つかりません。'
    foreach ($text in @($v1Bat, $v1Script, $v1Menu)) {
        Assert-True -Condition ($text.Contains('ExecutionPolicy RemoteSigned')) -Message 'V1 導線に RemoteSigned がありません。'
        Assert-True -Condition (-not $text.Contains('ExecutionPolicy Bypass')) -Message 'V1 導線に Bypass が残っています。'
    }
    Assert-True -Condition ($v1Menu.Contains('ProfileDirOverride')) -Message 'V1 legacy メニューが profile override を渡していません。'
    return 'run_pdf2excel.bat は V2 Secure、legacy 導線も RemoteSigned を確認'
}

$testResults += Invoke-UnitTest -Name 'テストとビルド導線に Bypass が残っていない' -Body {
    $integrationText = Get-Content -LiteralPath (Join-Path $projectRoot 'tests\run_integration_tests.ps1') -Raw -Encoding UTF8
    $scriptText = Get-Content -LiteralPath $runScript -Raw -Encoding UTF8
    Assert-True -Condition (-not $integrationText.Contains('-ExecutionPolicy Bypass')) -Message 'run_integration_tests.ps1 に Bypass が残っています。'
    Assert-True -Condition (-not $scriptText.Contains('-ExecutionPolicy Bypass')) -Message 'run_pdf2excel.ps1 に Bypass が残っています。'
    Assert-True -Condition ($integrationText.Contains('ExecutionPolicy RemoteSigned')) -Message 'run_integration_tests.ps1 に RemoteSigned が見つかりません。'
    return 'tests / build 呼び出しの Bypass 除去を確認'
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

$testResults += Invoke-UnitTest -Name '共通メニューは共有パス制約とプロファイル雛形導線を持つ' -Body {
    $menuPath = Join-Path $projectRoot 'scripts\run_pdf2excel_menu.ps1'
    $menuText = Get-Content -LiteralPath $menuPath -Raw -Encoding UTF8
    Assert-True -Condition ($menuText.Contains('[string]$VersionMode = ''v2''')) -Message '共通メニューの既定 VersionMode が v2 ではありません。'
    Assert-True -Condition ($menuText.Contains('Assert-ExecutionLocationAllowed')) -Message '共有パス制約の呼び出しが見つかりません。'
    Assert-True -Condition ($menuText.Contains('Assert-SecureLocalStorageAvailable')) -Message 'LOCALAPPDATA 必須チェックが見つかりません。'
    Assert-True -Condition ($menuText.Contains('VER2 Secure の正式運用では共有パス上から実行できません')) -Message '共有パス拒否メッセージが見つかりません。'
    Assert-True -Condition ($menuText.Contains('Invoke-ProfileScaffoldScript')) -Message 'プロファイル雛形生成の呼び出しが見つかりません。'
    Assert-True -Condition ($menuText.Contains('[6] プロファイル雛形を作成')) -Message 'メニュー項目 6 が見つかりません。'
    Assert-True -Condition ($menuText.Contains('[8] 終了')) -Message '終了メニュー番号の更新が見つかりません。'
    Assert-True -Condition (-not $menuText.Contains('ExecutionPolicy'', ''Bypass')) -Message '共通メニューに Bypass が残っています。'
    return '共有パス制約 / LOCALAPPDATA 必須 / プロファイル雛形 / メニュー番号を確認'
}

$testResults += Invoke-UnitTest -Name '共通メニューは完了メッセージを版別に分ける' -Body {
    $menuPath = Join-Path $projectRoot 'scripts\run_pdf2excel_menu.ps1'
    $menuText = Get-Content -LiteralPath $menuPath -Raw -Encoding UTF8
    Assert-True -Condition ($menuText.Contains('$completionSheetMessage = if ([string]::IsNullOrWhiteSpace($CompletionSheetMessageOverride))')) -Message '完了メッセージの override 分岐がありません。'
    Assert-True -Condition ($menuText.Contains('Result / Review / Errors / Summary を確認してください。')) -Message 'V2 用完了メッセージがありません。'
    Assert-True -Condition ($menuText.Contains('Result / Errors / Summary を確認してください。')) -Message 'V1 用完了メッセージがありません。'
    return 'V1/V2 完了メッセージと override 分岐を確認'
}

$testResults += Invoke-UnitTest -Name 'handoff ビルドは VBA モジュールを同梱する' -Body {
    $scriptPath = Join-Path $projectRoot 'scripts\build_handoff_package.ps1'
    $scriptText = Get-Content -LiteralPath $scriptPath -Raw -Encoding UTF8
    Assert-True -Condition ($scriptText.Contains("'template\vba'")) -Message 'handoff 作成先に template\vba がありません。'
    Assert-True -Condition ($scriptText.Contains("config\template-integrity.json")) -Message 'template-integrity.json の同梱が見つかりません。'
    Assert-True -Condition ($scriptText.Contains("template\vba\PDF2ExcelTemplateBuilder.bas")) -Message 'TemplateBuilder.bas の同梱が見つかりません。'
    Assert-True -Condition ($scriptText.Contains("template\vba\PDF2ExcelMacros.bas")) -Message 'PDF2ExcelMacros.bas の同梱が見つかりません。'
    Assert-True -Condition ($scriptText.Contains('function Sync-VbaModuleEncodingMirror')) -Message 'Shift_JIS ミラー同期関数が見つかりません。'
    Assert-True -Condition ($scriptText.Contains('PDF2ExcelTemplateBuilder.sjis.bas')) -Message 'TemplateBuilder.sjis.bas の同期が見つかりません。'
    Assert-True -Condition ($scriptText.Contains('new_profile_scaffold.ps1')) -Message 'new_profile_scaffold.ps1 の同梱が見つかりません。'
    return 'template\vba / sjis ミラー同期 / new_profile_scaffold.ps1 / TemplateBuilder.bas / PDF2ExcelMacros.bas を確認'
}

$testResults += Invoke-UnitTest -Name '生成済み V2 handoff は配布元ソースと同期している' -Body {
    $handoffRoot = Join-Path $projectRoot 'handoff\generated\PDF2Excel_V2_Minimal'
    $pairs = @(
        @{
            Source = Join-Path $projectRoot 'handoff\sources\HANDOFF_README_V2_SOURCE.md'
            Generated = Join-Path $handoffRoot 'HANDOFF_README.md'
        },
        @{
            Source = Join-Path $projectRoot 'run_pdf2excel_v2.bat'
            Generated = Join-Path $handoffRoot 'run_pdf2excel.bat'
        },
        @{
            Source = Join-Path $projectRoot 'scripts\run_pdf2excel.ps1'
            Generated = Join-Path $handoffRoot 'scripts\run_pdf2excel.ps1'
        },
        @{
            Source = Join-Path $projectRoot 'scripts\run_pdf2excel_menu.ps1'
            Generated = Join-Path $handoffRoot 'scripts\run_pdf2excel_menu.ps1'
        },
        @{
            Source = Join-Path $projectRoot 'scripts\new_profile_scaffold.ps1'
            Generated = Join-Path $handoffRoot 'scripts\new_profile_scaffold.ps1'
        },
        @{
            Source = Join-Path $projectRoot 'scripts\pdf2excel.common.ps1'
            Generated = Join-Path $handoffRoot 'scripts\pdf2excel.common.ps1'
        },
        @{
            Source = Join-Path $projectRoot 'config\profiles\v2\construction_transfer_poc.json'
            Generated = Join-Path $handoffRoot 'config\profiles\v2\construction_transfer_poc.json'
        }
    )

    foreach ($pair in $pairs) {
        Assert-True -Condition (Test-Path -LiteralPath $pair.Generated) -Message "handoff 生成物が見つかりません: $($pair.Generated)"
        $sourceText = Get-Content -LiteralPath $pair.Source -Raw -Encoding UTF8
        $generatedText = Get-Content -LiteralPath $pair.Generated -Raw -Encoding UTF8
        Assert-True -Condition ($sourceText -eq $generatedText) -Message "handoff 生成物が古いか差分があります: $($pair.Generated)"
    }

    return 'README / BAT / scripts / profile の handoff 同期を確認'
}

$testResults += Invoke-UnitTest -Name 'new_profile_scaffold は v1/v2 雛形を生成できる' -Body {
    $scaffoldRoot = Join-Path $resultsRoot 'profile-scaffold'
    Ensure-Directory -Path $scaffoldRoot
    $v1Path = Join-Path $scaffoldRoot 'sample_v1.json'
    $v2Path = Join-Path $scaffoldRoot 'sample_v2.json'
    foreach ($path in @($v1Path, $v2Path)) {
        if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force
        }
    }

    & powershell -NoProfile -ExecutionPolicy RemoteSigned -File $scaffoldScript -VersionMode v1 -ProfileName sample_v1 -DisplayName 'テストV1' -OutputPath $v1Path
    Assert-True -Condition ($LASTEXITCODE -eq 0) -Message "v1 雛形生成が失敗しました: $LASTEXITCODE"
    & powershell -NoProfile -ExecutionPolicy RemoteSigned -File $scaffoldScript -VersionMode v2 -ProfileName sample_v2 -DisplayName 'テストV2' -OutputPath $v2Path
    Assert-True -Condition ($LASTEXITCODE -eq 0) -Message "v2 雛形生成が失敗しました: $LASTEXITCODE"

    $v1 = Get-Content -LiteralPath $v1Path -Raw -Encoding UTF8 | ConvertFrom-Json
    $v2 = Get-Content -LiteralPath $v2Path -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-True -Condition ($v1.displayName -eq 'テストV1') -Message "v1 displayName が想定と異なります: $($v1.displayName)"
    Assert-True -Condition ($v1.sourceFileColumnName -eq 'SourceFile') -Message "v1 sourceFileColumnName が想定と異なります: $($v1.sourceFileColumnName)"
    Assert-True -Condition ($v2.displayName -eq 'テストV2') -Message "v2 displayName が想定と異なります: $($v2.displayName)"
    Assert-True -Condition ($v2.multiPageMergeMode -eq 'sameHeader') -Message "v2 multiPageMergeMode が想定と異なります: $($v2.multiPageMergeMode)"
    Assert-True -Condition ($v2.sourceFileColumnName -eq '元ファイル名') -Message "v2 sourceFileColumnName が想定と異なります: $($v2.sourceFileColumnName)"
    return 'v1/v2 雛形生成を確認'
}

$testResults += Invoke-UnitTest -Name 'new_profile_scaffold は v2 Wizard で主要設定を生成できる' -Body {
    $scaffoldRoot = Join-Path $resultsRoot 'profile-scaffold'
    Ensure-Directory -Path $scaffoldRoot
    $wizardPath = Join-Path $scaffoldRoot 'sample_v2_wizard.json'
    if (Test-Path -LiteralPath $wizardPath) {
        Remove-Item -LiteralPath $wizardPath -Force
    }

    $wizardInput = @(
        'sample_v2_wizard',
        'テストV2ウィザード',
        $wizardPath,
        'ウィザード説明',
        '28',
        '2',
        '6',
        'Y',
        'sameHeader',
        '元ファイル名',
        '項目',
        '勤怠,現場',
        '3',
        '5',
        '6',
        '7',
        '6:正規化入場1,7:正規化退場1',
        'Y'
    ) -join [Environment]::NewLine
    $wizardInput | & powershell -NoProfile -ExecutionPolicy RemoteSigned -File $scaffoldScript -VersionMode v2 -Wizard
    Assert-True -Condition ($LASTEXITCODE -eq 0) -Message "v2 Wizard 生成が失敗しました: $LASTEXITCODE"

    $wizardProfile = Get-Content -LiteralPath $wizardPath -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-True -Condition ($wizardProfile.displayName -eq 'テストV2ウィザード') -Message "Wizard displayName が想定と異なります: $($wizardProfile.displayName)"
    Assert-True -Condition ($wizardProfile.expectedColumns -eq 28) -Message "Wizard expectedColumns が想定と異なります: $($wizardProfile.expectedColumns)"
    Assert-True -Condition ($wizardProfile.allowMoreColumns -eq $true) -Message 'Wizard allowMoreColumns が true ではありません。'
    Assert-True -Condition (@($wizardProfile.preferredTableNameContains).Count -eq 2) -Message 'Wizard preferredTableNameContains が想定と異なります。'
    Assert-True -Condition (@($wizardProfile.normalizedTimeColumns).Count -eq 2) -Message 'Wizard normalizedTimeColumns が想定と異なります。'
    return 'v2 Wizard 雛形生成を確認'
}

$testResults += Invoke-UnitTest -Name 'build_handoff_package は V2 既定生成と legacy v1 参照を持つ' -Body {
    $scriptPath = Join-Path $projectRoot 'scripts\build_handoff_package.ps1'
    $scriptText = Get-Content -LiteralPath $scriptPath -Raw -Encoding UTF8
    Assert-True -Condition ($scriptText.Contains("[ValidateSet('all', 'v1', 'v2')]")) -Message 'TargetVersion の ValidateSet が見つかりません。'
    Assert-True -Condition ($scriptText.Contains('[string]$TargetVersion = ''v2''')) -Message 'TargetVersion の既定値が v2 ではありません。'
    Assert-True -Condition ($scriptText.Contains('$TargetVersion -ne ''all''')) -Message 'TargetVersion のフィルタ分岐が見つかりません。'
    Assert-True -Condition ($scriptText.Contains("RunBat = 'legacy\v1\run_pdf2excel_v1.bat'")) -Message 'legacy v1 BAT の参照が見つかりません。'
    Assert-True -Condition ($scriptText.Contains("ReadmeSource = 'legacy\v1\handoff\sources\HANDOFF_README_SOURCE.md'")) -Message 'legacy v1 handoff README の参照が見つかりません。'
    return 'TargetVersion 既定値 v2 と legacy v1 配布元を確認'
}

$testResults += Invoke-UnitTest -Name 'run_integration_tests は Suite 指定で smoke full legacy を切り替えられる' -Body {
    $integrationText = Get-Content -LiteralPath (Join-Path $projectRoot 'tests\run_integration_tests.ps1') -Raw -Encoding UTF8
    Assert-True -Condition ($integrationText.Contains("[ValidateSet('smoke', 'full', 'legacy', 'all')]")) -Message 'Suite の ValidateSet が見つかりません。'
    Assert-True -Condition ($integrationText.Contains('[string]$Suite = ''smoke''')) -Message 'Suite の既定値が smoke ではありません。'
    Assert-True -Condition ($integrationText.Contains('function Get-SelectedTestCases')) -Message 'Suite 切り替え関数が見つかりません。'
    Assert-True -Condition ($integrationText.Contains('- 実行スイート: {0}')) -Message 'レポートの実行スイート表示が見つかりません。'
    return 'Suite switch / report summary を確認'
}

$testResults += Invoke-UnitTest -Name 'run.lock は匿名化され、Secure cleanup 失敗時の警告を持つ' -Body {
    $scriptText = Get-Content -LiteralPath $runScript -Raw -Encoding UTF8
    Assert-True -Condition (-not $scriptText.Contains('machineName')) -Message 'run.lock に machineName が残っています。'
    Assert-True -Condition (-not $scriptText.Contains('userName')) -Message 'run.lock に userName が残っています。'
    Assert-True -Condition ($scriptText.Contains('端末再起動または管理者確認が必要')) -Message 'cleanup 失敗時の警告文が見つかりません。'
    Assert-True -Condition ($scriptText.Contains('throw $script:cleanupFailureMessage')) -Message 'cleanup 失敗時の非ゼロ終了分岐が見つかりません。'
    return 'run.lock 匿名化 / cleanup 警告強化を確認'
}

$testResults += Invoke-UnitTest -Name 'Write-RunReport は既存ファイルを原子的に置き換える' -Body {
    $tempReportPath = Join-Path $resultsRoot 'atomic-run-report.json'
    $originalRunReportPath = $RunReportPath
    $originalLogPath = $script:logPath

    try {
        [System.IO.File]::WriteAllText($tempReportPath, '{broken json', (New-Object System.Text.UTF8Encoding($false)))
        $RunReportPath = $tempReportPath
        $script:logPath = Join-Path $resultsRoot 'atomic-run-report.log'
        [System.IO.File]::WriteAllText($script:logPath, 'log', (New-Object System.Text.UTF8Encoding($false)))

        Write-RunReport -Status 'Success' -OutputPath 'C:\Temp\sample.xlsx' -OutputParent 'C:\Temp' -Profile ([pscustomobject]@{ Name = 'profile_a'; DisplayName = 'Profile A' }) -SourcePdfCount 1 -RunHistoryPath 'C:\Temp\run-history.csv'
        $report = Get-Content -LiteralPath $tempReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
        Assert-True -Condition ($report.Status -eq 'Success') -Message "RunReport の Status が想定と異なります: $($report.Status)"
        Assert-True -Condition ($report.ProfileName -eq 'profile_a') -Message "RunReport の ProfileName が想定と異なります: $($report.ProfileName)"
        $tempFiles = @(Get-ChildItem -LiteralPath $resultsRoot -Filter '*.tmp' -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -like '*atomic-run-report*' })
        Assert-True -Condition ($tempFiles.Count -eq 0) -Message 'RunReport の一時ファイルが残っています。'
        return 'RunReport の原子的置換を確認'
    } finally {
        $RunReportPath = $originalRunReportPath
        $script:logPath = $originalLogPath
        if (Test-Path -LiteralPath $tempReportPath) {
            Remove-Item -LiteralPath $tempReportPath -Force -ErrorAction SilentlyContinue
        }
    }
}

$testResults += Invoke-UnitTest -Name 'Append-RunHistory は既存 CSV を原子的に追記する' -Body {
    $historyTestRoot = Join-Path $resultsRoot 'atomic-run-history'
    $tempHistoryPath = Join-Path $historyTestRoot 'run-history.csv'
    $originalRunHistoryPath = $script:runHistoryPath
    $originalOutputFile = $OutputFile

    try {
        if (Test-Path -LiteralPath $historyTestRoot) {
            Remove-Item -LiteralPath $historyTestRoot -Recurse -Force -ErrorAction SilentlyContinue
        }

        Ensure-Directory -Path $historyTestRoot
        $script:runHistoryPath = $tempHistoryPath
        $OutputFile = 'C:\Temp\first.xlsx'

        Append-RunHistory -Status 'SUCCESS'
        $runPlan = [pscustomobject]@{
            Profile     = [pscustomobject]@{ Name = 'profile_b'; DisplayName = 'Profile B' }
            SourceFiles = @('a.pdf', 'b.pdf')
            OutputPath  = 'C:\Temp\second.xlsx'
        }
        Append-RunHistory -Status 'SUCCESS' -RunPlan $runPlan -SuccessPdfCount 2

        $rows = @(Import-Csv -LiteralPath $tempHistoryPath)
        Assert-True -Condition ($rows.Count -eq 2) -Message "run-history の件数が想定と異なります: $($rows.Count)"
        Assert-True -Condition ($rows[0].Status -eq 'SUCCESS') -Message "1件目の Status が想定と異なります: $($rows[0].Status)"
        Assert-True -Condition ($rows[1].ProfileName -eq 'profile_b') -Message "2件目の ProfileName が想定と異なります: $($rows[1].ProfileName)"
        Assert-True -Condition ($rows[1].OutputPath -eq 'C:\Temp\second.xlsx') -Message "2件目の OutputPath が想定と異なります: $($rows[1].OutputPath)"

        $tempFiles = @(Get-ChildItem -LiteralPath $historyTestRoot -Filter '*.tmp' -File -ErrorAction SilentlyContinue)
        Assert-True -Condition ($tempFiles.Count -eq 0) -Message 'run-history の一時ファイルが残っています。'
        return 'run-history の原子的追記を確認'
    } finally {
        $script:runHistoryPath = $originalRunHistoryPath
        $OutputFile = $originalOutputFile
        if (Test-Path -LiteralPath $historyTestRoot) {
            Remove-Item -LiteralPath $historyTestRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

$testResults += Invoke-UnitTest -Name 'Write-EnvironmentCheckReport は既存レポートを原子的に置き換える' -Body {
    $tempReportPath = Join-Path $resultsRoot 'atomic-environment-check.md'
    $originalEnvironmentReportPath = $script:environmentReportPath

    try {
        [System.IO.File]::WriteAllText($tempReportPath, 'broken report', (New-Object System.Text.UTF8Encoding($false)))
        $script:environmentReportPath = $tempReportPath

        $environmentResult = [pscustomobject]@{
            Items = @(
                [pscustomobject]@{
                    Name = 'Excel COM'
                    Status = 'OK'
                    Detail = '起動可能'
                    SuggestedAction = ''
                }
            )
            HasFailures = $false
            Summary = '主要な前提条件は満たしています。'
        }

        Write-EnvironmentCheckReport -EnvironmentResult $environmentResult
        $reportText = Get-Content -LiteralPath $tempReportPath -Raw -Encoding UTF8
        Assert-True -Condition ($reportText.Contains('PDF2Excel 環境チェックレポート')) -Message '環境チェックレポートのタイトルがありません。'
        Assert-True -Condition ($reportText.Contains('Excel COM')) -Message '環境チェックレポートの項目がありません。'
        $tempFiles = @(Get-ChildItem -LiteralPath $resultsRoot -Filter '*.tmp' -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -like '*atomic-environment-check*' })
        Assert-True -Condition ($tempFiles.Count -eq 0) -Message '環境チェックレポートの一時ファイルが残っています。'
        return '環境チェックレポートの原子的置換を確認'
    } finally {
        $script:environmentReportPath = $originalEnvironmentReportPath
        if (Test-Path -LiteralPath $tempReportPath) {
            Remove-Item -LiteralPath $tempReportPath -Force -ErrorAction SilentlyContinue
        }
    }
}

$testResults += Invoke-UnitTest -Name 'Get-RunLockState は stale lock を識別する' -Body {
    $tempLockPath = Join-Path $resultsRoot 'stale-run.lock'
    $originalLockPath = $script:lockFilePath

    try {
        $script:lockFilePath = $tempLockPath
        $staleLock = [ordered]@{
            runInstanceId = 'stale-lock'
            pid           = 999999
            startedAt     = (Get-Date).AddHours(-1).ToString('yyyy-MM-dd HH:mm:ss')
        } | ConvertTo-Json
        [System.IO.File]::WriteAllText($tempLockPath, $staleLock, (New-Object System.Text.UTF8Encoding($false)))

        $state = Get-RunLockState
        Assert-True -Condition ($state.Classification -eq 'STALE') -Message "stale lock の分類が想定と異なります: $($state.Classification)"
        Assert-True -Condition ($state.Detail.Contains('前回異常終了の可能性')) -Message "stale lock の詳細が想定と異なります: $($state.Detail)"
        return 'stale lock の識別を確認'
    } finally {
        $script:lockFilePath = $originalLockPath
        if (Test-Path -LiteralPath $tempLockPath) {
            Remove-Item -LiteralPath $tempLockPath -Force -ErrorAction SilentlyContinue
        }
    }
}

$testResults += Invoke-UnitTest -Name 'Get-RunLockState は壊れた lock を内容読取不可として保持する' -Body {
    $tempLockPath = Join-Path $resultsRoot 'broken-run.lock'
    $originalLockPath = $script:lockFilePath

    try {
        $script:lockFilePath = $tempLockPath
        [System.IO.File]::WriteAllText($tempLockPath, '{broken lock', (New-Object System.Text.UTF8Encoding($false)))

        $state = Get-RunLockState
        Assert-True -Condition ($state.Classification -eq 'UNREADABLE') -Message "壊れた lock の分類が想定と異なります: $($state.Classification)"
        Assert-True -Condition ($state.Detail.Contains('内容を読めませんでした')) -Message "壊れた lock の詳細が想定と異なります: $($state.Detail)"
        return '壊れた lock の unreadable 判定を確認'
    } finally {
        $script:lockFilePath = $originalLockPath
        if (Test-Path -LiteralPath $tempLockPath) {
            Remove-Item -LiteralPath $tempLockPath -Force -ErrorAction SilentlyContinue
        }
    }
}

$testResults += Invoke-UnitTest -Name 'Acquire-RunLock は放棄 mutex 回復の catch 分岐を持つ' -Body {
    $scriptText = Get-Content -LiteralPath $runScript -Raw -Encoding UTF8
    Assert-True -Condition ($scriptText.Contains('catch [System.Threading.AbandonedMutexException]')) -Message 'AbandonedMutexException の catch が見つかりません。'
    Assert-True -Condition ($scriptText.Contains('自動回復して処理を続行します')) -Message '放棄 mutex 回復時の警告ログが見つかりません。'
    return '放棄 mutex 回復分岐を確認'
}

$testResults += Invoke-UnitTest -Name 'テンプレート整合性マニフェストは必要ファイルを持つ' -Body {
    $manifestPath = Join-Path $projectRoot 'config\template-integrity.json'
    Assert-True -Condition (Test-Path -LiteralPath $manifestPath) -Message 'template-integrity.json が見つかりません。'
    $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $relativePaths = @($manifest.Files | ForEach-Object { [string]$_.RelativePath })
    foreach ($requiredPath in @(
        'template\PDF2Excel_V1_Converter.xlsm',
        'template\PDF2Excel_V2_Converter.xlsm',
        'template\vba\PDF2ExcelMacros.bas',
        'template\vba\PDF2ExcelTemplateBuilder.bas'
    )) {
        Assert-True -Condition ($relativePaths -contains $requiredPath) -Message "マニフェストに必要ファイルがありません: $requiredPath"
    }
    return 'template-integrity.json の主要エントリを確認'
}

$testResults += Invoke-UnitTest -Name '日本語コンソール出力スクリプトは UTF-8 初期化と保存形式を持つ' -Body {
    $scriptPaths = @(
        (Join-Path $projectRoot 'scripts\run_pdf2excel.ps1'),
        (Join-Path $projectRoot 'scripts\build_handoff_package.ps1'),
        (Join-Path $projectRoot 'scripts\build_excel_template.ps1'),
        (Join-Path $projectRoot 'scripts\new_profile_scaffold.ps1')
    )

    foreach ($scriptPath in $scriptPaths) {
        $scriptText = Get-Content -LiteralPath $scriptPath -Raw -Encoding UTF8
        Assert-True -Condition ($scriptText.Contains('[Console]::InputEncoding = New-Object System.Text.UTF8Encoding($false)')) -Message "$scriptPath に InputEncoding 初期化がありません。"
        Assert-True -Condition ($scriptText.Contains('[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)')) -Message "$scriptPath に OutputEncoding 初期化がありません。"
        Assert-True -Condition ($scriptText.Contains('$OutputEncoding = [Console]::OutputEncoding')) -Message "$scriptPath に OutputEncoding 反映がありません。"
    }

    $buildHandoffBytes = [System.IO.File]::ReadAllBytes((Join-Path $projectRoot 'scripts\build_handoff_package.ps1'))
    Assert-True -Condition ($buildHandoffBytes.Length -ge 3) -Message 'build_handoff_package.ps1 が空です。'
    Assert-True -Condition ($buildHandoffBytes[0] -eq 0xEF -and $buildHandoffBytes[1] -eq 0xBB -and $buildHandoffBytes[2] -eq 0xBF) -Message 'build_handoff_package.ps1 が UTF-8 BOM ではありません。'
    return 'run/build/scaffold scripts の UTF-8 初期化と build_handoff_package.ps1 の BOM を確認'
}

$testResults += Invoke-UnitTest -Name 'Write-Log は INFO で詳細パスを伏せ DEBUG を抑止する' -Body {
    $tempLogPath = Join-Path $resultsRoot 'write-log-info.log'
    $originalLogPath = $script:logPath
    $originalSensitiveMap = $script:sensitivePathMap
    $originalSecureMode = $script:isSecureMode
    $originalLogLevel = $LogLevel

    try {
        if (Test-Path -LiteralPath $tempLogPath) {
            Remove-Item -LiteralPath $tempLogPath -Force
        }

        $script:logPath = $tempLogPath
        $script:sensitivePathMap = [ordered]@{
            'C:\Users\TestUser\AppData\Local\PDF2Excel\logs\run_001.log' = '...\logs\run_001.log'
        }
        $script:isSecureMode = $true
        $LogLevel = 'INFO'

        Write-Log -Message 'INFO path=C:\Users\TestUser\AppData\Local\PDF2Excel\logs\run_001.log' -Level 'INFO'
        Write-Log -Message 'DEBUG path=C:\Users\TestUser\AppData\Local\PDF2Excel\logs\run_001.log' -Level 'DEBUG'

        $logText = Get-Content -LiteralPath $tempLogPath -Raw -Encoding UTF8
        Assert-True -Condition ($logText.Contains('...\logs\run_001.log')) -Message "INFO ログのマスクが効いていません: $logText"
        Assert-True -Condition (-not $logText.Contains('C:\Users\TestUser\AppData\Local\PDF2Excel\logs\run_001.log')) -Message "INFO ログにフルパスが残っています: $logText"
        Assert-True -Condition (-not $logText.Contains('DEBUG path=')) -Message "INFO レベルで DEBUG ログが出力されています: $logText"
        return 'INFO では path mask / DEBUG 抑止を確認'
    } finally {
        $script:logPath = $originalLogPath
        $script:sensitivePathMap = $originalSensitiveMap
        $script:isSecureMode = $originalSecureMode
        $LogLevel = $originalLogLevel
    }
}

$testResults += Invoke-UnitTest -Name 'Write-Log は DEBUG 指定時だけ詳細パスを出力する' -Body {
    $tempLogPath = Join-Path $resultsRoot 'write-log-debug.log'
    $originalLogPath = $script:logPath
    $originalSensitiveMap = $script:sensitivePathMap
    $originalSecureMode = $script:isSecureMode
    $originalLogLevel = $LogLevel

    try {
        if (Test-Path -LiteralPath $tempLogPath) {
            Remove-Item -LiteralPath $tempLogPath -Force
        }

        $script:logPath = $tempLogPath
        $script:sensitivePathMap = [ordered]@{
            'C:\Users\TestUser\AppData\Local\PDF2Excel\logs\run_001.log' = '...\logs\run_001.log'
        }
        $script:isSecureMode = $true
        $LogLevel = 'DEBUG'

        Write-Log -Message 'DEBUG path=C:\Users\TestUser\AppData\Local\PDF2Excel\logs\run_001.log' -Level 'DEBUG'

        $logText = Get-Content -LiteralPath $tempLogPath -Raw -Encoding UTF8
        Assert-True -Condition ($logText.Contains('DEBUG path=C:\Users\TestUser\AppData\Local\PDF2Excel\logs\run_001.log')) -Message "DEBUG ログで詳細パスが出力されていません: $logText"
        Assert-True -Condition (-not $logText.Contains('...\logs\run_001.log')) -Message "DEBUG ログで path mask が残っています: $logText"
        return 'DEBUG では詳細パスを確認'
    } finally {
        $script:logPath = $originalLogPath
        $script:sensitivePathMap = $originalSensitiveMap
        $script:isSecureMode = $originalSecureMode
        $LogLevel = $originalLogLevel
    }
}

$testResults += Invoke-UnitTest -Name 'ログ要約関数は件数と代表ファイル名を返す' -Body {
    $summary = Get-PathListLogSummary -Paths @(
        'C:\temp\alpha.pdf',
        'C:\temp\beta.pdf',
        'C:\temp\gamma.pdf',
        'C:\temp\delta.pdf'
    )
    Assert-True -Condition ($summary.Contains('4 件')) -Message "件数が想定と異なります: $summary"
    Assert-True -Condition ($summary.Contains('alpha.pdf')) -Message "代表ファイル名が含まれていません: $summary"
    Assert-True -Condition ($summary.Contains('ほか 1 件')) -Message "省略件数が含まれていません: $summary"
    return $summary
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
    $batPaths = @(
        (Join-Path $legacyRoot 'run_pdf2excel_v1.bat'),
        (Join-Path $projectRoot 'run_pdf2excel_v2.bat')
    )
    foreach ($batPath in $batPaths) {
        $batName = Split-Path -Leaf $batPath
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
    $menuPaths = @(
        (Join-Path $legacyRoot 'scripts\run_pdf2excel_menu_v1.ps1'),
        (Join-Path $projectRoot 'scripts\run_pdf2excel_menu_v2.ps1')
    )
    foreach ($menuPath in $menuPaths) {
        $menuName = Split-Path -Leaf $menuPath
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
