Set-StrictMode -Version Latest

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

function Reset-Directory {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (Test-Path -LiteralPath $Path) {
        try {
            Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
        } catch [System.IO.DirectoryNotFoundException] {
        }
    }

    New-Item -ItemType Directory -Path $Path -Force | Out-Null
}

function New-EmptyFile {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType File -Path $Path | Out-Null
    }
}

function Escape-MString {
    param([Parameter(Mandatory = $true)][string]$Value)

    return $Value.Replace('"', '""')
}

function ConvertTo-MTextLiteral {
    param([Parameter(Mandatory = $true)][string]$Value)

    return '"' + (Escape-MString -Value $Value) + '"'
}

function ConvertTo-MTextListLiteral {
    param([string[]]$Values)

    if ($null -eq $Values -or $Values.Count -eq 0) {
        return '{}'
    }

    $items = @($Values | ForEach-Object { ConvertTo-MTextLiteral -Value $_ })
    return '{' + ($items -join ', ') + '}'
}

function ConvertTo-MLogicalLiteral {
    param([bool]$Value)

    if ($Value) {
        return 'true'
    }

    return 'false'
}

function ConvertTo-MRecordListLiteral {
    param([object[]]$Records)

    if ($null -eq $Records -or $Records.Count -eq 0) {
        return '{}'
    }

    $items = @()
    foreach ($record in $Records) {
        $properties = @()
        foreach ($property in $record.PSObject.Properties) {
            $value = $property.Value
            $literal =
                if ($null -eq $value) {
                    'null'
                } elseif ($value -is [bool]) {
                    ConvertTo-MLogicalLiteral -Value $value
                } elseif ($value -is [int] -or $value -is [long] -or $value -is [double] -or $value -is [decimal]) {
                    [string]$value
                } else {
                    ConvertTo-MTextLiteral -Value ([string]$value)
                }
            $properties += ('{0} = {1}' -f $property.Name, $literal)
        }
        $items += ('[' + ($properties -join ', ') + ']')
    }

    return '{' + ($items -join ', ') + '}'
}

function Get-ProfileOutputColumnNames {
    param([Parameter(Mandatory = $true)]$Profile)

    $names = @($Profile.SourceFileColumnName)
    foreach ($index in 1..$Profile.ExpectedColumns) {
        $names += '{0}{1}' -f $Profile.DataColumnPrefix, $index
    }

    return $names
}

function Convert-MinutesToTimeText {
    param([Parameter(Mandatory = $true)][int]$MinutesFromMidnight)

    $hours = [int][Math]::Floor($MinutesFromMidnight / 60)
    $minutes = [int]($MinutesFromMidnight % 60)
    return ('{0:D2}:{1:D2}' -f $hours, $minutes)
}

function Count-TimeLikeTokens {
    param([AllowNull()]$Value)

    if ($null -eq $Value) {
        return 0
    }

    $normalized = ([string]$Value).Replace("`r", ' ').Replace("`n", ' ').Replace('　', ' ')
    $tokens = @($normalized -split '\s+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    return @(
        $tokens | Where-Object {
            $_ -match ':' -or
            $_ -match '：' -or
            $_ -match '時'
        }
    ).Count
}

function Normalize-TimeText {
    param([AllowNull()]$Value)

    if ($null -eq $Value) {
        return [pscustomobject]@{
            RawText              = ''
            CandidateText        = ''
            NormalizedText       = ''
            MinutesFromMidnight  = $null
            ExcelTimeValue       = $null
            Status               = 'EMPTY'
            Note                 = ''
        }
    }

    if ($Value -isnot [string] -and $Value -is [System.IConvertible]) {
        try {
            $numericValue = [double]$Value
            if ($numericValue -ge 0 -and $numericValue -le 1) {
                $minutesFromMidnight = [int][Math]::Round($numericValue * 1440, 0, [MidpointRounding]::AwayFromZero)
                if ($minutesFromMidnight -ge 0 -and $minutesFromMidnight -le 1440) {
                    return [pscustomobject]@{
                        RawText              = [string]$Value
                        CandidateText        = [string]$Value
                        NormalizedText       = Convert-MinutesToTimeText -MinutesFromMidnight $minutesFromMidnight
                        MinutesFromMidnight  = $minutesFromMidnight
                        ExcelTimeValue       = $minutesFromMidnight / 1440.0
                        Status               = 'OK'
                        Note                 = ''
                    }
                }
            }
        } catch {
        }
    }

    $rawText = [string]$Value
    $candidate = $rawText.Trim()
    if ([string]::IsNullOrWhiteSpace($candidate)) {
        return [pscustomobject]@{
            RawText              = $rawText
            CandidateText        = ''
            NormalizedText       = ''
            MinutesFromMidnight  = $null
            ExcelTimeValue       = $null
            Status               = 'EMPTY'
            Note                 = ''
        }
    }

    $widthMap = @{
        '０' = '0'; '１' = '1'; '２' = '2'; '３' = '3'; '４' = '4'
        '５' = '5'; '６' = '6'; '７' = '7'; '８' = '8'; '９' = '9'
    }
    foreach ($entry in $widthMap.GetEnumerator()) {
        $candidate = $candidate.Replace($entry.Key, $entry.Value)
    }

    $candidate = $candidate.Replace('：', ':')

    $fractionValue = 0.0
    if ($candidate.Contains('.') -and [double]::TryParse($candidate, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$fractionValue)) {
        if ($fractionValue -ge 0 -and $fractionValue -le 1) {
            $minutesFromMidnight = [int][Math]::Round($fractionValue * 1440, 0, [MidpointRounding]::AwayFromZero)
            return [pscustomobject]@{
                RawText              = $rawText
                CandidateText        = $candidate
                NormalizedText       = Convert-MinutesToTimeText -MinutesFromMidnight $minutesFromMidnight
                MinutesFromMidnight  = $minutesFromMidnight
                ExcelTimeValue       = $minutesFromMidnight / 1440.0
                Status               = 'OK'
                Note                 = ''
            }
        }
    }

    $candidate = $candidate.Replace('時', ':')
    $candidate = $candidate.Replace('分', '')
    $candidate = $candidate.Replace('頃', '')
    $candidate = $candidate.Replace('.', ':')
    $candidate = $candidate.Replace('　', '')
    $candidate = $candidate -replace '\s+', ''

    if ([string]::IsNullOrWhiteSpace($candidate)) {
        return [pscustomobject]@{
            RawText              = $rawText
            CandidateText        = ''
            NormalizedText       = ''
            MinutesFromMidnight  = $null
            ExcelTimeValue       = $null
            Status               = 'EMPTY'
            Note                 = ''
        }
    }

    if ($candidate -match '^\d{1,2}:$') {
        $candidate = $candidate + '00'
    } elseif ($candidate -match '^\d{1,2}$') {
        $candidate = $candidate + ':00'
    } elseif ($candidate -match '^\d{3,4}$') {
        $candidate = $candidate.Insert($candidate.Length - 2, ':')
    } elseif ($candidate -match '^(?<hours>\d{1,2}):(?<minutes>\d{1,2}):(?<seconds>\d{1,2})$') {
        if ([int]$Matches.seconds -ne 0) {
            return [pscustomobject]@{
                RawText              = $rawText
                CandidateText        = $candidate
                NormalizedText       = ''
                MinutesFromMidnight  = $null
                ExcelTimeValue       = $null
                Status               = 'INVALID'
                Note                 = '秒を含む時刻は 00 秒のみ補助正規化の対象です。'
            }
        }

        $candidate = ('{0}:{1}' -f $Matches.hours, $Matches.minutes)
    }

    if ($candidate -notmatch '^\d{1,2}:\d{1,2}$') {
        return [pscustomobject]@{
            RawText              = $rawText
            CandidateText        = $candidate
            NormalizedText       = ''
            MinutesFromMidnight  = $null
            ExcelTimeValue       = $null
            Status               = 'INVALID'
            Note                 = '時刻として解釈できません。'
        }
    }

    $parts = $candidate.Split(':')
    $hours = 0
    $minutes = 0
    if (-not [int]::TryParse($parts[0], [ref]$hours) -or -not [int]::TryParse($parts[1], [ref]$minutes)) {
        return [pscustomobject]@{
            RawText              = $rawText
            CandidateText        = $candidate
            NormalizedText       = ''
            MinutesFromMidnight  = $null
            ExcelTimeValue       = $null
            Status               = 'INVALID'
            Note                 = '時刻として解釈できません。'
        }
    }

    if ($hours -eq 24 -and $minutes -eq 0) {
        $minutesFromMidnight = 1440
        return [pscustomobject]@{
            RawText              = $rawText
            CandidateText        = $candidate
            NormalizedText       = '24:00'
            MinutesFromMidnight  = $minutesFromMidnight
            ExcelTimeValue       = 1.0
            Status               = 'OK'
            Note                 = ''
        }
    }

    if ($hours -lt 0 -or $hours -gt 23 -or $minutes -lt 0 -or $minutes -gt 59) {
        return [pscustomobject]@{
            RawText              = $rawText
            CandidateText        = $candidate
            NormalizedText       = ''
            MinutesFromMidnight  = $null
            ExcelTimeValue       = $null
            Status               = 'INVALID'
            Note                 = '時刻の範囲外です。'
        }
    }

    $minutesFromMidnight = ($hours * 60) + $minutes
    return [pscustomobject]@{
        RawText              = $rawText
        CandidateText        = $candidate
        NormalizedText       = ('{0:D2}:{1:D2}' -f $hours, $minutes)
        MinutesFromMidnight  = $minutesFromMidnight
        ExcelTimeValue       = $minutesFromMidnight / 1440.0
        Status               = 'OK'
        Note                 = ''
    }
}

function Get-TimeNormalizationAudit {
    param(
        [Parameter(Mandatory = $true)][object[]]$Definitions,
        [Parameter(Mandatory = $true)][hashtable]$RawValuesByDisplayName,
        [string]$ExistingReason = ''
    )

    $results = @{}
    $issues = @()
    $hasTimeValue = $false

    foreach ($definition in $Definitions) {
        $rawValue = if ($RawValuesByDisplayName.ContainsKey($definition.DisplayName)) { [string]$RawValuesByDisplayName[$definition.DisplayName] } else { '' }
        $normalized = Normalize-TimeText -Value $rawValue
        $results[$definition.DisplayName] = $normalized

        if (-not [string]::IsNullOrWhiteSpace($rawValue)) {
            $hasTimeValue = $true
        }

        if ((Count-TimeLikeTokens -Value $rawValue) -gt 1) {
            $issues += ('{0}: 複数の時刻らしき文字列があります。' -f $definition.DisplayName)
        }

        if ($normalized.Status -eq 'INVALID') {
            $issues += ('{0}: {1}' -f $definition.DisplayName, $normalized.Note)
        }
    }

    for ($index = 0; $index -lt $Definitions.Count; $index += 2) {
        $left = $Definitions[$index]
        $right = if (($index + 1) -lt $Definitions.Count) { $Definitions[$index + 1] } else { $null }
        if ($null -eq $right) {
            continue
        }

        $leftValue = if ($RawValuesByDisplayName.ContainsKey($left.DisplayName)) { [string]$RawValuesByDisplayName[$left.DisplayName] } else { '' }
        $rightValue = if ($RawValuesByDisplayName.ContainsKey($right.DisplayName)) { [string]$RawValuesByDisplayName[$right.DisplayName] } else { '' }
        $leftHasValue = -not [string]::IsNullOrWhiteSpace($leftValue)
        $rightHasValue = -not [string]::IsNullOrWhiteSpace($rightValue)
        if ($leftHasValue -xor $rightHasValue) {
            $issues += ('{0}/{1}: 片側の時刻だけ埋まっています。' -f $left.DisplayName, $right.DisplayName)
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($ExistingReason)) {
        $issues += $ExistingReason
    }

    $distinctIssues = @($issues | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
    return [pscustomobject]@{
        Results      = $results
        HasTimeValue = $hasTimeValue
        Status       = if ($distinctIssues.Count -gt 0) { '要確認' } elseif ($hasTimeValue) { 'OK' } else { '' }
        Note         = if ($distinctIssues.Count -gt 0) { $distinctIssues -join ' / ' } else { '' }
    }
}

function Get-ReviewReasonCategories {
    param(
        [Parameter(Mandatory = $true)][object[]]$Definitions,
        [Parameter(Mandatory = $true)][hashtable]$RawValuesByDisplayName,
        [string]$ExistingReason = '',
        [string]$ExistingCategoryCsv = '',
        $Audit = $null
    )

    if ($null -eq $Audit) {
        $Audit = Get-TimeNormalizationAudit -Definitions $Definitions -RawValuesByDisplayName $RawValuesByDisplayName
    }

    $categorySet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($category in @($ExistingCategoryCsv -split ',')) {
        $trimmed = [string]$category
        if (-not [string]::IsNullOrWhiteSpace($trimmed)) {
            [void]$categorySet.Add($trimmed.Trim())
        }
    }

    if ($ExistingReason -like '*安全に結合できませんでした*') {
        [void]$categorySet.Add('HEADER_MISMATCH')
    }

    foreach ($definition in $Definitions) {
        $rawValue = if ($RawValuesByDisplayName.ContainsKey($definition.DisplayName)) { [string]$RawValuesByDisplayName[$definition.DisplayName] } else { '' }
        if ((Count-TimeLikeTokens -Value $rawValue) -gt 1) {
            [void]$categorySet.Add('TIME_MULTI')
        }

        $normalized = if ($null -ne $Audit -and $Audit.Results.ContainsKey($definition.DisplayName)) { $Audit.Results[$definition.DisplayName] } else { Normalize-TimeText -Value $rawValue }
        if ($normalized.Status -eq 'INVALID') {
            [void]$categorySet.Add('TIME_INVALID')
        }
    }

    for ($index = 0; $index -lt $Definitions.Count; $index += 2) {
        $left = $Definitions[$index]
        $right = if (($index + 1) -lt $Definitions.Count) { $Definitions[$index + 1] } else { $null }
        if ($null -eq $right) {
            continue
        }

        $leftValue = if ($RawValuesByDisplayName.ContainsKey($left.DisplayName)) { [string]$RawValuesByDisplayName[$left.DisplayName] } else { '' }
        $rightValue = if ($RawValuesByDisplayName.ContainsKey($right.DisplayName)) { [string]$RawValuesByDisplayName[$right.DisplayName] } else { '' }
        $leftHasValue = -not [string]::IsNullOrWhiteSpace($leftValue)
        $rightHasValue = -not [string]::IsNullOrWhiteSpace($rightValue)
        if ($leftHasValue -xor $rightHasValue) {
            [void]$categorySet.Add('TIME_MISSING')
        }
    }

    $preferredOrder = @('HEADER_MISMATCH', 'TIME_MULTI', 'TIME_MISSING', 'TIME_INVALID')
    $ordered = New-Object System.Collections.Generic.List[string]

    foreach ($category in $preferredOrder) {
        if ($categorySet.Contains($category)) {
            [void]$ordered.Add($category)
        }
    }

    foreach ($category in $categorySet) {
        if ($preferredOrder -notcontains $category) {
            [void]$ordered.Add($category)
        }
    }

    return ($ordered.ToArray() -join ',')
}

function Get-NormalizedTimeColumnDefinitions {
    param(
        [Parameter(Mandatory = $true)]$Profile,
        [ValidateSet('v1', 'v2')]
        [string]$VersionMode = 'v1'
    )

    if ($VersionMode -ne 'v2') {
        return @()
    }

    $definitions = @()
    $profileDefinitions = $Profile.PSObject.Properties['NormalizedTimeColumns']
    if ($null -ne $profileDefinitions -and $null -ne $profileDefinitions.Value) {
        foreach ($entry in @($profileDefinitions.Value)) {
            if ($null -eq $entry) {
                continue
            }

            $sourceColumn = [int]$entry.SourceColumn
            $displayName = [string]$entry.DisplayName
            if ($sourceColumn -lt 1 -or [string]::IsNullOrWhiteSpace($displayName)) {
                continue
            }

            $minutesColumnName = if ([string]::IsNullOrWhiteSpace([string]$entry.MinutesColumnName)) {
                '{0}_分' -f $displayName
            } else {
                [string]$entry.MinutesColumnName
            }

            $definitions += [pscustomobject]@{
                SourceColumn        = $sourceColumn
                SourceColumnName    = '{0}{1}' -f $Profile.DataColumnPrefix, $sourceColumn
                DisplayName         = $displayName
                MinutesColumnName   = $minutesColumnName
                ReviewRawColumnName = '{0}_raw' -f $displayName
            }
        }
    }

    return $definitions
}

function Get-ResultOutputColumnNames {
    param(
        [Parameter(Mandatory = $true)]$Profile,
        [ValidateSet('v1', 'v2')]
        [string]$VersionMode = 'v1'
    )

    $names = @(Get-ProfileOutputColumnNames -Profile $Profile)
    foreach ($definition in @(Get-NormalizedTimeColumnDefinitions -Profile $Profile -VersionMode $VersionMode)) {
        $names += $definition.DisplayName
        $names += $definition.MinutesColumnName
    }

    if ($VersionMode -eq 'v2' -and @($names).Count -gt $Profile.ExpectedColumns + 1) {
        $names += '時刻正規化状態'
        $names += '時刻確認メモ'
    }

    return $names
}

function Get-ControlSheetStaticCells {
    param(
        [ValidateSet('v1', 'v2')]
        [string]$VersionMode = 'v1'
    )

    $usageText = if ($VersionMode -eq 'v2') {
        '1. run_pdf2excel_v2.bat を実行  2. 建設現場向け PDF を選択  3. 実行前チェックを確認  4. Result / Review / Summary / Errors を確認'
    } else {
        '1. run_pdf2excel_v1.bat を実行  2. PDF を選択  3. 実行前チェックを確認  4. Result / Summary / Errors を確認'
    }

    $checkText = if ($VersionMode -eq 'v2') {
        'Summary は全体件数、Result は raw 転記、Review は確認要行、Errors は失敗した PDF と理由です。'
    } else {
        'Summary は件数の全体像、Result は変換成功データ、Errors は失敗した PDF と理由です。'
    }

    return [ordered]@{
        'A1'  = '項目'
        'B1'  = '内容'
        'A2'  = '入力フォルダ'
        'A3'  = '出力ファイル'
        'A4'  = 'ログファイル'
        'A5'  = '最終実行日時'
        'A6'  = '状態'
        'A7'  = '対象PDF数'
        'A8'  = '取込データ行数'
        'A9'  = 'エラー件数'
        'A10' = '成功PDF数'
        'A11' = '失敗PDF数'
        'A12' = '処理時間(秒)'
        'A13' = '使用プロファイル'
        'A14' = 'プロファイル説明'
        'A15' = 'システム版'
        'A16' = 'かんたんな使い方'
        'B16' = $usageText
        'A17' = '確認ポイント'
        'B17' = $checkText
        'A18' = '注意'
        'B18' = '文字を選択できるテキスト PDF と、ほぼ同じレイアウトの帳票を想定しています。'
    }
}

function Set-ControlSheetStaticCells {
    param(
        [Parameter(Mandatory = $true)]$Worksheet,
        [ValidateSet('v1', 'v2')]
        [string]$VersionMode = 'v1'
    )

    foreach ($entry in (Get-ControlSheetStaticCells -VersionMode $VersionMode).GetEnumerator()) {
        $Worksheet.Range($entry.Key).Value2 = $entry.Value
    }
}

function Get-SummarySheetStaticCells {
    param(
        [ValidateSet('v1', 'v2')]
        [string]$VersionMode = 'v1'
    )

    $summaryDescription = if ($VersionMode -eq 'v2') {
        '実行結果の集計、PDF ごとの内訳、Review 件数を表示します。'
    } else {
        '実行結果の集計と、PDFごとの内訳を表示します。'
    }

    return [ordered]@{
        'A1'  = 'Summary'
        'A2'  = $summaryDescription
        'A4'  = '項目'
        'B4'  = '内容'
        'A13' = 'PDF別サマリー'
        'M13' = 'エラー分類別件数'
    }
}

function Set-SummarySheetStaticCells {
    param(
        [Parameter(Mandatory = $true)]$Worksheet,
        [ValidateSet('v1', 'v2')]
        [string]$VersionMode = 'v1'
    )

    foreach ($entry in (Get-SummarySheetStaticCells -VersionMode $VersionMode).GetEnumerator()) {
        $Worksheet.Range($entry.Key).Value2 = $entry.Value
    }
}

function Get-VersionDisplayName {
    param(
        [ValidateSet('v1', 'v2')]
        [Parameter(Mandatory = $true)][string]$VersionMode
    )

    if ($VersionMode -eq 'v2') {
        return 'VER2'
    }

    return 'VER1'
}

function Resolve-RunErrorInfo {
    param([Parameter(Mandatory = $true)][string]$Message)

    $normalizedMessage = $Message.ToLowerInvariant()
    if ($normalizedMessage.Contains('実行前チェックでキャンセル')) {
        return [pscustomobject]@{ ErrorCode = 'RUN_CANCELLED'; ErrorCategory = '実行キャンセル' }
    }
    if ($normalizedMessage.Contains('同名の pdf')) {
        return [pscustomobject]@{ ErrorCode = 'DUPLICATE_FILE_NAME'; ErrorCategory = '入力エラー' }
    }
    if ($normalizedMessage.Contains('別の pdf2excel 実行が進行中')) {
        return [pscustomobject]@{ ErrorCode = 'RUN_LOCKED'; ErrorCategory = '実行競合' }
    }
    if ($normalizedMessage.Contains('入力フォルダ') -or $normalizedMessage.Contains('入力ファイル')) {
        return [pscustomobject]@{ ErrorCode = 'INPUT_RESOLUTION_ERROR'; ErrorCategory = '入力エラー' }
    }
    if ($normalizedMessage.Contains('テンプレート')) {
        return [pscustomobject]@{ ErrorCode = 'TEMPLATE_ERROR'; ErrorCategory = 'テンプレートエラー' }
    }
    if ($normalizedMessage.Contains('excel')) {
        return [pscustomobject]@{ ErrorCode = 'EXCEL_RUNTIME_ERROR'; ErrorCategory = 'Excel実行エラー' }
    }

    return [pscustomobject]@{ ErrorCode = 'UNEXPECTED_RUN_ERROR'; ErrorCategory = 'システムエラー' }
}
