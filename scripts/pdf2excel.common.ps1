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

function Get-LocalAppDataPdf2ExcelPath {
    param([string]$ChildPath = '')

    if ([string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
        return $null
    }

    $basePath = Join-Path $env:LOCALAPPDATA 'PDF2Excel'
    if ([string]::IsNullOrWhiteSpace($ChildPath)) {
        return $basePath
    }

    return Join-Path $basePath $ChildPath
}

function Get-PathLocationInfo {
    param([Parameter(Mandatory = $true)][string]$Path)

    $normalizedPath = try {
        [System.IO.Path]::GetFullPath($Path)
    } catch {
        $Path
    }

    $rootPath = ''
    try {
        $rootPath = [System.IO.Path]::GetPathRoot($normalizedPath)
    } catch {
        $rootPath = ''
    }

    $isUnc = -not [string]::IsNullOrWhiteSpace($rootPath) -and $rootPath.StartsWith('\\')
    $driveType = ''
    $isNetworkDrive = $false

    if ($isUnc) {
        $driveType = [System.IO.DriveType]::Network.ToString()
    } elseif (-not [string]::IsNullOrWhiteSpace($rootPath)) {
        try {
            $driveInfo = New-Object System.IO.DriveInfo($rootPath)
            $driveType = $driveInfo.DriveType.ToString()
            $isNetworkDrive = ($driveInfo.DriveType -eq [System.IO.DriveType]::Network)
        } catch {
            $driveType = ''
        }
    }

    return [pscustomobject]@{
        NormalizedPath  = $normalizedPath
        RootPath        = $rootPath
        IsUnc           = $isUnc
        DriveType       = $driveType
        IsNetworkDrive  = $isNetworkDrive
        IsShared        = ($isUnc -or $isNetworkDrive)
    }
}

function ConvertTo-MaskedPathText {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $Path
    }

    $normalizedPath = try {
        [System.IO.Path]::GetFullPath($Path)
    } catch {
        $Path
    }

    $trimmedPath = $normalizedPath.TrimEnd('\')
    $leaf = try {
        Split-Path -Path $trimmedPath -Leaf
    } catch {
        $trimmedPath
    }

    $parentPath = try {
        Split-Path -Path $trimmedPath -Parent
    } catch {
        ''
    }

    $parentLeaf = if ([string]::IsNullOrWhiteSpace($parentPath)) {
        ''
    } else {
        try {
            Split-Path -Path $parentPath -Leaf
        } catch {
            ''
        }
    }

    if ([string]::IsNullOrWhiteSpace($leaf)) {
        return '...\'
    }

    if (-not [string]::IsNullOrWhiteSpace($parentLeaf) -and $parentLeaf -ne $leaf) {
        return ('...\{0}\{1}' -f $parentLeaf, $leaf)
    }

    return ('...\{0}' -f $leaf)
}

function Protect-MessagePaths {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [hashtable]$PathMap
    )

    $protectedMessage = $Message
    if ($PathMap) {
        $registeredPaths = @($PathMap.Keys | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object Length -Descending -Unique)
        foreach ($registeredPath in $registeredPaths) {
            $replacement = [string]$PathMap[$registeredPath]
            if ([string]::IsNullOrWhiteSpace($replacement)) {
                continue
            }

            $pattern = [System.Text.RegularExpressions.Regex]::Escape($registeredPath)
            $protectedMessage = [System.Text.RegularExpressions.Regex]::Replace(
                $protectedMessage,
                $pattern,
                [System.Text.RegularExpressions.MatchEvaluator]{ param($match) $replacement },
                [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
            )
        }
    }

    $genericPatterns = @(
        '(?i)(?:[A-Z]:\\|\\\\)[^\r\n''""]+?\.[A-Za-z0-9]{1,8}',
        '(?i)(?:[A-Z]:\\|\\\\)[^\r\n''""]+'
    )

    foreach ($pattern in $genericPatterns) {
        $protectedMessage = [System.Text.RegularExpressions.Regex]::Replace(
            $protectedMessage,
            $pattern,
            [System.Text.RegularExpressions.MatchEvaluator]{
                param($match)
                $rawValue = $match.Value
                $trimmedValue = $rawValue.TrimEnd(',', ';', '.', ')')
                $suffix = $rawValue.Substring($trimmedValue.Length)
                return (ConvertTo-MaskedPathText -Path $trimmedValue) + $suffix
            }
        )
    }

    return $protectedMessage
}

function Get-WindowProcessId {
    param([Parameter(Mandatory = $true)][int]$WindowHandle)

    if ($WindowHandle -le 0) {
        return $null
    }

    if (-not ('PDF2Excel.NativeMethods' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

namespace PDF2Excel {
    public static class NativeMethods {
        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
    }
}
'@
    }

    $processId = 0
    [PDF2Excel.NativeMethods]::GetWindowThreadProcessId([IntPtr]$WindowHandle, [ref]$processId) | Out-Null
    if ($processId -le 0) {
        return $null
    }

    return $processId
}

function Remove-PathWithRetryCommon {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [int]$MaxAttempts = 5,
        [int]$DelayMilliseconds = 500
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt += 1) {
        try {
            Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
            return
        } catch [System.IO.DirectoryNotFoundException] {
            return
        } catch [System.IO.FileNotFoundException] {
            return
        } catch {
            if ($attempt -ge $MaxAttempts) {
                throw
            }
            Start-Sleep -Milliseconds $DelayMilliseconds
        }
    }
}

function Reset-Directory {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (Test-Path -LiteralPath $Path) {
        Remove-PathWithRetryCommon -Path $Path
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

function Format-ProfileColumnDisplayName {
    param([AllowNull()]$Value)

    $text = if ($null -eq $Value) { '' } else { [string]$Value }
    $text = $text.Replace("`r", ' ').Replace("`n", ' ').Replace('　', ' ')
    $tokens = @($text -split '\s+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $normalized = ($tokens -join ' ').Trim()
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return '(空欄)'
    }

    return $normalized
}

function Get-UniqueColumnNameList {
    param([string[]]$Names)

    $counts = @{}
    $resolved = New-Object System.Collections.Generic.List[string]
    foreach ($name in @($Names)) {
        $baseName = Format-ProfileColumnDisplayName -Value $name
        $nextCount = if ($counts.ContainsKey($baseName)) { [int]$counts[$baseName] + 1 } else { 1 }
        $counts[$baseName] = $nextCount
        $finalName = if ($nextCount -le 1) { $baseName } else { '{0}_{1}' -f $baseName, $nextCount }
        [void]$resolved.Add($finalName)
    }

    return @($resolved)
}

function Test-ProfileHasCustomOutputColumnNames {
    param([Parameter(Mandatory = $true)]$Profile)

    return ($Profile.PSObject.Properties.Name -contains 'OutputColumnNames') -and $null -ne $Profile.OutputColumnNames -and @($Profile.OutputColumnNames).Count -gt 0
}

function Get-ProfileDataColumnNames {
    param([Parameter(Mandatory = $true)]$Profile)

    $expectedColumns = [int]$Profile.ExpectedColumns
    if ($expectedColumns -lt 1) {
        return @()
    }

    $dataColumnPrefix = if (
        ($Profile.PSObject.Properties.Name -contains 'DataColumnPrefix') -and
        -not [string]::IsNullOrWhiteSpace([string]$Profile.DataColumnPrefix)
    ) {
        [string]$Profile.DataColumnPrefix
    } else {
        'Column'
    }

    $seedNames = @()
    if (Test-ProfileHasCustomOutputColumnNames -Profile $Profile) {
        $seedNames += @($Profile.OutputColumnNames)
    }
    $seedNames += @(1..$expectedColumns | ForEach-Object { '{0}{1}' -f $dataColumnPrefix, $_ })

    return @((Get-UniqueColumnNameList -Names $seedNames) | Select-Object -First $expectedColumns)
}

function Get-ProfileOutputColumnNames {
    param([Parameter(Mandatory = $true)]$Profile)

    $names = @($Profile.SourceFileColumnName)
    $names += @(Get-ProfileDataColumnNames -Profile $Profile)
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

    if ($ExistingReason -like '*ヘッダー位置がずれ*') {
        [void]$categorySet.Add('HEADER_ROW_SHIFT')
    }
    if ($ExistingReason -like '*ヘッダー文言に揺れ*') {
        [void]$categorySet.Add('HEADER_TEXT_DRIFT')
    }
    if ($ExistingReason -like '*空欄で補完*') {
        [void]$categorySet.Add('COLUMN_MISSING')
    }
    if ($ExistingReason -like '*無視しました*') {
        [void]$categorySet.Add('COLUMN_EXTRA')
    }

    $preferredOrder = @('HEADER_MISMATCH', 'HEADER_ROW_SHIFT', 'HEADER_TEXT_DRIFT', 'COLUMN_MISSING', 'COLUMN_EXTRA', 'TIME_MULTI', 'TIME_MISSING', 'TIME_INVALID')
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
    $dataColumnNames = @(Get-ProfileDataColumnNames -Profile $Profile)
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
                SourceColumnName    = if ($sourceColumn -le $dataColumnNames.Count) { [string]$dataColumnNames[$sourceColumn - 1] } else { '{0}{1}' -f $Profile.DataColumnPrefix, $sourceColumn }
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
        '1. run_pdf2excel_v2.bat を実行  2. 対象 PDF を選択  3. 実行前チェックを確認  4. Result / Review / Summary / Errors を確認'
    } else {
        '1. run_pdf2excel_v1.bat を実行  2. PDF を選択  3. 実行前チェックを確認  4. Result / Summary / Errors を確認'
    }

    $checkText = if ($VersionMode -eq 'v2') {
        'Summary は全体件数、Result は生データ転記結果、Review は確認要行、Errors は失敗した PDF と理由です。'
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
    if ($normalizedMessage.Contains('共有パス上')) {
        return [pscustomobject]@{ ErrorCode = 'SHARED_EXECUTION_PATH'; ErrorCategory = 'セキュリティ制約' }
    }
    if ($normalizedMessage.Contains('localappdata')) {
        return [pscustomobject]@{ ErrorCode = 'LOCAL_RUNTIME_UNAVAILABLE'; ErrorCategory = 'セキュリティ制約' }
    }
    if ($normalizedMessage.Contains('実行前チェックでキャンセル')) {
        return [pscustomobject]@{ ErrorCode = 'RUN_CANCELLED'; ErrorCategory = '実行キャンセル' }
    }
    if ($normalizedMessage.Contains('選択されませんでした')) {
        return [pscustomobject]@{ ErrorCode = 'RUN_CANCELLED'; ErrorCategory = '実行キャンセル' }
    }
    if ($normalizedMessage.Contains('同名の pdf')) {
        return [pscustomobject]@{ ErrorCode = 'DUPLICATE_FILE_NAME'; ErrorCategory = '入力エラー' }
    }
    if ($normalizedMessage.Contains('別の pdf2excel 実行が進行中')) {
        return [pscustomobject]@{ ErrorCode = 'RUN_LOCKED'; ErrorCategory = '実行競合' }
    }
    if ($normalizedMessage.Contains('プロファイル')) {
        return [pscustomobject]@{ ErrorCode = 'PROFILE_RESOLUTION_ERROR'; ErrorCategory = '設定エラー' }
    }
    if ($normalizedMessage.Contains('入力フォルダ') -or $normalizedMessage.Contains('入力ファイル')) {
        return [pscustomobject]@{ ErrorCode = 'INPUT_RESOLUTION_ERROR'; ErrorCategory = '入力エラー' }
    }
    if (
        $normalizedMessage.Contains('出力ファイル') -or
        $normalizedMessage.Contains('書き込め') -or
        $normalizedMessage.Contains('saveas')
    ) {
        return [pscustomobject]@{ ErrorCode = 'OUTPUT_WRITE_ERROR'; ErrorCategory = '出力エラー' }
    }
    if ($normalizedMessage.Contains('テンプレート')) {
        return [pscustomobject]@{ ErrorCode = 'TEMPLATE_ERROR'; ErrorCategory = 'テンプレートエラー' }
    }
    if ($normalizedMessage.Contains('excel')) {
        return [pscustomobject]@{ ErrorCode = 'EXCEL_RUNTIME_ERROR'; ErrorCategory = 'Excel実行エラー' }
    }

    return [pscustomobject]@{ ErrorCode = 'UNEXPECTED_RUN_ERROR'; ErrorCategory = 'システムエラー' }
}

function Get-RunErrorGuidance {
    param($ErrorInfo)

    if ($null -eq $ErrorInfo) {
        return 'ログを確認し、入力内容と実行環境を見直してから再実行してください。'
    }

    switch ([string]$ErrorInfo.ErrorCode) {
        'SHARED_EXECUTION_PATH' {
            return 'ZIP をローカルへ展開し、共有フォルダ上ではなくローカルのコピーから再実行してください。'
        }
        'LOCAL_RUNTIME_UNAVAILABLE' {
            return 'LOCALAPPDATA を使える Windows ユーザー環境で再実行してください。端末制約がある場合は管理者へ確認してください。'
        }
        'RUN_CANCELLED' {
            return '入力 PDF、保存先、確認画面の内容を見直してから再実行してください。'
        }
        'DUPLICATE_FILE_NAME' {
            return '同名 PDF が混在しています。ファイル名を変更してから再実行してください。'
        }
        'RUN_LOCKED' {
            return '別の実行が完了するまで待ってから再実行してください。必要ならログフォルダで直近の実行状況を確認してください。'
        }
        'PROFILE_RESOLUTION_ERROR' {
            return '使用するプロファイルを選び直してください。必要ならメニューの「プロファイルを選ぶ」または雛形作成を利用してください。'
        }
        'INPUT_RESOLUTION_ERROR' {
            return '入力フォルダまたは PDF 選択内容を見直してください。PDF が存在し、拡張子が .pdf であることも確認してください。'
        }
        'OUTPUT_WRITE_ERROR' {
            return '保存先フォルダの書き込み権限と既存ファイルの使用状況を確認し、必要なら別の保存先で再実行してください。'
        }
        'TEMPLATE_ERROR' {
            return 'テンプレートの存在と整合性を確認してください。必要ならテンプレート再生成か環境チェックを実施してください。'
        }
        'EXCEL_RUNTIME_ERROR' {
            return 'Excel を閉じて再実行してください。改善しない場合は環境チェックを実施し、Excel COM やテンプレート状態を確認してください。'
        }
        default {
            return 'ログを確認し、改善しない場合は環境チェック結果とあわせて保守担当へ連絡してください。'
        }
    }
}
