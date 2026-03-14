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

function Get-ProfileOutputColumnNames {
    param([Parameter(Mandatory = $true)]$Profile)

    $names = @($Profile.SourceFileColumnName)
    foreach ($index in 1..$Profile.ExpectedColumns) {
        $names += '{0}{1}' -f $Profile.DataColumnPrefix, $index
    }

    return $names
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
