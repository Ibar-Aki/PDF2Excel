param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$samplesRoot = Join-Path $projectRoot 'samples'
$pdfRoot = Join-Path $samplesRoot 'pdf'
$sourceRoot = Join-Path $samplesRoot 'source'
$versionedSamplesRoot = Join-Path $samplesRoot 'v1'
$versionedSamplesRootV2 = Join-Path $samplesRoot 'v2'

function Ensure-Directory {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path | Out-Null
    }
}

function Reset-Directory {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Recurse -Force
    }
    Ensure-Directory -Path $Path
}

function Export-WorkbookAsSample {
    param(
        [Parameter(Mandatory = $true)][string]$WorkbookPath,
        [Parameter(Mandatory = $true)][string]$PdfPath,
        [Parameter(Mandatory = $true)][scriptblock]$PopulateWorkbook
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

        & $PopulateWorkbook $worksheet

        $workbook.SaveAs($WorkbookPath, 51)
        $workbook.ExportAsFixedFormat(0, $PdfPath)
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

function Convert-ColumnIndexToLetter {
    param([Parameter(Mandatory = $true)][int]$ColumnIndex)

    $letters = ''
    while ($ColumnIndex -gt 0) {
        $mod = ($ColumnIndex - 1) % 26
        $letters = [string][char](65 + $mod) + $letters
        $ColumnIndex = [Math]::Floor(($ColumnIndex - 1) / 26)
    }

    return $letters
}

function Apply-StandardLayout {
    param(
        [Parameter(Mandatory = $true)]$Worksheet,
        [Parameter(Mandatory = $true)][string]$TitleText,
        [Parameter(Mandatory = $true)][int]$LastColumn,
        [Parameter(Mandatory = $true)][int]$LastRow
    )

    $lastColumnLetter = Convert-ColumnIndexToLetter -ColumnIndex $LastColumn

    $titleRange = $Worksheet.Range("A1:${lastColumnLetter}1")
    $titleRange.Merge()
    $titleRange.Value2 = $TitleText
    $titleRange.HorizontalAlignment = -4108
    $titleRange.Font.Bold = $true
    $titleRange.Font.Size = 14

    $headerRange = $Worksheet.Range("A2:${lastColumnLetter}2")
    $headerRange.Font.Bold = $true
    $headerRange.Interior.Color = 15921906

    $usedRange = $Worksheet.Range("A2:${lastColumnLetter}${LastRow}")
    $usedRange.Borders.LineStyle = 1

    $Worksheet.Columns.AutoFit() | Out-Null
    $Worksheet.PageSetup.Orientation = 2
    $Worksheet.PageSetup.Zoom = $false
    $Worksheet.PageSetup.FitToPagesWide = 1
    $Worksheet.PageSetup.FitToPagesTall = 1
}

function New-JapaneseAttendanceWorkbookAndPdf {
    param(
        [Parameter(Mandatory = $true)][string]$WorkbookPath,
        [Parameter(Mandatory = $true)][string]$PdfPath,
        [Parameter(Mandatory = $true)][string]$MonthLabel
    )

    $memberRows = @(
        @('A001', '佐藤花子', '営業部', '通常'),
        @('A002', '鈴木一郎', '営業部', '在宅'),
        @('A003', '田中美咲', '管理部', '通常'),
        @('A004', '高橋健太', '管理部', '通常'),
        @('A005', '伊藤直子', '開発部', '在宅'),
        @('A006', '渡辺大輔', '開発部', '通常')
    )
    $statusCycle = @('出勤', '在宅', '休暇', '半休', '遅刻', '出勤', '出勤', '在宅')

    Export-WorkbookAsSample -WorkbookPath $WorkbookPath -PdfPath $PdfPath -PopulateWorkbook {
        param($worksheet)

        $worksheet.Name = '勤怠表'
        $headers = @('社員番号', '氏名', '所属', '勤務区分')
        $headers += @(1..30 | ForEach-Object { '{0}日' -f $_ })
        $headers += @('備考')

        for ($column = 1; $column -le $headers.Count; $column += 1) {
            $worksheet.Cells.Item(2, $column).Value2 = $headers[$column - 1]
        }

        for ($rowIndex = 0; $rowIndex -lt $memberRows.Count; $rowIndex += 1) {
            $excelRow = $rowIndex + 3
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

        Apply-StandardLayout -Worksheet $worksheet -TitleText "$MonthLabel 勤怠管理表" -LastColumn 35 -LastRow 8
    }
}

function New-JapaneseSalesWorkbookAndPdf {
    param(
        [Parameter(Mandatory = $true)][string]$WorkbookPath,
        [Parameter(Mandatory = $true)][string]$PdfPath,
        [Parameter(Mandatory = $true)][string]$ReportLabel,
        [Parameter(Mandatory = $true)][string]$StoreName
    )

    $rows = @(
        @($ReportLabel, $StoreName, '飲料', 'ミネラルウォーター', '24', '2880', '2160', '720', '山田芽衣', '現金', '販促展開あり', '確定'),
        @($ReportLabel, $StoreName, '文具', 'ノートA5', '15', '2250', '1500', '750', '山田芽衣', '現金', '新学期需要', '確定'),
        @($ReportLabel, $StoreName, '食品', 'カレーセット', '12', '3960', '2820', '1140', '高田修平', 'クレジット', 'まとめ買い', '確定'),
        @($ReportLabel, $StoreName, '雑貨', '折りたたみ傘', '8', '6400', '4240', '2160', '高田修平', '電子マネー', '雨天対応', '確定'),
        @($ReportLabel, $StoreName, '家電', 'USB充電器', '6', '5340', '3720', '1620', '山田芽衣', 'クレジット', '週末販促', '確定')
    )

    Export-WorkbookAsSample -WorkbookPath $WorkbookPath -PdfPath $PdfPath -PopulateWorkbook {
        param($worksheet)

        $worksheet.Name = '売上日報'
        $headers = @('売上日', '店舗名', '部門', '商品名', '数量', '売上金額', '原価', '粗利', '担当者', '支払区分', '備考', '確定状態')
        for ($column = 1; $column -le $headers.Count; $column += 1) {
            $worksheet.Cells.Item(2, $column).Value2 = $headers[$column - 1]
        }
        for ($rowIndex = 0; $rowIndex -lt $rows.Count; $rowIndex += 1) {
            $excelRow = $rowIndex + 3
            for ($column = 1; $column -le $headers.Count; $column += 1) {
                $worksheet.Cells.Item($excelRow, $column).Value2 = $rows[$rowIndex][$column - 1]
            }
        }

        Apply-StandardLayout -Worksheet $worksheet -TitleText "$ReportLabel 売上日報 ($StoreName)" -LastColumn 12 -LastRow 7
    }
}

function New-JapaneseInventoryWorkbookAndPdf {
    param(
        [Parameter(Mandatory = $true)][string]$WorkbookPath,
        [Parameter(Mandatory = $true)][string]$PdfPath,
        [Parameter(Mandatory = $true)][string]$SeasonLabel,
        [Parameter(Mandatory = $true)][string]$WarehouseName
    )

    $rows = @(
        @('P-1001', 'コピー用紙A4', '文具', $WarehouseName, '180', '60', '480', 'A-01', "$SeasonLabel 棚卸済み", '2026/03/10'),
        @('P-1002', 'ボールペン黒', '文具', $WarehouseName, '320', '120', '95', 'A-03', "$SeasonLabel 棚卸済み", '2026/03/10'),
        @('P-2001', 'キーボード', 'IT機器', $WarehouseName, '42', '15', '2980', 'B-11', "$SeasonLabel 棚卸済み", '2026/03/10'),
        @('P-2002', 'マウス', 'IT機器', $WarehouseName, '65', '20', '1480', 'B-12', "$SeasonLabel 棚卸済み", '2026/03/10'),
        @('P-3001', '折りたたみ椅子', '備品', $WarehouseName, '24', '8', '3980', 'C-02', "$SeasonLabel 棚卸済み", '2026/03/10'),
        @('P-3002', '延長コード', '備品', $WarehouseName, '58', '18', '1280', 'C-08', "$SeasonLabel 棚卸済み", '2026/03/10')
    )

    Export-WorkbookAsSample -WorkbookPath $WorkbookPath -PdfPath $PdfPath -PopulateWorkbook {
        param($worksheet)

        $worksheet.Name = '在庫一覧'
        $headers = @('品番', '品名', '区分', '倉庫', '在庫数', '発注点', '単価', '棚番', '備考', '更新日')
        for ($column = 1; $column -le $headers.Count; $column += 1) {
            $worksheet.Cells.Item(2, $column).Value2 = $headers[$column - 1]
        }
        for ($rowIndex = 0; $rowIndex -lt $rows.Count; $rowIndex += 1) {
            $excelRow = $rowIndex + 3
            for ($column = 1; $column -le $headers.Count; $column += 1) {
                $worksheet.Cells.Item($excelRow, $column).Value2 = $rows[$rowIndex][$column - 1]
            }
        }

        Apply-StandardLayout -Worksheet $worksheet -TitleText "$SeasonLabel 在庫一覧 ($WarehouseName)" -LastColumn 10 -LastRow 8
    }
}

function New-JapaneseInquiryWorkbookAndPdf {
    param(
        [Parameter(Mandatory = $true)][string]$WorkbookPath,
        [Parameter(Mandatory = $true)][string]$PdfPath,
        [Parameter(Mandatory = $true)][string]$WeekLabel
    )

    $rows = @(
        @("$WeekLabel 1日", 'Q-001', '株式会社青葉', 'メール', '請求書再発行', '石井彩乃', '対応中', "$WeekLabel 3日", '経理確認待ち'),
        @("$WeekLabel 2日", 'Q-002', '合同会社みなと', '電話', '納期確認', '川口優', '完了', "$WeekLabel 2日", '当日回答'),
        @("$WeekLabel 3日", 'Q-003', '山城商事', 'フォーム', '返品相談', '石井彩乃', '対応中', "$WeekLabel 5日", '写真受領済み'),
        @("$WeekLabel 4日", 'Q-004', '西東京サービス', 'メール', '仕様問い合わせ', '中島直人', '保留', "$WeekLabel 8日", '開発確認中'),
        @("$WeekLabel 5日", 'Q-005', '北辰物流', '電話', '操作方法', '中島直人', '完了', "$WeekLabel 5日", 'マニュアル送付済み')
    )

    Export-WorkbookAsSample -WorkbookPath $WorkbookPath -PdfPath $PdfPath -PopulateWorkbook {
        param($worksheet)

        $worksheet.Name = '問い合わせ'
        $headers = @('受付日', '問い合わせID', '顧客名', 'チャネル', '件名', '担当者', '状態', '回答期限', '備考')
        for ($column = 1; $column -le $headers.Count; $column += 1) {
            $worksheet.Cells.Item(2, $column).Value2 = $headers[$column - 1]
        }
        for ($rowIndex = 0; $rowIndex -lt $rows.Count; $rowIndex += 1) {
            $excelRow = $rowIndex + 3
            for ($column = 1; $column -le $headers.Count; $column += 1) {
                $worksheet.Cells.Item($excelRow, $column).Value2 = $rows[$rowIndex][$column - 1]
            }
        }

        Apply-StandardLayout -Worksheet $worksheet -TitleText "$WeekLabel 問い合わせ管理表" -LastColumn 9 -LastRow 7
    }
}

function New-ConstructionTransferPocWorkbookAndPdf {
    param(
        [Parameter(Mandatory = $true)][string]$WorkbookPath,
        [Parameter(Mandatory = $true)][string]$PdfPath
    )

    $headers = @(
        '日付', '曜日', '氏名', '所属', '現場1', '入場1', '退場1', '現場2', '入場2', '退場2',
        '現場3', '入場3', '退場3', '休憩', '当日小計', '事務所入', '事務所出', '移動', '宿泊', '資格',
        '工種', '天候', '体調', '安全確認', '備考1', '備考2', '確認者', '管理者メモ', '予備1', '予備2'
    )

    $rows = @(
        @("2/4`n（金）", '金', '佐藤 花子', '一次協力', "東京駅前再開発A棟`n東工区", ' 10:30', '19：00', '仮設事務所', '9時00分', '9:20', '', '', '', '60', '7:30', '8:30', '20:00', 'あり', 'なし', '職長教育', '鉄筋', '晴', '良好', '済', "入場時に`n安全帯確認", '', '山田健', 'A棟は午後からB1へ応援', '', ''),
        @("2/5`n（土）", '土', '鈴木一郎', '一次協力', "東京駅前再開発Ａ棟", '10時30分', '18時 00分', "南口歩道橋更新`nその2", '18:25', '20：10', '', '', '', '45', '8:10', '9:10', '20:20', 'あり', 'なし', '高所作業', '足場', '曇', '良好', '済', '', "現場名が微妙に揺れ", '山田健', '', '', ''),
        @("2/6`n（日）", '日', '田中美咲', '直用', "湾岸物流センター`n新築工事", '9：05', '17:30', '', '', '', '', '', '', '60', '7:25', '8:10', '18:10', 'あり', 'なし', '玉掛', '搬入', '晴', '良好', '済', '朝礼あり', '', '岡本進', "氏名表記ゆれなし", '', ''),
        @("2/7`n（月）", '月', '高橋健太', '直用', "湾岸物流センタ-新築工事", ' 9時 15分', '17時30分', "東京駅前再開発A棟", '18:00', '20:15', '', '', '', '60', '9:30', '8:50', '20:40', 'あり', 'なし', '職長教育', '鉄骨', '雨', '普通', '済', "現場名の記号ゆれ", '', '岡本進', '', '', ''),
        @("2/8`n（火）", '火', '渡辺大輔', '二次協力', "渋谷駅西口改良工事`n南工区", '10：3０', '19:05', '', '', '', '', '', '', '60', '7:35', '9:00', '19:30', 'なし', 'なし', '誘導員', '雑工', '晴', '良好', '済', "全角数字混在`n確認要", '', '山田健', '', '', '')
    )

    Export-WorkbookAsSample -WorkbookPath $WorkbookPath -PdfPath $PdfPath -PopulateWorkbook {
        param($worksheet)

        $worksheet.Name = '勤怠一覧PoC'
        $worksheet.Range('A1:AD1').Merge()
        $worksheet.Range('A1').Value2 = '建設現場別 延べ作業時間算出用 勤怠一覧（PoC）'
        $worksheet.Range('A2').Value2 = '管理者: 山田健'
        $worksheet.Range('A3').Value2 = '対象月: 2026年02月'

        for ($column = 1; $column -le $headers.Count; $column += 1) {
            $worksheet.Cells.Item(4, $column).Value2 = $headers[$column - 1]
        }

        for ($rowIndex = 0; $rowIndex -lt $rows.Count; $rowIndex += 1) {
            $excelRow = $rowIndex + 5
            for ($column = 1; $column -le $headers.Count; $column += 1) {
                $worksheet.Cells.Item($excelRow, $column).Value2 = $rows[$rowIndex][$column - 1]
            }
        }

        $worksheet.Range('A1:AD1').HorizontalAlignment = -4108
        $worksheet.Range('A1:AD1').Font.Bold = $true
        $worksheet.Range('A1:AD1').Font.Size = 14
        $worksheet.Range('A2:A3').Font.Bold = $true
        $worksheet.Range('A4:AD4').Font.Bold = $true
        $worksheet.Range('A4:AD4').Interior.Color = 15921906
        $worksheet.Range('A4:AD9').Borders.LineStyle = 1
        $worksheet.Range('A4:AD9').WrapText = $true

        foreach ($rowNumber in 5..9) {
            $worksheet.Rows.Item($rowNumber).RowHeight = 42
        }

        $worksheet.PageSetup.PaperSize = 8
        $worksheet.PageSetup.Orientation = 2
        $worksheet.PageSetup.Zoom = $false
        $worksheet.PageSetup.FitToPagesWide = 2
        $worksheet.PageSetup.FitToPagesTall = 1
        $worksheet.Columns.AutoFit() | Out-Null
    }
}

function New-ConstructionTransferPagedPocWorkbookAndPdf {
    param(
        [Parameter(Mandatory = $true)][string]$WorkbookPath,
        [Parameter(Mandatory = $true)][string]$PdfPath,
        [Parameter(Mandatory = $true)][int]$PageCount,
        [Parameter(Mandatory = $true)][string]$TitleLabel,
        [Parameter(Mandatory = $true)][string[]]$MonthLabels
    )

    $headers = @(
        '日付', '曜日', '氏名', '所属', '現場1', '入場1', '退場1', '現場2', '入場2', '退場2',
        '現場3', '入場3', '退場3', '休憩', '当日小計', '事務所入', '事務所出', '移動', '宿泊', '資格',
        '工種', '天候', '体調', '安全確認', '備考1', '備考2', '確認者', '管理者メモ', '予備1', '予備2'
    )

    $baseRows = @(
        @("2/4`n（金）", '金', '佐藤 花子', '一次協力', "東京駅前再開発A棟`n東工区", ' 10:30', '19：00', '仮設事務所', '9時00分', '9:20', '', '', '', '60', '7:30', '8:30', '20:00', 'あり', 'なし', '職長教育', '鉄筋', '晴', '良好', '済', "入場時に`n安全帯確認", '', '山田健', 'A棟は午後からB1へ応援', '', ''),
        @("2/5`n（土）", '土', '鈴木一郎', '一次協力', "東京駅前再開発Ａ棟", '10時30分', '18時 00分', "南口歩道橋更新`nその2", '18:25', '20：10', '', '', '', '45', '8:10', '9:10', '20:20', 'あり', 'なし', '高所作業', '足場', '曇', '良好', '済', '', "現場名が微妙に揺れ", '山田健', '', '', ''),
        @("2/6`n（日）", '日', '田中美咲', '直用', "湾岸物流センター`n新築工事", '9：05', '17:30', '', '', '', '', '', '', '60', '7:25', '8:10', '18:10', 'あり', 'なし', '玉掛', '搬入', '晴', '良好', '済', '朝礼あり', '', '岡本進', "氏名表記ゆれなし", '', ''),
        @("2/7`n（月）", '月', '高橋健太', '直用', "湾岸物流センタ-新築工事", ' 9時 15分', '17時30分', "東京駅前再開発A棟", '18:00', '20:15', '', '', '', '60', '9:30', '8:50', '20:40', 'あり', 'なし', '職長教育', '鉄骨', '雨', '普通', '済', "現場名の記号ゆれ", '', '岡本進', '', '', ''),
        @("2/8`n（火）", '火', '渡辺大輔', '二次協力', "渋谷駅西口改良工事`n南工区", '10：3０', '19:05', '', '', '', '', '', '', '60', '7:35', '9:00', '19:30', 'なし', 'なし', '誘導員', '雑工', '晴', '良好', '済', "全角数字混在`n確認要", '', '山田健', '', '', '')
    )

    Export-WorkbookAsSample -WorkbookPath $WorkbookPath -PdfPath $PdfPath -PopulateWorkbook {
        param($worksheet)

        $worksheet.Name = '勤怠一覧MultiPoC'

        $worksheet.Range('A1:AD1').Merge()
        $worksheet.Range('A1').Value2 = $TitleLabel
        $worksheet.Range('A2').Value2 = '管理者: 山田健'
        $worksheet.Range('A3').Value2 = ('対象月: ' + ($MonthLabels -join ' / '))

        $worksheet.Range('A1:AD1').HorizontalAlignment = -4108
        $worksheet.Range('A1:AD1').Font.Bold = $true
        $worksheet.Range('A1:AD1').Font.Size = 14
        $worksheet.Range('A2:A3').Font.Bold = $true

        $currentRow = 4
        for ($page = 1; $page -le $PageCount; $page += 1) {
            if ($page -gt 1) {
                $worksheet.Cells.Item($currentRow, 1).Value2 = "ページ$page 補助欄"
                $worksheet.Rows.Item($currentRow).RowHeight = 16
                $currentRow += 1
            }

            for ($column = 1; $column -le $headers.Count; $column += 1) {
                $worksheet.Cells.Item($currentRow, $column).Value2 = $headers[$column - 1]
            }

            $headerRow = $currentRow
            $worksheet.Range("A${headerRow}:AD${headerRow}").Font.Bold = $true
            $worksheet.Range("A${headerRow}:AD${headerRow}").Interior.Color = 15921906
            $currentRow += 1

            for ($rowIndex = 0; $rowIndex -lt $baseRows.Count; $rowIndex += 1) {
                $excelRow = $currentRow + $rowIndex
                $rowValues = @($baseRows[$rowIndex])
                $monthLabel = $MonthLabels[[Math]::Min($page - 1, $MonthLabels.Count - 1)]
                $daySeed = (($page - 1) * 5) + $rowIndex + 1
                $rowValues[0] = "{0}`n（{1}）" -f ("{0}/{1}" -f ($page + 1), $daySeed), @('月', '火', '水', '木', '金', '土', '日')[($page + $rowIndex) % 7]
                $rowValues[4] = if ($rowIndex % 2 -eq 0) { "$monthLabel`n$rowValues[4]" } else { $rowValues[4] }
                $rowValues[27] = "ページ$page / $monthLabel / " + $rowValues[27]
                for ($column = 1; $column -le $headers.Count; $column += 1) {
                    $worksheet.Cells.Item($excelRow, $column).Value2 = $rowValues[$column - 1]
                }
                $worksheet.Rows.Item($excelRow).RowHeight = 42
            }

            $lastDataRow = $currentRow + $baseRows.Count - 1
            $worksheet.Range("A${headerRow}:AD${lastDataRow}").Borders.LineStyle = 1
            $worksheet.Range("A${headerRow}:AD${lastDataRow}").WrapText = $true
            $currentRow = $lastDataRow + 1

            if ($page -lt $PageCount) {
                [void]$worksheet.HPageBreaks.Add($worksheet.Cells.Item($currentRow, 1))
            }
        }

        $worksheet.PageSetup.PaperSize = 8
        $worksheet.PageSetup.Orientation = 2
        $worksheet.PageSetup.Zoom = $false
        $worksheet.PageSetup.FitToPagesWide = 1
        $worksheet.PageSetup.FitToPagesTall = $false
        $worksheet.Columns.AutoFit() | Out-Null
    }
}

$managedDirectories = @(
    'attendance_jp',
    'sales_daily_jp',
    'inventory_jp',
    'inquiry_jp',
    'construction_transfer_poc'
)

Ensure-Directory -Path $pdfRoot
Ensure-Directory -Path $sourceRoot
Ensure-Directory -Path (Join-Path $versionedSamplesRoot 'pdf')
Ensure-Directory -Path (Join-Path $versionedSamplesRoot 'source')
Ensure-Directory -Path (Join-Path $versionedSamplesRootV2 'pdf')
Ensure-Directory -Path (Join-Path $versionedSamplesRootV2 'source')

foreach ($directoryName in $managedDirectories) {
    Reset-Directory -Path (Join-Path $pdfRoot $directoryName)
    Reset-Directory -Path (Join-Path $sourceRoot $directoryName)
}

$sampleDefinitions = @(
    [pscustomobject]@{
        PdfDirectory      = 'attendance_jp'
        SourceDirectory   = 'attendance_jp'
        WorkbookName      = '2026年03月_勤怠管理表.xlsx'
        PdfName           = '2026年03月_勤怠管理表.pdf'
        CreateSample      = { param($workbookPath, $pdfPath) New-JapaneseAttendanceWorkbookAndPdf -WorkbookPath $workbookPath -PdfPath $pdfPath -MonthLabel '2026年03月' }
    },
    [pscustomobject]@{
        PdfDirectory      = 'attendance_jp'
        SourceDirectory   = 'attendance_jp'
        WorkbookName      = '2026年04月_勤怠管理表.xlsx'
        PdfName           = '2026年04月_勤怠管理表.pdf'
        CreateSample      = { param($workbookPath, $pdfPath) New-JapaneseAttendanceWorkbookAndPdf -WorkbookPath $workbookPath -PdfPath $pdfPath -MonthLabel '2026年04月' }
    },
    [pscustomobject]@{
        PdfDirectory      = 'sales_daily_jp'
        SourceDirectory   = 'sales_daily_jp'
        WorkbookName      = '2026-03-15_売上日報_東京店.xlsx'
        PdfName           = '2026-03-15_売上日報_東京店.pdf'
        CreateSample      = { param($workbookPath, $pdfPath) New-JapaneseSalesWorkbookAndPdf -WorkbookPath $workbookPath -PdfPath $pdfPath -ReportLabel '2026-03-15' -StoreName '東京店' }
    },
    [pscustomobject]@{
        PdfDirectory      = 'sales_daily_jp'
        SourceDirectory   = 'sales_daily_jp'
        WorkbookName      = '2026-03-16_売上日報_横浜店.xlsx'
        PdfName           = '2026-03-16_売上日報_横浜店.pdf'
        CreateSample      = { param($workbookPath, $pdfPath) New-JapaneseSalesWorkbookAndPdf -WorkbookPath $workbookPath -PdfPath $pdfPath -ReportLabel '2026-03-16' -StoreName '横浜店' }
    },
    [pscustomobject]@{
        PdfDirectory      = 'inventory_jp'
        SourceDirectory   = 'inventory_jp'
        WorkbookName      = '春季_在庫一覧_倉庫A.xlsx'
        PdfName           = '春季_在庫一覧_倉庫A.pdf'
        CreateSample      = { param($workbookPath, $pdfPath) New-JapaneseInventoryWorkbookAndPdf -WorkbookPath $workbookPath -PdfPath $pdfPath -SeasonLabel '春季' -WarehouseName '倉庫A' }
    },
    [pscustomobject]@{
        PdfDirectory      = 'inventory_jp'
        SourceDirectory   = 'inventory_jp'
        WorkbookName      = '春季_在庫一覧_倉庫B.xlsx'
        PdfName           = '春季_在庫一覧_倉庫B.pdf'
        CreateSample      = { param($workbookPath, $pdfPath) New-JapaneseInventoryWorkbookAndPdf -WorkbookPath $workbookPath -PdfPath $pdfPath -SeasonLabel '春季' -WarehouseName '倉庫B' }
    },
    [pscustomobject]@{
        PdfDirectory      = 'inquiry_jp'
        SourceDirectory   = 'inquiry_jp'
        WorkbookName      = '2026年03月_問い合わせ管理表_第1週.xlsx'
        PdfName           = '2026年03月_問い合わせ管理表_第1週.pdf'
        CreateSample      = { param($workbookPath, $pdfPath) New-JapaneseInquiryWorkbookAndPdf -WorkbookPath $workbookPath -PdfPath $pdfPath -WeekLabel '2026年03月 第1週' }
    },
    [pscustomobject]@{
        PdfDirectory      = 'inquiry_jp'
        SourceDirectory   = 'inquiry_jp'
        WorkbookName      = '2026年03月_問い合わせ管理表_第2週.xlsx'
        PdfName           = '2026年03月_問い合わせ管理表_第2週.pdf'
        CreateSample      = { param($workbookPath, $pdfPath) New-JapaneseInquiryWorkbookAndPdf -WorkbookPath $workbookPath -PdfPath $pdfPath -WeekLabel '2026年03月 第2週' }
    },
    [pscustomobject]@{
        PdfDirectory      = 'construction_transfer_poc'
        SourceDirectory   = 'construction_transfer_poc'
        WorkbookName      = '2026年02月_作業員勤怠一覧_PoC.xlsx'
        PdfName           = '2026年02月_作業員勤怠一覧_PoC.pdf'
        CreateSample      = { param($workbookPath, $pdfPath) New-ConstructionTransferPocWorkbookAndPdf -WorkbookPath $workbookPath -PdfPath $pdfPath }
    },
    [pscustomobject]@{
        PdfDirectory      = 'construction_transfer_poc'
        SourceDirectory   = 'construction_transfer_poc'
        WorkbookName      = '2026年02月_作業員勤怠一覧_PoC_2ページ同一列.xlsx'
        PdfName           = '2026年02月_作業員勤怠一覧_PoC_2ページ同一列.pdf'
        CreateSample      = { param($workbookPath, $pdfPath) New-ConstructionTransferPagedPocWorkbookAndPdf -WorkbookPath $workbookPath -PdfPath $pdfPath -PageCount 2 -TitleLabel '建設現場別 延べ作業時間算出用 勤怠一覧（2ページ同一列PoC）' -MonthLabels @('2026年02月', '2026年02月') }
    },
    [pscustomobject]@{
        PdfDirectory      = 'construction_transfer_poc'
        SourceDirectory   = 'construction_transfer_poc'
        WorkbookName      = '2026年04月-06月_作業員勤怠一覧_PoC_6ページ同一列.xlsx'
        PdfName           = '2026年04月-06月_作業員勤怠一覧_PoC_6ページ同一列.pdf'
        CreateSample      = { param($workbookPath, $pdfPath) New-ConstructionTransferPagedPocWorkbookAndPdf -WorkbookPath $workbookPath -PdfPath $pdfPath -PageCount 6 -TitleLabel '建設現場別 延べ作業時間算出用 勤怠一覧（6ページ同一列PoC）' -MonthLabels @('2026年04月', '2026年04月', '2026年05月', '2026年05月', '2026年06月', '2026年06月') }
    }
)

foreach ($definition in $sampleDefinitions) {
    $pdfDirectory = Join-Path $pdfRoot $definition.PdfDirectory
    $sourceDirectory = Join-Path $sourceRoot $definition.SourceDirectory
    Ensure-Directory -Path $pdfDirectory
    Ensure-Directory -Path $sourceDirectory

    $workbookPath = Join-Path $sourceDirectory $definition.WorkbookName
    $pdfPath = Join-Path $pdfDirectory $definition.PdfName
    & $definition.CreateSample $workbookPath $pdfPath
}

$v1SampleDirectories = @('attendance_jp', 'sales_daily_jp', 'inventory_jp', 'inquiry_jp')
$v2SampleDirectories = @('construction_transfer_poc')

foreach ($directoryName in $v1SampleDirectories) {
    $targetPdfDir = Join-Path (Join-Path $versionedSamplesRoot 'pdf') $directoryName
    $targetSourceDir = Join-Path (Join-Path $versionedSamplesRoot 'source') $directoryName
    Reset-Directory -Path $targetPdfDir
    Reset-Directory -Path $targetSourceDir
    Copy-Item -Path (Join-Path (Join-Path $pdfRoot $directoryName) '*') -Destination $targetPdfDir -Recurse -Force
    Copy-Item -Path (Join-Path (Join-Path $sourceRoot $directoryName) '*') -Destination $targetSourceDir -Recurse -Force
}

foreach ($directoryName in $v2SampleDirectories) {
    $targetPdfDir = Join-Path (Join-Path $versionedSamplesRootV2 'pdf') $directoryName
    $targetSourceDir = Join-Path (Join-Path $versionedSamplesRootV2 'source') $directoryName
    Reset-Directory -Path $targetPdfDir
    Reset-Directory -Path $targetSourceDir
    Copy-Item -Path (Join-Path (Join-Path $pdfRoot $directoryName) '*') -Destination $targetPdfDir -Recurse -Force
    Copy-Item -Path (Join-Path (Join-Path $sourceRoot $directoryName) '*') -Destination $targetSourceDir -Recurse -Force
}

Write-Host '日本語サンプル帳票を作成しました。'
Write-Host "PDF: $pdfRoot"
Write-Host "元Excel: $sourceRoot"
Write-Host "V1 PDF: $(Join-Path $versionedSamplesRoot 'pdf')"
Write-Host "V2 PDF: $(Join-Path $versionedSamplesRootV2 'pdf')"
