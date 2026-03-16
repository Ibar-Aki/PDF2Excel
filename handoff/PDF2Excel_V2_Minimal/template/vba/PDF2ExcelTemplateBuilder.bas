Attribute VB_Name = "PDF2ExcelTemplateBuilder"
Option Explicit

Private Const CONTROL_SHEET As String = "Control"
Private Const RESULT_SHEET As String = "Result"
Private Const ERRORS_SHEET As String = "Errors"
Private Const SUMMARY_SHEET As String = "Summary"
Private Const REVIEW_SHEET As String = "Review"
Private Const TEMPLATE_VARIANT_V1 As String = "v1"
Private Const TEMPLATE_VARIANT_V2 As String = "v2"

Public Sub BuildPDF2ExcelTemplateInActiveWorkbook()
    BuildTemplateInActiveWorkbook TEMPLATE_VARIANT_V1
End Sub

Public Sub BuildPDF2ExcelV1TemplateInActiveWorkbook()
    BuildTemplateInActiveWorkbook TEMPLATE_VARIANT_V1
End Sub

Public Sub BuildPDF2ExcelV2TemplateInActiveWorkbook()
    BuildTemplateInActiveWorkbook TEMPLATE_VARIANT_V2
End Sub

Private Sub BuildTemplateInActiveWorkbook(ByVal templateVariant As String)
    On Error GoTo EH

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    EnsureWorkbookHasRequiredSheets ThisWorkbook, templateVariant

    SetControlSheetLayout ThisWorkbook.Worksheets.Item(1), templateVariant
    SetDataSheetLayout ThisWorkbook.Worksheets.Item(2), RESULT_SHEET, _
        "変換成功データがここに読み込まれます。A列はPDFファイル名、B列以降はプロファイルに応じた表データです。"
    SetDataSheetLayout ThisWorkbook.Worksheets.Item(3), ERRORS_SHEET, _
        "失敗したPDFと、分類されたエラー理由がここに一覧表示されます。"
    SetSummarySheetLayout ThisWorkbook.Worksheets.Item(4), templateVariant

    If IsV2Variant(templateVariant) Then
        SetReviewSheetLayout ThisWorkbook.Worksheets.Item(5)
    End If

    ThisWorkbook.Worksheets(CONTROL_SHEET).Activate
    SetStatus "待機中"

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
EH:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Err.Raise Err.Number, "BuildTemplateInActiveWorkbook", Err.Description
End Sub

Private Function ControlSheet() As Worksheet
    Set ControlSheet = ThisWorkbook.Worksheets(CONTROL_SHEET)
End Function

Private Sub SetStatus(ByVal statusText As String)
    ControlSheet.Range("B6").Value = statusText
    ControlSheet.Range("B5").Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Sub

Private Function IsV2Variant(ByVal templateVariant As String) As Boolean
    IsV2Variant = (LCase$(Trim$(templateVariant)) = TEMPLATE_VARIANT_V2)
End Function

Private Function GetRequiredSheetCount(ByVal templateVariant As String) As Long
    If IsV2Variant(templateVariant) Then
        GetRequiredSheetCount = 5
    Else
        GetRequiredSheetCount = 4
    End If
End Function

Private Function GetVersionDisplayName(ByVal templateVariant As String) As String
    If IsV2Variant(templateVariant) Then
        GetVersionDisplayName = "VER2"
    Else
        GetVersionDisplayName = "VER1"
    End If
End Function

Private Sub EnsureWorkbookHasRequiredSheets(ByVal targetWorkbook As Workbook, ByVal templateVariant As String)
    Dim requiredSheetCount As Long

    requiredSheetCount = GetRequiredSheetCount(templateVariant)

    Do While targetWorkbook.Worksheets.Count < requiredSheetCount
        targetWorkbook.Worksheets.Add After:=targetWorkbook.Worksheets(targetWorkbook.Worksheets.Count)
    Loop

    Do While targetWorkbook.Worksheets.Count > requiredSheetCount
        targetWorkbook.Worksheets.Item(targetWorkbook.Worksheets.Count).Delete
    Loop
End Sub

Private Sub SetControlSheetLayout(ByVal targetSheet As Worksheet, ByVal templateVariant As String)
    targetSheet.Cells.Clear
    targetSheet.Name = CONTROL_SHEET

    SetControlSheetStaticCells targetSheet, templateVariant

    targetSheet.Range("A1:B1").Font.Bold = True
    targetSheet.Range("A1:B1").Interior.Color = 15773696
    targetSheet.Range("A15:A18").Font.Bold = True
    targetSheet.Columns.Item("A").ColumnWidth = 18
    targetSheet.Columns.Item("B").ColumnWidth = 92
    targetSheet.Range("A1:B18").VerticalAlignment = -4160
    targetSheet.Range("A1:B18").WrapText = True
    targetSheet.Range("B5").Value = ""
    targetSheet.Range("B6").Value = "待機中"
    targetSheet.Range("B7:B14").Value = ""
    targetSheet.Range("B15").Value = GetVersionDisplayName(templateVariant)
    targetSheet.Activate
    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True
End Sub

Private Sub SetDataSheetLayout(ByVal targetSheet As Worksheet, ByVal sheetName As String, ByVal description As String)
    targetSheet.Cells.Clear
    targetSheet.Name = sheetName
    targetSheet.Range("A1").Value = sheetName
    targetSheet.Range("A2").Value = description
    targetSheet.Range("A1").Font.Bold = True
    targetSheet.Range("A1").Font.Size = 14
    targetSheet.Columns.Item("A").ColumnWidth = 24
    targetSheet.Columns.Item("B").ColumnWidth = 24
End Sub

Private Sub SetSummarySheetLayout(ByVal targetSheet As Worksheet, ByVal templateVariant As String)
    targetSheet.Cells.Clear
    targetSheet.Name = SUMMARY_SHEET

    SetSummarySheetStaticCells targetSheet, templateVariant

    targetSheet.Range("A1").Font.Bold = True
    targetSheet.Range("A1").Font.Size = 14
    targetSheet.Range("A4:B4").Font.Bold = True
    targetSheet.Range("A4:B4").Interior.Color = 15773696
    targetSheet.Range("A13").Font.Bold = True
    targetSheet.Range("M13").Font.Bold = True
    targetSheet.Columns.Item("A").ColumnWidth = 24
    targetSheet.Columns.Item("B").ColumnWidth = 28
    targetSheet.Columns.Item("C").ColumnWidth = 22
    targetSheet.Columns.Item("D").ColumnWidth = 18
    targetSheet.Columns.Item("E").ColumnWidth = 18
    targetSheet.Columns.Item("F").ColumnWidth = 20
    targetSheet.Columns.Item("G").ColumnWidth = 16
    targetSheet.Columns.Item("H").ColumnWidth = 16
    targetSheet.Columns.Item("I").ColumnWidth = 18
    targetSheet.Columns.Item("J").ColumnWidth = 18
    targetSheet.Columns.Item("K").ColumnWidth = 18
    targetSheet.Columns.Item("L").ColumnWidth = 16
    targetSheet.Columns.Item("M").ColumnWidth = 18
    targetSheet.Columns.Item("N").ColumnWidth = 18
    targetSheet.Columns.Item("O").ColumnWidth = 16
End Sub

Private Sub SetReviewSheetLayout(ByVal targetSheet As Worksheet)
    targetSheet.Cells.Clear
    targetSheet.Name = REVIEW_SHEET
    targetSheet.Range("A1").Value = REVIEW_SHEET
    targetSheet.Range("A2").Value = "確認が必要な生データ転記行を一覧化します。"
    targetSheet.Range("A4").Value = "元ファイル名"
    targetSheet.Range("B4").Value = "ページ"
    targetSheet.Range("C4").Value = "氏名 raw"
    targetSheet.Range("D4").Value = "現場 raw"
    targetSheet.Range("E4").Value = "入場 raw"
    targetSheet.Range("F4").Value = "退場 raw"
    targetSheet.Range("G4").Value = "確認要理由"
    targetSheet.Range("A1").Font.Bold = True
    targetSheet.Range("A1").Font.Size = 14
    targetSheet.Range("A4:G4").Font.Bold = True
    targetSheet.Range("A4:G4").Interior.Color = 15773696
    targetSheet.Columns.Item("A").ColumnWidth = 28
    targetSheet.Columns.Item("B").ColumnWidth = 10
    targetSheet.Columns.Item("C").ColumnWidth = 18
    targetSheet.Columns.Item("D").ColumnWidth = 30
    targetSheet.Columns.Item("E").ColumnWidth = 18
    targetSheet.Columns.Item("F").ColumnWidth = 18
    targetSheet.Columns.Item("G").ColumnWidth = 44
End Sub

Private Sub SetControlSheetStaticCells(ByVal targetSheet As Worksheet, ByVal templateVariant As String)
    Dim usageText As String
    Dim checkText As String

    If IsV2Variant(templateVariant) Then
        usageText = "1. run_pdf2excel_v2.bat を実行  2. 対象 PDF を選択  3. 実行前チェックを確認  4. Result / Review / Summary / Errors を確認"
        checkText = "Summary は全体件数、Result は生データ転記結果、Review は確認要行、Errors は失敗した PDF と理由です。"
    Else
        usageText = "1. run_pdf2excel_v1.bat を実行  2. PDF を選択  3. 実行前チェックを確認  4. Result / Summary / Errors を確認"
        checkText = "Summary は件数の全体像、Result は変換成功データ、Errors は失敗した PDF と理由です。"
    End If

    targetSheet.Range("A1").Value = "項目"
    targetSheet.Range("B1").Value = "内容"
    targetSheet.Range("A2").Value = "入力フォルダ"
    targetSheet.Range("A3").Value = "出力ファイル"
    targetSheet.Range("A4").Value = "ログファイル"
    targetSheet.Range("A5").Value = "最終実行日時"
    targetSheet.Range("A6").Value = "状態"
    targetSheet.Range("A7").Value = "対象PDF数"
    targetSheet.Range("A8").Value = "取込データ行数"
    targetSheet.Range("A9").Value = "エラー件数"
    targetSheet.Range("A10").Value = "成功PDF数"
    targetSheet.Range("A11").Value = "失敗PDF数"
    targetSheet.Range("A12").Value = "処理時間(秒)"
    targetSheet.Range("A13").Value = "使用プロファイル"
    targetSheet.Range("A14").Value = "プロファイル説明"
    targetSheet.Range("A15").Value = "システム版"
    targetSheet.Range("A16").Value = "かんたんな使い方"
    targetSheet.Range("B16").Value = usageText
    targetSheet.Range("A17").Value = "確認ポイント"
    targetSheet.Range("B17").Value = checkText
    targetSheet.Range("A18").Value = "注意"
    targetSheet.Range("B18").Value = "文字を選択できるテキスト PDF と、ほぼ同じレイアウトの帳票を想定しています。"
End Sub

Private Sub SetSummarySheetStaticCells(ByVal targetSheet As Worksheet, ByVal templateVariant As String)
    Dim summaryDescription As String

    If IsV2Variant(templateVariant) Then
        summaryDescription = "実行結果の集計、PDF ごとの内訳、Review 件数を表示します。"
    Else
        summaryDescription = "実行結果の集計と、PDFごとの内訳を表示します。"
    End If

    targetSheet.Range("A1").Value = "Summary"
    targetSheet.Range("A2").Value = summaryDescription
    targetSheet.Range("A4").Value = "項目"
    targetSheet.Range("B4").Value = "内容"
    targetSheet.Range("A13").Value = "PDF別サマリー"
    targetSheet.Range("M13").Value = "エラー分類別件数"
End Sub
