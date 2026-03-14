Attribute VB_Name = "PDF2ExcelTemplateBuilder"
Option Explicit

Private Const CONTROL_SHEET As String = "Control"
Private Const RESULT_SHEET As String = "Result"
Private Const ERRORS_SHEET As String = "Errors"
Private Const SUMMARY_SHEET As String = "Summary"

Public Sub BuildPDF2ExcelTemplateInActiveWorkbook()
    On Error GoTo EH

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    EnsureWorkbookHasRequiredSheets ThisWorkbook

    SetControlSheetLayout ThisWorkbook.Worksheets.Item(1)
    SetDataSheetLayout ThisWorkbook.Worksheets.Item(2), RESULT_SHEET, _
        "変換成功データがここに読み込まれます。A列はPDFファイル名、B列以降はプロファイルに応じた表データです。"
    SetDataSheetLayout ThisWorkbook.Worksheets.Item(3), ERRORS_SHEET, _
        "失敗したPDFと、分類されたエラー理由がここに一覧表示されます。"
    SetSummarySheetLayout ThisWorkbook.Worksheets.Item(4)

    ThisWorkbook.Worksheets(CONTROL_SHEET).Activate
    SetStatus "待機中"

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
EH:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Err.Raise Err.Number, "BuildPDF2ExcelTemplateInActiveWorkbook", Err.Description
End Sub

Public Sub RefreshAndBuildWorkbook()
    On Error GoTo EH

    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    SetStatus "Refreshing"

    ThisWorkbook.RefreshAll
    Application.CalculateUntilAsyncQueriesDone

    ThisWorkbook.Worksheets(RESULT_SHEET).Columns.AutoFit
    ThisWorkbook.Worksheets(ERRORS_SHEET).Columns.AutoFit
    ThisWorkbook.Worksheets(SUMMARY_SHEET).Columns.AutoFit
    SetStatus "Refreshed"
    Application.ScreenUpdating = True

    Exit Sub
EH:
    Application.ScreenUpdating = True
    SetStatus "MacroError"
    Err.Raise Err.Number, "RefreshAndBuildWorkbook", Err.Description
End Sub

Public Sub ExportResultAsXlsx()
    On Error GoTo EH

    Dim outputPath As String
    Dim outputWorkbook As Workbook
    Dim folderPath As String

    outputPath = Trim$(CStr(ControlSheet.Range("B3").Value))
    If Len(outputPath) = 0 Then
        Err.Raise vbObjectError + 700, "ExportResultAsXlsx", "Control!B3 does not contain an output path."
    End If

    folderPath = Left$(outputPath, InStrRev(outputPath, "\") - 1)
    If Len(folderPath) > 0 Then
        EnsureFolderTree folderPath
    End If

    SetStatus "Exporting"
    ThisWorkbook.Worksheets(Array(CONTROL_SHEET, SUMMARY_SHEET, RESULT_SHEET, ERRORS_SHEET)).Copy
    Set outputWorkbook = ActiveWorkbook

    Application.DisplayAlerts = False
    outputWorkbook.SaveAs Filename:=outputPath, FileFormat:=xlOpenXMLWorkbook
    outputWorkbook.Close SaveChanges:=False
    Application.DisplayAlerts = True

    SetStatus "Exported"
    Exit Sub
EH:
    SetStatus "MacroError"
    Application.DisplayAlerts = True
    Err.Raise Err.Number, "ExportResultAsXlsx", Err.Description
End Sub

Private Function ControlSheet() As Worksheet
    Set ControlSheet = ThisWorkbook.Worksheets(CONTROL_SHEET)
End Function

Private Sub SetStatus(ByVal statusText As String)
    ControlSheet.Range("B6").Value = statusText
    ControlSheet.Range("B5").Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Sub

Private Sub EnsureFolderTree(ByVal folderPath As String)
    Dim normalizedPath As String
    Dim parts() As String
    Dim currentPath As String
    Dim i As Long

    normalizedPath = Replace(folderPath, "/", "\")
    If Len(normalizedPath) = 0 Then Exit Sub

    parts = Split(normalizedPath, "\")
    If UBound(parts) < 0 Then Exit Sub

    currentPath = parts(0)
    If Right$(currentPath, 1) = ":" Then
        currentPath = currentPath & "\"
    End If

    For i = 1 To UBound(parts)
        If Len(parts(i)) > 0 Then
            If Right$(currentPath, 1) <> "\" Then currentPath = currentPath & "\"
            currentPath = currentPath & parts(i)
            If Dir$(currentPath, vbDirectory) = vbNullString Then
                MkDir currentPath
            End If
        End If
    Next i
End Sub

Private Sub EnsureWorkbookHasRequiredSheets(ByVal targetWorkbook As Workbook)
    Do While targetWorkbook.Worksheets.Count < 4
        targetWorkbook.Worksheets.Add After:=targetWorkbook.Worksheets(targetWorkbook.Worksheets.Count)
    Loop

    Do While targetWorkbook.Worksheets.Count > 4
        targetWorkbook.Worksheets.Item(targetWorkbook.Worksheets.Count).Delete
    Loop
End Sub

Private Sub SetControlSheetLayout(ByVal targetSheet As Worksheet)
    targetSheet.Cells.Clear
    targetSheet.Name = CONTROL_SHEET

    SetControlSheetStaticCells targetSheet

    targetSheet.Range("A1:B1").Font.Bold = True
    targetSheet.Range("A1:B1").Interior.Color = 15773696
    targetSheet.Range("A16:A18").Font.Bold = True
    targetSheet.Columns.Item("A").ColumnWidth = 18
    targetSheet.Columns.Item("B").ColumnWidth = 92
    targetSheet.Range("A1:B18").VerticalAlignment = -4160
    targetSheet.Range("A1:B18").WrapText = True
    targetSheet.Range("B5").Value = ""
    targetSheet.Range("B6").Value = "待機中"
    targetSheet.Range("B7:B14").Value = ""
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

Private Sub SetSummarySheetLayout(ByVal targetSheet As Worksheet)
    targetSheet.Cells.Clear
    targetSheet.Name = SUMMARY_SHEET

    SetSummarySheetStaticCells targetSheet

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

Private Sub SetControlSheetStaticCells(ByVal targetSheet As Worksheet)
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
    targetSheet.Range("A16").Value = "かんたんな使い方"
    targetSheet.Range("B16").Value = "1. run_pdf2excel.bat を実行  2. PDF を選択  3. 実行前チェックを確認  4. Result / Summary / Errors を確認"
    targetSheet.Range("A17").Value = "確認ポイント"
    targetSheet.Range("B17").Value = "Summary は件数の全体像、Result は変換成功データ、Errors は失敗した PDF と理由です。"
    targetSheet.Range("A18").Value = "注意"
    targetSheet.Range("B18").Value = "文字を選択できるテキスト PDF と、ほぼ同じレイアウトの帳票を想定しています。"
End Sub

Private Sub SetSummarySheetStaticCells(ByVal targetSheet As Worksheet)
    targetSheet.Range("A1").Value = "Summary"
    targetSheet.Range("A2").Value = "実行結果の集計と、PDFごとの内訳を表示します。"
    targetSheet.Range("A4").Value = "項目"
    targetSheet.Range("B4").Value = "内容"
    targetSheet.Range("A13").Value = "PDF別サマリー"
    targetSheet.Range("M13").Value = "エラー分類別件数"
End Sub
