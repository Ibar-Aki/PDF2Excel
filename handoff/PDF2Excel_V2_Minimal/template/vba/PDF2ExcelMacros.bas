Attribute VB_Name = "PDF2ExcelMacros"
Option Explicit

Private Const CONTROL_SHEET As String = "Control"
Private Const RESULT_SHEET As String = "Result"
Private Const ERRORS_SHEET As String = "Errors"
Private Const SUMMARY_SHEET As String = "Summary"
Private Const REVIEW_SHEET As String = "Review"

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
    If WorksheetExists(REVIEW_SHEET) Then
        ThisWorkbook.Worksheets(REVIEW_SHEET).Columns.AutoFit
    End If
    SetStatus "Refreshed"
    Application.ScreenUpdating = True

    Exit Sub
EH:
    Application.ScreenUpdating = True
    SetStatus "MacroError"
    Err.Raise Err.Number, "RefreshAndBuildWorkbook", Err.Description
End Sub

Private Function WorksheetExists(ByVal worksheetName As String) As Boolean
    Dim sheet As Worksheet

    On Error Resume Next
    Set sheet = ThisWorkbook.Worksheets(worksheetName)
    WorksheetExists = Not sheet Is Nothing
    Set sheet = Nothing
    On Error GoTo 0
End Function

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
    If WorksheetExists(REVIEW_SHEET) Then
        ThisWorkbook.Worksheets(Array(CONTROL_SHEET, SUMMARY_SHEET, RESULT_SHEET, ERRORS_SHEET, REVIEW_SHEET)).Copy
    Else
        ThisWorkbook.Worksheets(Array(CONTROL_SHEET, SUMMARY_SHEET, RESULT_SHEET, ERRORS_SHEET)).Copy
    End If
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
