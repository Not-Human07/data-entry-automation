Attribute VB_Name = "ValidationRules"
Option Explicit

' ValidationRules.bas
' Applies data validation (dropdown list) and conditional-formatting
' checks (duplicate IDs, blank required fields) to the MasterData sheet.
'
' Edit COL_ID / COL_STATUS below to match your actual column layout.

Sub ApplyValidationRules()

    Const COL_ID As String = "A"       ' Record ID column
    Const COL_STATUS As String = "F"   ' Status column

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rngStatus As Range

    Set ws = ThisWorkbook.Worksheets("MasterData")
    lastRow = ws.Cells(ws.Rows.Count, COL_ID).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "No data found to validate.", vbExclamation
        Exit Sub
    End If

    ' 1. Add a dropdown list to the Status column
    Set rngStatus = ws.Range(COL_STATUS & "2:" & COL_STATUS & lastRow)
    With rngStatus.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="Verified,Pending,Flagged"
        .InputTitle = "Record Status"
        .InputMessage = "Choose a status for this record."
        .ShowInput = True
    End With

    ' 2. Highlight duplicate Record IDs
    HighlightDuplicates ws, COL_ID, lastRow

    ' 3. Highlight blank Status entries
    HighlightBlanks ws, COL_STATUS, lastRow

    MsgBox "Validation rules applied to " & (lastRow - 1) & " records.", vbInformation

End Sub

Private Sub HighlightDuplicates(ws As Worksheet, col As String, lastRow As Long)
    Dim rng As Range
    Set rng = ws.Range(col & "2:" & col & lastRow)
    rng.FormatConditions.Delete
    rng.FormatConditions.AddUniqueValues
    rng.FormatConditions(rng.FormatConditions.Count).DupeUnique = xlDuplicate
    rng.FormatConditions(rng.FormatConditions.Count).Interior.Color = RGB(255, 199, 206)
End Sub

Private Sub HighlightBlanks(ws As Worksheet, col As String, lastRow As Long)
    Dim rng As Range
    Set rng = ws.Range(col & "2:" & col & lastRow)
    rng.FormatConditions.Delete
    rng.FormatConditions.Add Type:=xlBlanksCondition
    rng.FormatConditions(rng.FormatConditions.Count).Interior.Color = RGB(255, 235, 156)
End Sub
