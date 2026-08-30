Attribute VB_Name = "PivotBuilder"
Option Explicit

' PivotBuilder.bas
' Builds (or rebuilds) a summary pivot table from MasterData into a
' PivotSummary sheet, and offers a quick refresh for existing pivots.
'
' Edit ROW_FIELD / DATA_FIELD below to match your actual column headers.

Sub BuildPivotTable()

    Const ROW_FIELD As String = "Category"    ' column to group rows by
    Const DATA_FIELD As String = "Record ID"  ' column to count

    Dim wsData As Worksheet, wsPivot As Worksheet
    Dim pCache As PivotCache
    Dim pTable As PivotTable
    Dim dataRange As Range
    Dim lastRow As Long, lastCol As Long

    Set wsData = ThisWorkbook.Worksheets("MasterData")

    With wsData
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        Set dataRange = .Range(.Cells(1, 1), .Cells(lastRow, lastCol))
    End With

    If lastRow < 2 Then
        MsgBox "No data available to summarize.", vbExclamation
        Exit Sub
    End If

    ' Remove any existing PivotSummary sheet, then recreate it
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("PivotSummary").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set wsPivot = ThisWorkbook.Worksheets.Add
    wsPivot.Name = "PivotSummary"

    ' Build the pivot cache and table
    Set pCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, SourceData:=dataRange)

    Set pTable = pCache.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="RecordsSummary")

    With pTable
        .PivotFields(ROW_FIELD).Orientation = xlRowField
        With .PivotFields(DATA_FIELD)
            .Orientation = xlDataField
            .Function = xlCount
            .NumberFormat = "#,##0"
        End With
    End With

    wsPivot.Range("A1").Value = "Records Summary - Generated " & Format(Now, "mm/dd/yyyy hh:mm")
    wsPivot.Columns.AutoFit

    MsgBox "Pivot table built successfully.", vbInformation

End Sub

Sub RefreshAllPivots()
    ' Run this after re-importing data instead of rebuilding from scratch.
    Dim pc As PivotCache
    For Each pc In ThisWorkbook.PivotCaches
        pc.Refresh
    Next pc
    MsgBox "All pivot tables refreshed.", vbInformation
End Sub
