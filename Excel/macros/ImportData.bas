Attribute VB_Name = "ImportData"
Option Explicit

' ImportData.bas
' Imports the cleaned/verified CSV produced by the Python scripts
' (data_entry.py / cross_verify.py) into the MasterData sheet of this
' workbook, replacing whatever was there before

Sub ImportRawData()

    Dim wsTarget As Worksheet
    Dim filePath As Variant
    Dim wbSource As Workbook
    Dim lastRow As Long, lastCol As Long

    On Error GoTo ErrHandler

    ' Let the user pick the cleaned CSV file to import
    filePath = Application.GetOpenFilename("CSV Files (*.csv), *.csv", , "Select cleaned data file")
    If filePath = False Then
        MsgBox "Import cancelled.", vbInformation
        Exit Sub
    End If

    ' Prepare the target sheet
    Set wsTarget = ThisWorkbook.Worksheets("MasterData")
    wsTarget.Cells.Clear

    ' Open the source CSV and copy its contents into MasterData
    Application.ScreenUpdating = False
    Set wbSource = Workbooks.Open(filePath)
    wbSource.Sheets(1).UsedRange.Copy wsTarget.Range("A1")
    wbSource.Close SaveChanges:=False

    ' Basic formatting: bold header row, shaded, autofit, filter dropdowns
    With wsTarget
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        .Range(.Cells(1, 1), .Cells(1, lastCol)).Font.Bold = True
        .Range(.Cells(1, 1), .Cells(1, lastCol)).Interior.Color = RGB(217, 217, 217)
        .Columns.AutoFit
        .Rows(1).AutoFilter
    End With

    Application.ScreenUpdating = True
    MsgBox "Import complete: " & (lastRow - 1) & " records loaded.", vbInformation
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Import failed: " & Err.Description, vbCritical

End Sub
