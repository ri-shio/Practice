Attribute VB_Name = "Module1"
Option Explicit

Sub Q42()
    Sheets("階層DB").Cells.Clear

    Sheets("階層").Cells(1, 1).CurrentRegion.Copy _
    Destination:=Sheets("階層DB").Cells(1, 1)

    Dim dataRange As Range
    Set dataRange = Sheets("階層DB").Cells(1, 1).CurrentRegion
    
    Dim i As Long, j As Long
    
    For i = 2 To dataRange.Rows.Count
        For j = 1 To 3
            If Sheets("階層DB").Cells(i, j) = "" Then Sheets("階層DB").Cells(i, j) = Sheets("階層DB").Cells(i - 1, j)
        Next
    Next
    
    For i = dataRange.Rows.Count To 2 Step -1
        If Sheets("階層DB").Cells(i, 4) = "" Then Sheets("階層DB").Rows(i).Delete
    Next
    
End Sub
