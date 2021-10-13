Attribute VB_Name = "Module1"
Option Explicit

Sub Q42()
    Sheets("�K�wDB").Cells.Clear

    Sheets("�K�w").Cells(1, 1).CurrentRegion.Copy _
    Destination:=Sheets("�K�wDB").Cells(1, 1)

    Dim dataRange As Range
    Set dataRange = Sheets("�K�wDB").Cells(1, 1).CurrentRegion
    
    Dim i As Long, j As Long
    
    For i = 2 To dataRange.Rows.Count
        For j = 1 To 3
            If Sheets("�K�wDB").Cells(i, j) = "" Then Sheets("�K�wDB").Cells(i, j) = Sheets("�K�wDB").Cells(i - 1, j)
        Next
    Next
    
    For i = dataRange.Rows.Count To 2 Step -1
        If Sheets("�K�wDB").Cells(i, 4) = "" Then Sheets("�K�wDB").Rows(i).Delete
    Next
    
End Sub
