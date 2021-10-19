Attribute VB_Name = "Module1"
Option Explicit

Sub Q5()
    Dim i As Long
    
    For i = 1 To Cells(2, 2).CurrentRegion.Rows.Count
        If Cells(i + 2, 2) <> "" And Cells(i + 2, 3) <> "" Then
             Cells(i + 2, 4).Value = Cells(i + 2, 2).Value * Cells(i + 2, 3).Value
             Cells(i + 2, 4).NumberFormatLocal = "\#,##0"
        End If
    Next

End Sub
