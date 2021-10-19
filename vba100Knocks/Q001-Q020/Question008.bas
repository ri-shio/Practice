Attribute VB_Name = "Module1"
Option Explicit

Sub Q8()
    Dim i As Long
    For i = 2 To Cells(1, 1).CurrentRegion.Rows.Count
        If WorksheetFunction.Sum(Range(Cells(i, 2), Cells(i, 6))) >= 350 And WorksheetFunction.Min(Cells(i, 2), Cells(i, 6)) >= 50 Then
            Cells(i, 7) = "‡Ši"
        End If
    Next
End Sub
