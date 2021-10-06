Attribute VB_Name = "Module2"
Option Explicit

Sub Q11()
    Dim i As Long, j As Long
    For i = 1 To Cells(1, 1).CurrentRegion.Columns.Count
        For j = 1 To Cells(1, 1).CurrentRegion.Rows.Count
            If Cells(j, i).MergeCells Then
                Cells(j, i).ClearComments
                Cells(j, i).AddComment "åãçáÉZÉã"
            End If
        Next
    Next
End Sub
