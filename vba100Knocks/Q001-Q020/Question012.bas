Attribute VB_Name = "Module2"
Option Explicit

Sub Q12()
    Dim i As Long, j As Long, k As Long
    Dim mgrange As Variant
    Dim divcount As Integer
    For i = 1 To Cells(1, 1).CurrentRegion.Columns.Count
        For j = 1 To Cells(1, 1).CurrentRegion.Rows.Count
            If Cells(j, i).MergeCells Then
                Set mgrange = Cells(j, i).MergeArea
                Cells(j, i).MergeCells = False
                divcount = Cells(j, i).Value
                
                For k = 0 To mgrange.Count - 1
                    Cells(j + k, i).Value = WorksheetFunction.RoundUp(divcount / (mgrange.Count - k), 0)
                    divcount = divcount - Cells(j + k, i).Value
                Next
            End If
        Next
    Next
End Sub
