Attribute VB_Name = "Module1"
Option Explicit

Sub Q6()
    Dim i As Long
    For i = 1 To Cells(1, 1).CurrentRegion.Rows.Count - 1
        If Not Cells(i + 1, 1).Value Like "*-*" Then
            Cells(i + 1, 4).FormulaR1C1 = "=RC[-2] * RC[-1]"
        End If
    Next

End Sub

