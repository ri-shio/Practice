Attribute VB_Name = "Module1"
Option Explicit

Sub Q9()
    Dim ws As Worksheet
    Dim flg As Boolean
    flg = False

    For Each ws In Worksheets
        If ws.Name = "合格者" Then flg = True
    Next ws

    If flg = False Then
        Set ws = Worksheets.Add
        ws.Name = "合格者"
    End If

    Worksheets("合格者").Cells.Clear

    Dim i As Long, j As Long
    j = 1

    For i = 2 To Worksheets("成績表").Cells(1, 1).CurrentRegion.Rows.Count
        If Worksheets("成績表").Cells(i, 7) = "合格" Then
            Worksheets("合格者").Cells(j, 1) = Worksheets("成績表").Cells(i, 1)
            j = j + 1
        End If
    Next

End Sub
