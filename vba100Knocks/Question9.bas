Attribute VB_Name = "Module1"
Option Explicit

Sub Q9()
    Dim ws As Worksheet
    Dim flg As Boolean
    flg = False
    
    For Each ws In Worksheets
        If ws.Name = "���i��" Then flg = True
    Next ws
    
    If flg = False Then
        Set ws = Worksheets.Add
        ws.Name = "���i��"
    End If
    
    Worksheets("���i��").Cells.Clear
    
    Dim i As Long, j As Long
    j = 1
    
    For i = 2 To Worksheets("���ѕ\").Cells(1, 1).CurrentRegion.Rows.Count
        If Worksheets("���ѕ\").Cells(i, 7) = "���i" Then
            Worksheets("���i��").Cells(j, 1) = Worksheets("���ѕ\").Cells(i, 1)
            j = j + 1
        End If
    Next
    
End Sub
