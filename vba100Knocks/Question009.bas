Attribute VB_Name = "Module1"
Option Explicit

Sub Q9()
    Dim ws As Worksheet
    Dim flg As Boolean
    flg = False
    
    For Each ws In Worksheets
        If ws.Name = "çáäié“" Then flg = True
    Next ws
    
    If flg = False Then
        Set ws = Worksheets.Add
        ws.Name = "çáäié“"
    End If
    
    Worksheets("çáäié“").Cells.Clear
    
    Dim i As Long, j As Long
    j = 1
    
    For i = 2 To Worksheets("ê¨ê—ï\").Cells(1, 1).CurrentRegion.Rows.Count
        If Worksheets("ê¨ê—ï\").Cells(i, 7) = "çáäi" Then
            Worksheets("çáäié“").Cells(j, 1) = Worksheets("ê¨ê—ï\").Cells(i, 1)
            j = j + 1
        End If
    Next
    
End Sub
