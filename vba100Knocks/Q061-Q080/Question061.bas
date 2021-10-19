Attribute VB_Name = "Module1"
Option Explicit

Sub Q61()
    Dim wsData As Worksheet, wsMaster As Worksheet
    Set wsData = Sheets("data"): Set wsMaster = Sheets("マスタ")

    Dim i As Long
    
    For i = 2 To wsData.Cells(2, 1).CurrentRegion.Rows.Count
        If wsMaster.Columns(1).Find(what:=wsData.Cells(i, 1).Value, lookat:=xlWhole, MatchCase:=True) Is Nothing Then
            wsData.Cells(i, 1).Font.Color = vbRed
        Else
            wsData.Cells(i, 1).Font.ColorIndex = xlAutomatic
            wsData.Cells(i, 1).Phonetic.Text = wsMaster.Columns(1).Find(what:=wsData.Cells(i, 1).Value, lookat:=xlWhole, MatchCase:=True).Phonetic.Text
        End If
    Next
End Sub
