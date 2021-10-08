Attribute VB_Name = "Module1"
Option Explicit

Sub Q27()
    Dim ws_links As Object
    Dim i As Integer
    
    Set ws_links = ActiveSheet.Hyperlinks
    
    For i = ws_links.Count To 1 Step -1
        If ws_links.Item(i).Type = 0 Then
            Cells(ws_links.Item(i).Range.Row, ws_links.Item(i).Range.Column + 1) = ws_links.Item(i).Address
            ws_links.Item(i).Delete
        End If
    Next
End Sub

