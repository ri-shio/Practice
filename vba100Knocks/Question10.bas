Attribute VB_Name = "Module1"
Option Explicit

Sub Q10()
    Dim i As Long
    For i = Cells(1, 1).CurrentRegion.Rows.Count To 2 Step -1
        If Cells(i, 3) = "" Then
            If Cells(i, 4) Like "*ïsóv*" Or Cells(i, 4) Like "*çÌèú*" Then
                Cells(i, 1).EntireRow.Delete
            End If
        End If
    Next
End Sub

