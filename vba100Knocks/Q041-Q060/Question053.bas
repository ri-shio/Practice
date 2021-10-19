Attribute VB_Name = "Module1"
Option Explicit

Sub Q53()
    Dim dateThisYear As String
    dateThisYear = CDate("12/31/" & Format(Date, "yyyy"))

    Dim i As Long
    For i = 2 To Cells(1, 1).CurrentRegion.Rows.Count
        If Cells(i, 2) = "男" And DateDiff("yyyy", CDate(Cells(i, 3)), dateThisYear) >= 35 And Cells(i, 4) = "東京都" Then
            Cells(i, 5) = "対象"
        End If
    Next
    
End Sub
