Attribute VB_Name = "Module1"
Option Explicit

Sub Q7()
'�𓚂����ăs���I�h��؂��Date�^�ɔF���ł���悤�C���B
    Dim i As Long
    Dim d As Variant
    For i = 2 To Cells(2, 1).CurrentRegion.Rows.Count
        d = Replace(Cells(i, 1).Value, ".", "/")
        If IsDate(d) Then
            Cells(i, 2) = CDate(Year(d) & "/" & Month(d) + 1 & "/1") - 1
            Cells(i, 2).NumberFormatLocal = "mmdd"
        End If
    Next

End Sub
