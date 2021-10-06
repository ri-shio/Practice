Attribute VB_Name = "Module1"
Option Explicit

Sub Q21()
    Dim keep_date(30) As Long
    Dim buckup_date As Long
    Dim i As Integer
    Dim buf As String
    Dim flg_keep As Boolean: flg_keep = False
    
    For i = 0 To 30
        keep_date(i) = Format(Date - i, "yyyymmdd")
    Next

    buf = Dir(ThisWorkbook.Path & "\BUCKUP\*.xlsm")
    Do
        If buf = "" Then Exit Do
        buckup_date = Left(Right(buf, 17), 8)
        For i = 0 To 30
            If keep_date(i) = buckup_date Then
                flg_keep = True
                Exit For
            End If
        Next
        If flg_keep = False Then Kill ThisWorkbook.Path & "\BUCKUP\" & buf
        buf = Dir()
    Loop

End Sub
