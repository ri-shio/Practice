Attribute VB_Name = "Module1"
Option Explicit

Sub Q50()
    Dim tribonacci() As LongLong
    ReDim tribonacci(1 To 3)
    tribonacci(1) = 0: tribonacci(2) = 0: tribonacci(3) = 1

    Dim i As Long

    'この回数繰り返す。76回以上はオーバーフロー。
    Const toNum As Long = 75

    For i = 1 To toNum
    On Error GoTo ErrHandler
        If i >= 4 Then
            ReDim Preserve tribonacci(1 To i)
            tribonacci(i) = tribonacci(i - 3) + tribonacci(i - 2) + tribonacci(i - 1)
        End If

        Cells(i, 1) = tribonacci(i)
    Next

ErrHandler:
End Sub
