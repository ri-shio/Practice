Attribute VB_Name = "Module1"
Option Explicit

Sub Q60()
    'Functionの動作確認のため、サンプルとして適当なプロシージャを用意した。
    'このプロシージャはアクティブシートのA1～B9を使用する。

    Dim testStr As Variant
    testStr = Array("株）テストテスト", "テストテスト" & ChrW(13183), "テストテスト（株", "テストテスト(株)", "株)テストテスト", "（株）テストテスト", "テストテスト(株", "㈱テストテスト")

    Dim i As Integer
    
    Cells(1, 1) = "変換前"
    Cells(1, 2) = "変換後"
    
    For i = LBound(testStr) To UBound(testStr)
        Cells(i + 2, 1) = testStr(i)
        Cells(i + 2, 2) = varKabu(testStr(i))
    Next
End Sub

Function varKabu(ByVal chkStr As String) As String
    Dim checkList As Variant
    
    checkList = Array(Array("（株）", "*（株）*"), Array("(株)", "*(株)*"), Array("㈱", "*㈱*"), _
    Array(ChrW(13183), "*" & ChrW(13183) & "*"), Array("株）", "株）*"), Array("株)", "株)*"), _
    Array("（株", "*（株"), Array("(株", "*(株"))


    Dim i As Integer
    For i = LBound(checkList) To UBound(checkList)
        If chkStr Like checkList(i)(1) Then
            chkStr = Replace(chkStr, checkList(i)(0), "株式会社")
        End If
    Next
    
    varKabu = chkStr
End Function
