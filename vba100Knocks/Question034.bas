Attribute VB_Name = "Module1"
Option Explicit

Sub Q34()
    '回転方向を指定する。
    Const r_dir As String = "right"

    'テスト用に配列を作成。
    Dim MyArray As Variant
    MyArray = Range(Cells(1, 1), Cells(3, 4))
    
    'Functionを呼び出し、指定された方向に従って配列を回転させる。
    Call Q34_Rotate(MyArray, r_dir)
    Range(Cells(1, 6), Cells(4, 8)) = MyArray
    
End Sub

Function Q34_Rotate(ByRef MyArray As Variant, r_dir As String) As Variant
    
    '受け取った配列の要素数を取得し、バッファ用の配列を作成。
    Dim arrayElem1 As Integer, arrayElem2 As Integer
    arrayElem1 = UBound(MyArray, 1): arrayElem2 = UBound(MyArray, 2)
    
    Dim varBuffer As Variant
    ReDim varBuffer(1 To arrayElem2, 1 To arrayElem1)

    '方向の指示に従い、バッファ用配列に要素を書き込む。
    Dim i As Integer, j As Integer
        For i = 1 To arrayElem2
            For j = 1 To arrayElem1
                If r_dir = "right" Then
                    varBuffer(i, arrayElem1 - j + 1) = MyArray(j, i)
                ElseIf r_dir = "left" Then
                    varBuffer(arrayElem2 - i + 1, j) = MyArray(j, i)
                End If
            Next
        Next
        
    'バッファ用配列から値を受け取り、Functionは終了。
    MyArray = varBuffer
End Function
