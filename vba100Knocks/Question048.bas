Attribute VB_Name = "Module1"
Option Explicit

Sub Q48()
    Dim test_var As Long
    Dim test_arr1(2) As Variant
    Dim test_arr2(2, 2) As Variant
    Dim test_arr3(2, 2, 2) As Variant
    Dim result_func As Variant
    Dim i As Integer, j As Integer, k As Integer

    test_var = 3
    For i = 0 To 2
        test_arr1(i) = -1.5 + 1.5 * i

        For j = 0 To 1
            test_arr2(i, j) = -3 + 1.5 * i + 2.3 * j

            For k = 0 To 2
                test_arr3(i, j, k) = -5 + 1.5 * i + 2.3 * j + 0.8 + k
            Next
        Next
    Next

    test_arr2(0, 2) = "文字列テスト"
    test_arr3(1, 2, 1) = "文字列テスト2"

    result_func = adjustToInt(test_var)
    MsgBox ("操作前の値：" & test_var & vbCrLf & "操作後の値：" & result_func)

    result_func = adjustToInt(test_arr1)
    MsgBox ("操作前の値：" & test_arr1(0) & "," & test_arr1(1) & "," & test_arr1(2) & vbCrLf & _
    "操作後の値：" & result_func(0) & "," & result_func(1) & "," & result_func(2))

    result_func = adjustToInt(test_arr2)
    MsgBox ("操作前の値：" & vbCrLf & _
    test_arr2(0, 0) & "," & test_arr2(0, 1) & "," & test_arr2(0, 2) & vbCrLf & _
    test_arr2(1, 0) & "," & test_arr2(1, 1) & "," & test_arr2(1, 2) & vbCrLf & _
    test_arr2(2, 0) & "," & test_arr2(2, 1) & "," & test_arr2(2, 2) & vbCrLf & _
    "操作後の値：" & vbCrLf & _
    result_func(0, 0) & "," & result_func(0, 1) & "," & result_func(0, 2) & vbCrLf & _
    result_func(1, 0) & "," & result_func(1, 1) & "," & result_func(1, 2) & vbCrLf & _
    result_func(2, 0) & "," & result_func(2, 1) & "," & result_func(2, 2))

    result_func = adjustToInt(test_arr3)
    MsgBox ("要素数が多いため一部のみ表示" & vbCrLf & _
    "操作前の値(2,0,1)：" & test_arr3(2, 0, 1) & vbCrLf & _
    "操作後の値(2,0,1)：" & result_func(2, 0, 1) & vbCrLf & _
    "操作前の値(1,2,1)：" & test_arr3(1, 2, 1) & vbCrLf & _
    "操作後の値(1,2,1)：" & result_func(1, 2, 1))

End Sub

Function adjustToInt(ByVal varObj As Variant) As Variant

    Dim i As Long: i = 1
    Dim buffer As Variant

    Err.Clear

    On Error Resume Next
    Do While Err.Number = 0

        buffer = UBound(varObj, i)
        i = i + 1

    Loop
    On Error GoTo 0

    Dim j As Long, k As Long

    If i = 3 Then
        For j = LBound(varObj) To UBound(varObj)
            If IsNumeric(varObj(j)) Then varObj(j) = WorksheetFunction.RoundDown(varObj(j), 0)
        Next
    ElseIf i = 4 Then
        For j = LBound(varObj, 1) To UBound(varObj, 1)
            For k = LBound(varObj, 2) To UBound(varObj, 2)
                If IsNumeric(varObj(j, k)) Then varObj(j, k) = WorksheetFunction.RoundDown(varObj(j, k), 0)
            Next
        Next
    End If

        adjustToInt = varObj
End Function
