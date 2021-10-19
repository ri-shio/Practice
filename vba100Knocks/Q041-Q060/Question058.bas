Attribute VB_Name = "Module1"
Option Explicit

Sub Q58()
    '加工用のテスト配列
    Dim testArray As Variant
    testArray = Array(1, 2, 3, 5, 8, 9, 11, 12, 13, 14, 15, 17, 19, 20, 21, 22)

    'ここで指定した数値以上に1刻みの増加が連続していた場合、1つの要素にまとめる。
    Const rNum As Long = 2

    'Functionを呼び出し、修正前と修正後を比較できるようMsgBoxを表示。
    Dim modArray As Variant
    modArray = modArr(testArray, rNum)

    Dim msg1 As String, msg2 As String
    Dim i As Long

    For i = LBound(testArray) To UBound(testArray)
        msg1 = msg1 & testArray(i) & ", "
    Next
        msg1 = "修正前の配列：{" & Left(msg1, Len(msg1) - 2) & "}"

    For i = LBound(modArray) To UBound(modArray)
        msg2 = msg2 & modArray(i) & ", "
    Next
        msg2 = "修正後の配列：{" & Left(msg2, Len(msg2) - 2) & "}"

    MsgBox msg1 & vbCrLf & msg2

End Sub

Function modArr(ByVal Arr As Variant, ByVal rn As Long) As Variant
    Dim newArr() As Variant
    Dim conCnt As Long: conCnt = 1
    Dim arrCnt As Long: arrCnt = 0
    Dim i As Long, j As Long

    For i = LBound(Arr) + 1 To UBound(Arr)

        'ひとつ前の要素が自分より1低い数（1刻みの増加が続いている）のときは、conCntを増加させる。
        If Arr(i) = Arr(i - 1) + 1 Then
            conCnt = conCnt + 1

            'ひとつ前の要素が自分より1低い数ではない（1刻みの増加ではない・1刻みの増加が終わった）のときは、
            'conCntが既定の数（rNumおよびrn）以上か、それ未満かで判定を分ける。
        Else

        '既定の数以上の場合は数をまとめて新しい配列に格納する。
            If conCnt >= rn Then
                ReDim Preserve newArr(arrCnt)
                newArr(arrCnt) = Arr(i - conCnt) & "-" & Arr(i - 1)
                arrCnt = arrCnt + 1
        '既定の数に満たない場合は、それまでの1つ以上の要素をそのまま新しい配列に格納する。
            Else
                For j = i - conCnt To i - 1
                    ReDim Preserve newArr(arrCnt)
                    newArr(arrCnt) = Arr(j)
                    arrCnt = arrCnt + 1
                Next

            End If

            conCnt = 1

        End If
    Next

    '上記のForループはi - 1までの処理しか行わないため、最終要素の処理は別建てで行う。
    'ただし、上記Forループを抜け出した時点でiは渡された配列（testArrayおよびArr）の最大要素数より1多い状態なため、
    '下記の処理ではi - 1を用いるのが正しい。
    If conCnt > 1 Then
        If conCnt >= rn Then
            ReDim Preserve newArr(arrCnt)
            newArr(arrCnt) = Arr(i - conCnt) & "-" & Arr(i - 1)
        Else
            For j = i - conCnt To i - 1
                ReDim Preserve newArr(arrCnt)
                newArr(arrCnt) = Arr(j)
                arrCnt = arrCnt + 1
            Next

        End If
    Else
        ReDim Preserve newArr(arrCnt)
        newArr(arrCnt) = Arr(i - 1)
    End If

    '作成し終えた配列を返してFunctionは終了。
    modArr = newArr
End Function
