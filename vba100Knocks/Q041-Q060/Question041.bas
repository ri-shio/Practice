Attribute VB_Name = "Module1"
Option Explicit

Sub Q41()

    'Rnd関数を用いて被演算子1を0～99、被演算子2を0～9の範囲でランダムに決定する。
    'また、同様にRnd関数を用いて演算子を決定する。
    Dim operand(1) As Integer: Dim operator As String
    Dim answer As String
    
    Dim i As Integer
    Dim score As Integer: score = 0
    
    For i = 1 To 10
        Randomize
        operand(0) = Rnd * 100: operand(1) = Rnd * 10
        operator = Rnd * 4
        Select Case operator
            Case Is < 1
                operator = "+"
            Case Is < 2
                operator = "-"
            Case Is < 3
                operator = "*"
            Case Else
                operator = "/"
        End Select
        
        '0の除算を避ける。
        If operand(1) = 0 And operator = "/" Then operand(1) = 2
        
        'InputBoxでユーザーから文字列を取得。
        answer = InputBox(operand(0) & operator & operand(1) & "=?", "暗算問題 第" & i & "問 （割り算は小数点第1位を有効数字とし、四捨五入すること。）")
        
        '計算結果の判定
        If IsNumeric(StrConv(answer, vbNarrow)) Then
            If CDbl(StrConv(answer, vbNarrow)) = Round(Evaluate(operand(0) & operator & operand(1)), 1) Then score = score + 1
        End If
        
    Next

    '点数を表示し、マクロは終了。
    MsgBox ("あなたの点数は10点満点中" & score & "点です。")
End Sub
