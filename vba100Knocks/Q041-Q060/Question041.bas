Attribute VB_Name = "Module1"
Option Explicit

Sub Q41()

    'Rnd�֐���p���Ĕ퉉�Z�q1��0�`99�A�퉉�Z�q2��0�`9�͈̔͂Ń����_���Ɍ��肷��B
    '�܂��A���l��Rnd�֐���p���ĉ��Z�q�����肷��B
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
        
        '0�̏��Z�������B
        If operand(1) = 0 And operator = "/" Then operand(1) = 2
        
        'InputBox�Ń��[�U�[���當������擾�B
        answer = InputBox(operand(0) & operator & operand(1) & "=?", "�ÎZ��� ��" & i & "�� �i����Z�͏����_��1�ʂ�L�������Ƃ��A�l�̌ܓ����邱�ƁB�j")
        
        '�v�Z���ʂ̔���
        If IsNumeric(StrConv(answer, vbNarrow)) Then
            If CDbl(StrConv(answer, vbNarrow)) = Round(Evaluate(operand(0) & operator & operand(1)), 1) Then score = score + 1
        End If
        
    Next

    '�_����\�����A�}�N���͏I���B
    MsgBox ("���Ȃ��̓_����10�_���_��" & score & "�_�ł��B")
End Sub
