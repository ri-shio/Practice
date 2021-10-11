Attribute VB_Name = "Module1"
Option Explicit

Sub Q34()
    '��]�������w�肷��B
    Const r_dir As String = "right"

    '�e�X�g�p�ɔz����쐬�B
    Dim MyArray As Variant
    MyArray = Range(Cells(1, 1), Cells(3, 4))
    
    'Function���Ăяo���A�w�肳�ꂽ�����ɏ]���Ĕz�����]������B
    Call Q34_Rotate(MyArray, r_dir)
    Range(Cells(1, 6), Cells(4, 8)) = MyArray
    
End Sub

Function Q34_Rotate(ByRef MyArray As Variant, r_dir As String) As Variant
    
    '�󂯎�����z��̗v�f�����擾���A�o�b�t�@�p�̔z����쐬�B
    Dim arrayElem1 As Integer, arrayElem2 As Integer
    arrayElem1 = UBound(MyArray, 1): arrayElem2 = UBound(MyArray, 2)
    
    Dim varBuffer As Variant
    ReDim varBuffer(1 To arrayElem2, 1 To arrayElem1)

    '�����̎w���ɏ]���A�o�b�t�@�p�z��ɗv�f���������ށB
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
        
    '�o�b�t�@�p�z�񂩂�l���󂯎��AFunction�͏I���B
    MyArray = varBuffer
End Function
