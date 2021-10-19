Attribute VB_Name = "Module2"
Option Explicit

Sub Q33()
    '�񕪒T���𗘗p���邽�߁A�}�X�^�R�[�h�ŏ����ɕ��בւ���B
    Sheets("�}�X�^").Cells(1, 1).Sort key1:=Sheets("�}�X�^").Cells(1, 1), order1:=xlAscending, Header:=xlYes

    '�f�[�^�E�}�X�^�̗񐔂��擾���A�܂��A�l�����ꂼ��z��Ɋi�[����B
    Dim masterTable As Variant
    Dim masterRowNum As Long
    Dim dataTable As Variant
    Dim dataRowNum As Long
    
    masterRowNum = Sheets("�}�X�^").Cells(1, 1).CurrentRegion.Rows.Count
    Sheets("�}�X�^").Activate
    masterTable = Sheets("�}�X�^").Range(Cells(2, 1), Cells(masterRowNum, 3)).Value
    
    dataRowNum = Sheets("�f�[�^").Cells(1, 1).CurrentRegion.Rows.Count
    Sheets("�f�[�^").Activate
    dataTable = Sheets("�f�[�^").Range(Cells(2, 1), Cells(dataRowNum, 3)).Value
    
    '�}�X�^�̒T�����ʂ��i�[���邽�߂̔z����`����B
    Dim dataTable2 As Variant
    ReDim dataTable2(1 To dataRowNum, 1 To 2)
    
    '�񕪒T���ɂ�荂���ɒT�����s���B
    Dim lNum As Long, hNum As Long
    Dim i As Long
    
    For i = 1 To dataRowNum - 1
        lNum = 1: hNum = masterRowNum - 1
        Do
            If dataTable(i, 2) = masterTable(Int((lNum + hNum) / 2), 1) Then
                dataTable2(i, 1) = masterTable(Int((lNum + hNum) / 2), 2)
                dataTable2(i, 2) = masterTable(Int((lNum + hNum) / 2), 3)
                Exit Do
            ElseIf dataTable(i, 2) > masterTable(Int((lNum + hNum) / 2), 1) Then
                lNum = Int((lNum + hNum) / 2)
            ElseIf dataTable(i, 2) < masterTable(Int((lNum + hNum) / 2), 1) Then
                hNum = Int((lNum + hNum) / 2)
            End If
        Loop
    Next
    
    '�z�񂩂�l�����o���A�Z���ɓ��͂���B
    Sheets("�f�[�^").Activate
    Sheets("�f�[�^").Range(Cells(2, 4), Cells(masterRowNum, 5)).Value = dataTable2
    
    '�f�[�^�V�[�g��F��ɂ͌v�Z������͂���B
    With Sheets("�f�[�^").Cells(2, 6)
        .FormulaR1C1 = "=RC[-3]*RC[-1]"
        .AutoFill Destination:=Range(Cells(2, 6), Cells(dataRowNum, 6))
    End With
    
    '�f�[�^�V�[�g��A1��I�����A�}�N���͏I���B
    Sheets("�f�[�^").Cells(1, 1).Activate
End Sub
