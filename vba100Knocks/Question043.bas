Attribute VB_Name = "Module1"
Option Explicit

Sub Q43()
    '���Excel�{�̂��I�������邽�߁A�܂��ۑ�������B
    ThisWorkbook.Save
    
    '�A�N�e�B�u�V�[�g���f�[�^�V�[�g�Ƃ��A���`�p�ɐV�������[�N�V�[�g���쐬����B
    Dim data_ws As Worksheet
    Dim csvout_ws As Worksheet
    
    Set data_ws = ActiveSheet
    Set csvout_ws = Worksheets.Add
    
    data_ws.Cells(1, 1).CurrentRegion.Copy _
    Destination:=csvout_ws.Cells(1, 1)
    
    '�v���ɏ]���AA���yyyy-mm-dd�`���ɁAB����J���}���������AC����J���}��������2���Ƃ���B
    csvout_ws.Columns(1).NumberFormatLocal = "yyyy-mm-dd"
    csvout_ws.Columns(2).NumberFormatLocal = "0"
    csvout_ws.Columns(3).NumberFormatLocal = "0.00"
    
    '�v���ɏ]���AD����̃_�u���N�H�[�e�[�V�������G�X�P�[�v�ł���悤�ɂ���B
    Dim i As Long
    
    For i = 1 To csvout_ws.Cells(1, 4).CurrentRegion.Count
        If InStr(csvout_ws.Cells(i, 4), """") > 0 Then
            csvout_ws.Cells(i, 4) = "" & Replace(csvout_ws.Cells(i, 4), "", "" & "") & ""
        End If
    Next
    
    '���b�Z�[�W�m�F��False�ɂ��ACSV�G�N�X�|�[�g�E���`�p�V�[�g�폜�AExcel�{�̂̏I�����s���B
    Application.DisplayAlerts = False

    ThisWorkbook.SaveAs _
    Filename:=ThisWorkbook.Path & "\CSVOutput.csv", _
    FileFormat:=xlCSV, _
    local:=True

    csvout_ws.Delete
    Application.Quit
    
End Sub
