Attribute VB_Name = "Module1"
Option Explicit

Sub Q23()
    '����Ώۂ̃��[�N�u�b�N�͒萔�Ƃ��Ē�`�B
    Const wb1_name As String = "Book_20201101.xlsx"
    Const wb2_name As String = "Book_20201102.xlsx"
    
    'wb1,wb2�ɂ��ꂼ�ꔻ�肷�郏�[�N�u�b�N���Z�b�g����B
    Dim wb1 As Workbook, wb2 As Workbook
    Dim wb1_sheets() As String, wb2_sheets() As String

    Workbooks.Open ThisWorkbook.Path & "\" & wb1_name
    Set wb1 = Workbooks(wb1_name)

    Workbooks.Open ThisWorkbook.Path & "\" & wb2_name
    Set wb2 = Workbooks(wb2_name)
    
    'wb1_sheets,wb2_sheets�̔z�񐔂𓮓I�Ɍ��肵�A�V�[�g�����擾����B
    ReDim wb1_sheets(1 To wb1.Sheets.Count), wb2_sheets(1 To wb2.Sheets.Count)
    
    Dim i As Integer
    Dim ws As Object
    
    i = 1
    For Each ws In wb1.Sheets
        wb1_sheets(i) = ws.Name
        i = i + 1
    Next
    
    i = 1
    For Each ws In wb2.Sheets
        wb2_sheets(i) = ws.Name
        i = i + 1
    Next
    
    '���ꂼ��̔z����ŃV�[�g���̏��Ԃ𐮂���B
    Dim j As Integer
    Dim swap As String
    
    For i = 1 To wb1.Sheets.Count
        For j = wb1.Sheets.Count To i Step -1
            If wb1_sheets(i) > wb1_sheets(j) Then
                swap = wb1_sheets(i)
                wb1_sheets(i) = wb1_sheets(j)
                wb1_sheets(j) = swap
            End If
        Next
    Next
    
    For i = 1 To wb2.Sheets.Count
        For j = wb2.Sheets.Count To i Step -1
            If wb2_sheets(i) > wb2_sheets(j) Then
                swap = wb2_sheets(i)
                wb2_sheets(i) = wb2_sheets(j)
                wb2_sheets(j) = swap
            End If
        Next
    Next
    
    '�V�[�g������уV�[�g�����r���A�s��v������������False�Ƃ���B�ǂ������v�̏ꍇ��True���c��B
    Dim isSame As Boolean: isSame = True
    
    If wb1.Sheets.Count <> wb2.Sheets.Count Then isSame = False
    
    For i = 1 To wb1.Sheets.Count
        If wb1_sheets(i) <> wb2_sheets(i) Then isSame = False
    Next
    
    'T/F�ɏ]�����b�Z�[�W���쐬���A���b�Z�[�W�{�b�N�X�ɕ\���B����Ώۂ̃��[�N�u�b�N������B
    Dim message As String
    If isSame Then
        message = "��v"
    Else
        message = "�s��v"
    End If
    
    wb1.Close
    wb2.Close
    
    MsgBox (message)
    
End Sub
