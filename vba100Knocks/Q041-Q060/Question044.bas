Attribute VB_Name = "Module1"
Option Explicit

Sub Q44()
    Dim list_sht As Worksheet
    
    Set list_sht = ActiveSheet
    
    list_sht.Cells.Clear
    list_sht.Cells(1, 1) = "�e�[�u����": list_sht.Cells(1, 2) = "�V�[�g��": list_sht.Cells(1, 3) = "�Z���͈�"
    list_sht.Cells(1, 4) = "���X�g�s��": list_sht.Cells(1, 5) = "���X�g��"
    
    Dim sht As Object
    Dim tbl As Object
    Dim tbl_cnt As Integer: tbl_cnt = 1
    
    For Each sht In ThisWorkbook.Sheets
        For Each tbl In sht.ListObjects

            list_sht.Cells(tbl_cnt + 1, 1) = tbl.Name
            list_sht.Cells(tbl_cnt + 1, 2) = tbl.Parent.Name
            list_sht.Cells(tbl_cnt + 1, 3) = tbl.Range.Address
            list_sht.Cells(tbl_cnt + 1, 4) = tbl.ListRows.Count
            list_sht.Cells(tbl_cnt + 1, 5) = tbl.ListColumns.Count
            tbl_cnt = tbl_cnt + 1
        Next
    Next
    
End Sub
