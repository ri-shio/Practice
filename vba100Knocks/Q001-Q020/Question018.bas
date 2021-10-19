Attribute VB_Name = "Module1"
Option Explicit

Sub Q018()
    Dim wb_item As Variant
    Dim del_cnt As Integer: del_cnt = 0
    Dim unvisible_cnt As Integer: unvisible_cnt = 0
    Dim ref_to As String

    For Each wb_item In ActiveWorkbook.Names
        If wb_item.Visible = False Then
            wb_item.Visible = True
            unvisible_cnt = unvisible_cnt + 1
        End If
        ref_to = wb_item.RefersTo
        If ref_to Like "*[#]REF[!]*" Then
            Debug.Print "���O�F" & wb_item.Name & " �Q�Ɣ͈́F" & ref_to
            wb_item.Delete
            del_cnt = del_cnt + 1
        End If
    Next

    MsgBox ("��\���̖��O��`�F" & unvisible_cnt & "��" & vbCrLf & _
    "�폜�������O��`�F" & del_cnt & "��")
    
End Sub
