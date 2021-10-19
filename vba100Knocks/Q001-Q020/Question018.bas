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
            Debug.Print "名前：" & wb_item.Name & " 参照範囲：" & ref_to
            wb_item.Delete
            del_cnt = del_cnt + 1
        End If
    Next

    MsgBox ("非表示の名前定義：" & unvisible_cnt & "件" & vbCrLf & _
    "削除した名前定義：" & del_cnt & "件")
    
End Sub
