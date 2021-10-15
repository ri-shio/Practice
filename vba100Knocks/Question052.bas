Attribute VB_Name = "Module1"
Option Explicit

Sub Q52()
    Dim sht As Object
    Dim print_sht As Variant

    For Each sht In ThisWorkbook.Sheets
        '解答を見て.visibleの状態を判定に追加。
        If sht.Name Like "*印刷*" And sht.Visible = xlSheetVisible Then
            print_sht = print_sht & sht.Name & ","
        End If
    Next

    print_sht = Split(Left(print_sht, Len(print_sht) - 1), ",")
    Sheets(print_sht).PrintPreview

End Sub
