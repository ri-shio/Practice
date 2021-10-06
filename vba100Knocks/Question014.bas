Attribute VB_Name = "Module1"
Option Explicit

Sub Q14()
    Dim sht As Object
    
    For Each sht In sheets
        sht.Visible = xlSheetVisible    '解答を見て追記
        sht.Cells.Copy
        sht.Cells(1, 1).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        sht.Activate
        sht.Cells(1, 1).Select
    Next
    
    Application.DisplayAlerts = False
    
    For Each sht In sheets
        If sht.Name Like "*社外秘*" Then sht.Delete
    Next
    
    Application.DisplayAlerts = True
    
End Sub
