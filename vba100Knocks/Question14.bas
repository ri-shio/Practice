Attribute VB_Name = "Module1"
Option Explicit

Sub Q14()
    Dim sht As Object
    
    For Each sht In sheets
        sht.Visible = xlSheetVisible    '‰ð“š‚ðŒ©‚Ä’Ç‹L
        sht.Cells.Copy
        sht.Cells(1, 1).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        sht.Activate
        sht.Cells(1, 1).Select
    Next
    
    Application.DisplayAlerts = False
    
    For Each sht In sheets
        If sht.Name Like "*ŽÐŠO”é*" Then sht.Delete
    Next
    
    Application.DisplayAlerts = True
    
End Sub
