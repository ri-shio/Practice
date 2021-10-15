Attribute VB_Name = "Module1"
Option Explicit

Sub Q47()

    Dim wdw As Window
    Dim wb As Workbook
    Dim sht As Object
    
    Set wb = ThisWorkbook
    For Each wdw In wb.Windows
                
        wdw.Zoom = 85
        wdw.DisplayGridlines = False
        wdw.View = xlNormalView

    Next
    
    For Each sht In wb.Sheets
            
        sht.Rows(1).Hidden = False
        sht.Columns(1).Hidden = False
        sht.Cells(1, 1).Select
        sht.PageSetup.Orientation = xlLandscape

    Next
    
End Sub
