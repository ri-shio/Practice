Attribute VB_Name = "Module1"
Option Explicit

Sub Q35()
    Dim ws As Object
    Set ws = Sheets("èåèïtÇ´èëéÆ")
    
    ws.Cells.FormatConditions.Delete
    
    Dim col_num() As Variant
    col_num = Array(5, 7)
    
    Dim i As Integer
    
    For i = 0 To 1
        
        Columns(col_num(i)).FormatConditions.Add(xlBlanksCondition).Font.ColorIndex = xlAutomatic
        Columns(col_num(i)).FormatConditions.Add(xlCellValue, xlLess, 0.9).Interior.Color = vbRed
        Columns(col_num(i)).FormatConditions.Add(xlCellValue, xlLess, 1).Font.Color = vbRed
        
    Next
End Sub
