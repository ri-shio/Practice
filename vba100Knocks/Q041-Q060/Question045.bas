Attribute VB_Name = "Module1"
Option Explicit

Sub Q45()
    Dim tbl As Object
    Set tbl = ActiveSheet.ListObjects(1)
    
    tbl.ListColumns.Add Position:=4
    tbl.HeaderRowRange(4) = "çáåvóÒ1"
    tbl.Range.Cells(2, 4).FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"
    
    tbl.ListColumns.Add Position:=7
    tbl.HeaderRowRange(7) = "çáåvóÒ2"
    tbl.Range.Cells(2, 7).FormulaR1C1 = "=SUM(RC[-2]:RC[-1])"
    
End Sub
