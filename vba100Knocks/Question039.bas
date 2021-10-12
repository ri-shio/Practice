Attribute VB_Name = "Module1"
Option Explicit

Sub Q39()
    Columns(3).Clear
    
    Range(Cells(1, 1), Cells(Rows.Count, 1).End(xlUp)).Copy _
    Destination:=Cells(1, 3)
    Range(Cells(1, 2), Cells(Rows.Count, 2).End(xlUp)).Copy _
    Destination:=Cells(1, 3).End(xlDown).Offset(1, 0)
    
    Range(Cells(1, 3), Cells(1, 3).End(xlDown)).Sort key1:=Cells(1, 3), order1:=xlAscending, Header:=xlNo
    
    Dim i As Long
    For i = Cells(1, 3).End(xlDown).Row To 2 Step -1
        If Cells(i, 3) = Cells(i - 1, 3) Then Cells(i, 3).Delete (xlShiftUp)
    Next
End Sub
