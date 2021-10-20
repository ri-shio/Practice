Attribute VB_Name = "Module1"
Option Explicit

Sub Q63()
    Dim dataArr As Variant
    ReDim dataArr(1 To ThisWorkbook.Sheets.Count)
    
    Dim i As Long: i = 0
    Dim sht As Object
    
    For Each sht In ThisWorkbook.Sheets
        i = i + 1
        dataArr(i) = Range(sht.Cells(2, 1), sht.Cells(sht.Cells(2, 1).CurrentRegion.Rows.Count, sht.Cells(2, 1).CurrentRegion.Columns.Count))
    Next
    
    Dim newWS As Worksheet
    Set newWS = Worksheets.Add(before:=Sheets(1))
    Sheets(2).Rows(1).Copy Destination:=newWS.Rows(1)
    
    For i = 1 To UBound(dataArr, 1)
        newWS.Cells(Cells(1, 1).CurrentRegion.Rows.Count, 1).Offset(1, 0).Resize(UBound(dataArr(i), 1), UBound(dataArr(i), 2)) = dataArr(i)
    Next
    
End Sub
