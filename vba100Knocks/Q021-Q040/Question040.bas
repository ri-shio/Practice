Attribute VB_Name = "Module1"
Option Explicit

Sub Q40()
    Dim filePath As String
    Dim fileName As String
    
    filePath = ThisWorkbook.Path & "\data\"
    fileName = Dir(filePath & "*.xls*")
    
    Dim wb As Workbook, sht As Object
    Dim sht_exists As Boolean
    
    Do Until fileName = ""
        sht_exists = False
        Set wb = Workbooks.Open(filePath & fileName)
            For Each sht In wb.Sheets
                If sht.Name = "2020”N12ŒŽ" Then
                    sht.Range(sht.Cells(2, 1), sht.Cells(sht.Cells(2, 1).CurrentRegion.Rows.Count, sht.Cells(2, 1).CurrentRegion.Columns.Count)).Copy _
                    Destination:=ThisWorkbook.Sheets("2020”N12ŒŽ").Cells(ThisWorkbook.Sheets("2020”N12ŒŽ").Cells(1, 1).CurrentRegion.Rows.Count, 1).Offset(1, 0)
                    Exit For
                End If
            Next
        wb.Close
        fileName = Dir()
    Loop
End Sub
