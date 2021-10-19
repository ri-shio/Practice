Attribute VB_Name = "Module1"
Option Explicit

Sub Q25()
    Dim i As Integer, j As Integer
    Dim dbRowCnt As Integer

    For i = 2 To Sheets("売上").Cells(1, 1).CurrentRegion.Rows.Count Step 2
        Sheets("売上").Cells(i, 1).UnMerge
        Sheets("売上").Cells(i + 1, 1) = Cells(i, 1)
    Next
    
    dbRowCnt = 2
    For i = 2 To Sheets("売上").Cells(1, 1).CurrentRegion.Rows.Count
        For j = 3 To Sheets("売上").Cells(1, 1).CurrentRegion.Columns.Count
            Sheets("売上DB").Cells(dbRowCnt, 1) = Sheets("売上").Cells(i, 1)
            Sheets("売上DB").Cells(dbRowCnt, 2) = Sheets("売上").Cells(i, 2)
            Sheets("売上DB").Cells(dbRowCnt, 3) = Sheets("売上").Cells(1, j)
            Sheets("売上DB").Cells(dbRowCnt, 4) = Sheets("売上").Cells(i, j)
            dbRowCnt = dbRowCnt + 1
        Next
    Next
    
    Sheets("売上DB").Range("C:C").NumberFormatLocal = "yyyy/mm/dd"
    
End Sub
