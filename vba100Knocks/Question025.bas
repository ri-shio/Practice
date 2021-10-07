Attribute VB_Name = "Module1"
Option Explicit

Sub Q25()
    Dim i As Integer, j As Integer
    Dim dbRowCnt As Integer

    For i = 2 To Sheets("ã").Cells(1, 1).CurrentRegion.Rows.Count Step 2
        Sheets("ã").Cells(i, 1).UnMerge
        Sheets("ã").Cells(i + 1, 1) = Cells(i, 1)
    Next
    
    dbRowCnt = 2
    For i = 2 To Sheets("ã").Cells(1, 1).CurrentRegion.Rows.Count
        For j = 3 To Sheets("ã").Cells(1, 1).CurrentRegion.Columns.Count
            Sheets("ãDB").Cells(dbRowCnt, 1) = Sheets("ã").Cells(i, 1)
            Sheets("ãDB").Cells(dbRowCnt, 2) = Sheets("ã").Cells(i, 2)
            Sheets("ãDB").Cells(dbRowCnt, 3) = Sheets("ã").Cells(1, j)
            Sheets("ãDB").Cells(dbRowCnt, 4) = Sheets("ã").Cells(i, j)
            dbRowCnt = dbRowCnt + 1
        Next
    Next
    
    Sheets("ãDB").Range("C:C").NumberFormatLocal = "yyyy/mm/dd"
    
End Sub
