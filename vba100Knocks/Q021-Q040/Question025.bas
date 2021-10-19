Attribute VB_Name = "Module1"
Option Explicit

Sub Q25()
    Dim i As Integer, j As Integer
    Dim dbRowCnt As Integer

    For i = 2 To Sheets("îÑè„").Cells(1, 1).CurrentRegion.Rows.Count Step 2
        Sheets("îÑè„").Cells(i, 1).UnMerge
        Sheets("îÑè„").Cells(i + 1, 1) = Cells(i, 1)
    Next
    
    dbRowCnt = 2
    For i = 2 To Sheets("îÑè„").Cells(1, 1).CurrentRegion.Rows.Count
        For j = 3 To Sheets("îÑè„").Cells(1, 1).CurrentRegion.Columns.Count
            Sheets("îÑè„DB").Cells(dbRowCnt, 1) = Sheets("îÑè„").Cells(i, 1)
            Sheets("îÑè„DB").Cells(dbRowCnt, 2) = Sheets("îÑè„").Cells(i, 2)
            Sheets("îÑè„DB").Cells(dbRowCnt, 3) = Sheets("îÑè„").Cells(1, j)
            Sheets("îÑè„DB").Cells(dbRowCnt, 4) = Sheets("îÑè„").Cells(i, j)
            dbRowCnt = dbRowCnt + 1
        Next
    Next
    
    Sheets("îÑè„DB").Range("C:C").NumberFormatLocal = "yyyy/mm/dd"
    
End Sub
