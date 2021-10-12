Attribute VB_Name = "Module1"
Option Explicit

Sub Q36()
    Dim col_buffer() As Variant
    Dim col_count As Long: col_count = Cells(1, 1).CurrentRegion.Columns.Count
    Dim i As Long, j As Long
    
    ReDim col_buffer(1 To col_count, 1)

    For i = 1 To col_count
        col_buffer(i, 0) = Columns(i)
        col_buffer(i, 1) = Mid(Cells(1, i), InStrRev(Cells(1, i), "(") + 1, InStrRev(Cells(1, i), ")") - InStrRev(Cells(1, i), "(") - 1)
    Next
    
    Dim swap As Variant
    
    For i = 1 To col_count
        For j = col_count To i Step -1
            If CLng(col_buffer(i, 1)) > CLng(col_buffer(j, 1)) Then
                swap = Array(col_buffer(i, 0), col_buffer(i, 1))
                col_buffer(i, 0) = col_buffer(j, 0)
                col_buffer(i, 1) = col_buffer(j, 1)
                col_buffer(j, 0) = swap(0)
                col_buffer(j, 1) = swap(1)
            End If
        Next
    Next
    
    For i = 1 To col_count
        Columns(i) = col_buffer(i, 0)
    Next
End Sub
