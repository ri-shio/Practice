Attribute VB_Name = "Module1"
Option Explicit

Sub Q15()

    Dim sht_name() As Variant
    Dim sht_count As Integer
    Dim i As Integer
    
    sht_count = Sheets.Count
    ReDim sht_name(1 To sht_count)
    
    For i = 1 To sht_count
        sht_name(i) = Sheets(i).Name
    Next
    
    Dim swap As String
    Dim j As Integer
    
    For i = 1 To sht_count
        For j = sht_count To i Step -1
            If sht_name(i) > sht_name(j) Then
                swap = sht_name(i)
                sht_name(i) = sht_name(j)
                sht_name(j) = swap
            End If
        Next
    Next
    
    For i = 2 To sht_count
        Sheets(sht_name(i)).Move after:=Sheets(sht_name(i - 1))
    Next
    
End Sub


