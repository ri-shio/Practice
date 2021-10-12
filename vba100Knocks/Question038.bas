Attribute VB_Name = "Module1"
Option Explicit

Sub Q38()
    Dim h_date As Variant
    
    h_date = Sheets("j“ú").Range(Sheets("j“ú").Cells(1, 1), Sheets("j“ú").Cells(Sheets("j“ú").Cells(1, 1).CurrentRegion.Rows.Count, 1))
    
    Dim i As Long, j As Long
    Dim wd_cnt As Long, hd_cnt As Long
    Dim isHoliday As Boolean
    wd_cnt = 2: hd_cnt = 2
    
    For i = 2 To Sheets("”„ã").Cells(1, 1).CurrentRegion.Rows.Count
    
        isHoliday = False
        For j = 1 To UBound(h_date)
            If Sheets("”„ã").Cells(i, 1) = h_date(j, 1) Then
            isHoliday = True
            Exit For
            End If
        Next
        
        If Format(Sheets("”„ã").Cells(i, 1), "aaa") = "“y" Or Format(Sheets("”„ã").Cells(i, 1), "aaa") = "“ú" Then isHoliday = True
        If isHoliday = True Then
            Sheets("”„ã").Cells(i, 1).Resize(1, 6).Copy _
            Destination:=Sheets("“y“új").Cells(hd_cnt, 1)
            hd_cnt = hd_cnt + 1
        Else
            Sheets("”„ã").Cells(i, 1).Resize(1, 6).Copy _
            Destination:=Sheets("•½“ú").Cells(wd_cnt, 1)
            wd_cnt = wd_cnt + 1
        End If
    Next
    
End Sub
