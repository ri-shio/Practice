Attribute VB_Name = "Module1"
Option Explicit

Sub Q38()
    Dim h_date As Variant
    
    h_date = Sheets("�j��").Range(Sheets("�j��").Cells(1, 1), Sheets("�j��").Cells(Sheets("�j��").Cells(1, 1).CurrentRegion.Rows.Count, 1))
    
    Dim i As Long, j As Long
    Dim wd_cnt As Long, hd_cnt As Long
    Dim isHoliday As Boolean
    wd_cnt = 2: hd_cnt = 2
    
    For i = 2 To Sheets("����").Cells(1, 1).CurrentRegion.Rows.Count
    
        isHoliday = False
        For j = 1 To UBound(h_date)
            If Sheets("����").Cells(i, 1) = h_date(j, 1) Then
            isHoliday = True
            Exit For
            End If
        Next
        
        If Format(Sheets("����").Cells(i, 1), "aaa") = "�y" Or Format(Sheets("����").Cells(i, 1), "aaa") = "��" Then isHoliday = True
        If isHoliday = True Then
            Sheets("����").Cells(i, 1).Resize(1, 6).Copy _
            Destination:=Sheets("�y���j").Cells(hd_cnt, 1)
            hd_cnt = hd_cnt + 1
        Else
            Sheets("����").Cells(i, 1).Resize(1, 6).Copy _
            Destination:=Sheets("����").Cells(wd_cnt, 1)
            wd_cnt = wd_cnt + 1
        End If
    Next
    
End Sub
