Attribute VB_Name = "Module1"
Option Explicit

Sub Q30()
    Application.CutCopyMode = False

    Dim i As Integer
    For i = 1 To Sheets("����").Cells(1, 1).CurrentRegion.Rows.Count - 1
        If i Mod 2 = 1 Then
            If i >= 3 Then
                Sheets("���D").Rows(i).RowHeight = Sheets("���D").Rows(1).RowHeight
                Sheets("���D").Rows(i + 1).RowHeight = Sheets("���D").Rows(2).RowHeight
                Sheets("���D").Range("A1:A2").Copy
                DoEvents
                Sheets("���D").Range(Cells(i, 1), Cells(i + 1, 1)).PasteSpecial Paste:=xlPasteFormats
            End If
            
            Sheets("���D").Cells(i, 1) = Sheets("����").Cells(i + 1, 2)
            Sheets("���D").Cells(i + 1, 1) = Sheets("����").Cells(i + 1, 3)
            
        Else
            If i >= 4 Then
                Sheets("���D").Range("B1:B2").Copy
                DoEvents
                Sheets("���D").Range(Cells(i - 1, 2), Cells(i, 2)).PasteSpecial Paste:=xlPasteFormats
            End If
            
            Sheets("���D").Cells(i - 1, 2) = Sheets("����").Cells(i + 1, 2)
            Sheets("���D").Cells(i, 2) = Sheets("����").Cells(i + 1, 3)
        End If
    Next
    
End Sub
