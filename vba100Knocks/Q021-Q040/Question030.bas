Attribute VB_Name = "Module1"
Option Explicit

Sub Q30()
    Application.CutCopyMode = False

    Dim i As Integer
    For i = 1 To Sheets("名簿").Cells(1, 1).CurrentRegion.Rows.Count - 1
        If i Mod 2 = 1 Then
            If i >= 3 Then
                Sheets("名札").Rows(i).RowHeight = Sheets("名札").Rows(1).RowHeight
                Sheets("名札").Rows(i + 1).RowHeight = Sheets("名札").Rows(2).RowHeight
                Sheets("名札").Range("A1:A2").Copy
                DoEvents
                Sheets("名札").Range(Cells(i, 1), Cells(i + 1, 1)).PasteSpecial Paste:=xlPasteFormats
            End If
            
            Sheets("名札").Cells(i, 1) = Sheets("名簿").Cells(i + 1, 2)
            Sheets("名札").Cells(i + 1, 1) = Sheets("名簿").Cells(i + 1, 3)
            
        Else
            If i >= 4 Then
                Sheets("名札").Range("B1:B2").Copy
                DoEvents
                Sheets("名札").Range(Cells(i - 1, 2), Cells(i, 2)).PasteSpecial Paste:=xlPasteFormats
            End If
            
            Sheets("名札").Cells(i - 1, 2) = Sheets("名簿").Cells(i + 1, 2)
            Sheets("名札").Cells(i, 2) = Sheets("名簿").Cells(i + 1, 3)
        End If
    Next
    
End Sub
