Attribute VB_Name = "Module1"
Option Explicit

Sub Q30()
    Application.CutCopyMode = False

    Dim i As Integer
    For i = 1 To Sheets("–¼•ë").Cells(1, 1).CurrentRegion.Rows.Count - 1
        If i Mod 2 = 1 Then
            If i >= 3 Then
                Sheets("–¼ŽD").Rows(i).RowHeight = Sheets("–¼ŽD").Rows(1).RowHeight
                Sheets("–¼ŽD").Rows(i + 1).RowHeight = Sheets("–¼ŽD").Rows(2).RowHeight
                Sheets("–¼ŽD").Range("A1:A2").Copy
                DoEvents
                Sheets("–¼ŽD").Range(Cells(i, 1), Cells(i + 1, 1)).PasteSpecial Paste:=xlPasteFormats
            End If
            
            Sheets("–¼ŽD").Cells(i, 1) = Sheets("–¼•ë").Cells(i + 1, 2)
            Sheets("–¼ŽD").Cells(i + 1, 1) = Sheets("–¼•ë").Cells(i + 1, 3)
            
        Else
            If i >= 4 Then
                Sheets("–¼ŽD").Range("B1:B2").Copy
                DoEvents
                Sheets("–¼ŽD").Range(Cells(i - 1, 2), Cells(i, 2)).PasteSpecial Paste:=xlPasteFormats
            End If
            
            Sheets("–¼ŽD").Cells(i - 1, 2) = Sheets("–¼•ë").Cells(i + 1, 2)
            Sheets("–¼ŽD").Cells(i, 2) = Sheets("–¼•ë").Cells(i + 1, 3)
        End If
    Next
    
End Sub
