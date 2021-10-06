Attribute VB_Name = "Module1"
Option Explicit

Sub Q17()

    Sheets("部・課マスタ").Cells.ClearContents
    
    Dim i As Integer, j As Integer
    
    For i = 1 To Sheets("社員").Cells(1, 1).CurrentRegion.Rows.Count
        For j = 1 To 4
        Sheets("部・課マスタ").Cells(i, j) = Sheets("社員").Cells(i, j + 2)
        Next
    Next
    
    With Sheets("部・課マスタ").Cells(1, 1).CurrentRegion
        .RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes
        .Sort Key1:=Sheets("部・課マスタ").Cells(1, 1), Order1:=xlAscending, Key2:=Sheets("部・課マスタ").Cells(1, 2), Order2:=xlAscending, Header:=xlYes
    End With
End Sub

'Sheets("社員").Range(Cells(i, 3), Cells(i, 6)).Copy Destination:=Sheets("部・課マスタ").Cells(i, 1)
