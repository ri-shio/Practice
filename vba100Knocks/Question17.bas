Attribute VB_Name = "Module1"
Option Explicit

Sub Q17()

    Sheets("���E�ۃ}�X�^").Cells.ClearContents
    
    Dim i As Integer, j As Integer
    
    For i = 1 To Sheets("�Ј�").Cells(1, 1).CurrentRegion.Rows.Count
        For j = 1 To 4
        Sheets("���E�ۃ}�X�^").Cells(i, j) = Sheets("�Ј�").Cells(i, j + 2)
        Next
    Next
    
    With Sheets("���E�ۃ}�X�^").Cells(1, 1).CurrentRegion
        .RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes
        .Sort Key1:=Sheets("���E�ۃ}�X�^").Cells(1, 1), Order1:=xlAscending, Key2:=Sheets("���E�ۃ}�X�^").Cells(1, 2), Order2:=xlAscending, Header:=xlYes
    End With
End Sub

'Sheets("�Ј�").Range(Cells(i, 3), Cells(i, 6)).Copy Destination:=Sheets("���E�ۃ}�X�^").Cells(i, 1)
