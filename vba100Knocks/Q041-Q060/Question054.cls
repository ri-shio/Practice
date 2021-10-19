VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Column <> 2 Or Target.Row = 1 Then Exit Sub

    Dim masCodeColumn As Range
    Set masCodeColumn = Sheets("マスタ").Range(Sheets("マスタ").Cells(2, 1), Sheets("マスタ").Cells(Sheets("マスタ").Cells(2, 1).CurrentRegion.Rows.Count, 1))
    If masCodeColumn.Find(what:=Target.Value) Is Nothing Then
        Target.Font.Color = vbRed
        Target.Offset(0, 1).Clear
        Target.Offset(0, 3).Clear

    Else
        Target.Font.ColorIndex = xlAutomatic
        Target.Offset(0, 1) = masCodeColumn.Find(what:=Target.Value).Offset(0, 1).Value
        Target.Offset(0, 3) = masCodeColumn.Find(what:=Target.Value).Offset(0, 2).Value
    End If
End Sub
