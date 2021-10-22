VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Workbookのイベントプロシージャとして利用する。

Option Explicit

Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)

    If Target.MergeCells = False Then Exit Sub
    
    Dim bufRng As Range
    Dim bufValue As String
    
    Set bufRng = Target
    bufValue = Target.Cells(1).Value
    
    Target.HorizontalAlignment = xlLeft
    Target.MergeCells = False
    bufRng.Value = bufValue

    Cancel = True
End Sub
