Attribute VB_Name = "Module1"
Option Explicit

Sub Q2()

Worksheets("Sheet1").Range("A1:C5").Copy
Worksheets("Sheet2").Range("A1").PasteSpecial Paste:=xlPasteValues
Worksheets("Sheet2").Range("A1").PasteSpecial Paste:=xlPasteFormats
Application.CutCopyMode = False

End Sub
