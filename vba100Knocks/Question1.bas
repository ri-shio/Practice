Attribute VB_Name = "Module1"
Option Explicit
Sub Q1()

Worksheets("Sheet1").Range("A1:C5").Copy Destination:=Worksheets("Sheet2").Range("A1")

'�𓚂����ĒǋL
Application.CutCopyMode = False

End Sub


