Attribute VB_Name = "Module1"
Option Explicit
Sub Q1()

Worksheets("Sheet1").Range("A1:C5").Copy Destination:=Worksheets("Sheet2").Range("A1")

'解答を見て追記：コピーモードを解除
Application.CutCopyMode = False

End Sub
