Attribute VB_Name = "Module1"
Option Explicit

Sub Q3()

Sheets("Sheet1").Range(Cells(2, 2), Cells(Cells(2, 2).CurrentRegion.Rows.Count, Cells(2, 2).CurrentRegion.Columns.Count)).ClearContents

End Sub
