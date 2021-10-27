Attribute VB_Name = "Module1"
Option Explicit

Sub Q71()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")

    Dim ppApp As PowerPoint.Application
    Set ppApp = CreateObject("PowerPoint.Application")

    Dim ppPrs As PowerPoint.Presentation
    Set ppPrs = ppApp.Presentations.Open(ThisWorkbook.Path & "\prezen1.pptx")
    

    ppPrs.Slides(1).Shapes(1).Delete
    
    ws.ChartObjects("グラフ1").Chart.CopyPicture xlScreen, xlPicture
    ppPrs.Slides(1).Shapes.Paste
    
End Sub
