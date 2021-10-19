Attribute VB_Name = "Module1"
Option Explicit

Sub Q59()
    Const bDate As Date = #4/1/2020#

    Dim wbNameM As String, wbNameX As String
    Dim wbPathM As String, wbPathX As String
    Dim wbSheets(2) As String
    Dim wb As Workbook
    Dim sht As Object
    
    Dim i As Integer, j As Integer
    Dim within As Boolean
    
    For i = 1 To 4
        wbNameM = i & "Q.xlsm": wbNameX = i & "Q.xlsx"
        wbPathM = ThisWorkbook.Path & "\" & wbNameM
        wbPathX = ThisWorkbook.Path & "\" & wbNameX
        
        If Dir(wbPathX) <> "" Then
            MsgBox "同一フォルダに四半期集計用のファイルが既に存在します。このマクロは動作を終了します。"
            Exit Sub
        End If
        
        wbSheets(0) = Format(DateAdd("m", i * 3 - 3, bDate), "yyyy年mm月")
        wbSheets(1) = Format(DateAdd("m", i * 3 - 2, bDate), "yyyy年mm月")
        wbSheets(2) = Format(DateAdd("m", i * 3 - 1, bDate), "yyyy年mm月")
        
        ThisWorkbook.SaveCopyAs (wbPathM)
        Set wb = Workbooks.Open(wbPathM)
        
        Application.DisplayAlerts = False
        
        For Each sht In wb.Sheets
        
            within = False
            
            For j = 0 To 2
                If sht.Name = wbSheets(j) Then within = True
            Next
            
            If within = False Then sht.Delete
            
        Next
        
        wb.SaveAs Filename:=wbPathX, FileFormat:=xlOpenXMLWorkbook
        wb.Close
        Kill wbPathM
        
        Application.DisplayAlerts = True
        
    Next
    
End Sub
