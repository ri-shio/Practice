Attribute VB_Name = "Module1"
Option Explicit

Sub Q28()
    Dim sht As Object
    Dim sht_name() As String
    Dim separator As String: separator = "_"
    Dim i As Integer: i = 1
    
    ReDim sht_name(1 To ThisWorkbook.Sheets.Count, 1 To 2)
    
    For Each sht In ThisWorkbook.Sheets
        sht_name(i, 1) = Left(sht.Name, InStr(sht.Name, separator) - 1)
        sht_name(i, 2) = Mid(sht.Name, InStr(sht.Name, separator) + 1)
        i = i + 1
    Next
    
    Dim buf As String
    Dim checkdir() As String
    Dim dir_num As Integer
    Dim exists As Boolean: exists = False
    
    Dim new_wb As Workbook
    Dim new_wb_name As String
    Dim new_wb_path As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    For i = 1 To ThisWorkbook.Sheets.Count
        If fso.FolderExists(ThisWorkbook.Path & "\" & sht_name(i, 1)) = False Then MkDir (ThisWorkbook.Path & "\" & sht_name(i, 1))
        
        If fso.FileExists(ThisWorkbook.Path & "\" & sht_name(i, 1) & "\" & sht_name(i, 2) & ".xlsx") = False Then
            Set new_wb = Workbooks.Add
            new_wb.Save
            new_wb_name = new_wb.Name
            new_wb_path = new_wb.Path & "\" & new_wb.Name
            new_wb.Close
        
            Call fso.MoveFile(new_wb_path, ThisWorkbook.Path & "\" & sht_name(i, 1) & "\")
            fso.GetFile(ThisWorkbook.Path & "\" & sht_name(i, 1) & "\" & new_wb_name).Name = sht_name(i, 2) & ".xlsx"
        End If
    Next

End Sub
