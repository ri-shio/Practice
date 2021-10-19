Attribute VB_Name = "Module1"
Option Explicit

Sub Q20()
    Dim master_name As String: master_name = ThisWorkbook.Name
    Dim master_path As String: master_path = ThisWorkbook.Path
    Dim wb As Workbook
    Dim book_name As String
    
    Dim teststr As String
    teststr = Dir(ThisWorkbook.Path & "\BUCKUP", vbDirectory)

    If Dir(master_path & "\BUCKUP", vbDirectory) = "" Then
        MkDir master_path & "\BUCKUP"
    End If
    
    Set wb = ThisWorkbook
    book_name = Left(master_name, InStrRev(master_name, ".") - 1)
    'book_name = Left(master_name, InStr(Len(master_name) - 5, master_name, ".xls") - 1)
    
    wb.Save
    wb.SaveAs Filename:=master_path & "\BUCKUP\" & book_name & "_" & Format(Now, "yyyymmddhhnn") & ".xlsm"
    Workbooks.Open master_path & "\" & master_name
    wb.Close

End Sub
