Attribute VB_Name = "Module1"
Option Explicit

Sub Q31()
    Dim sht As Object
    Dim valid_list As String
    
    For Each sht In ThisWorkbook.Sheets
        valid_list = valid_list & "," & sht.Name
    Next
    
    valid_list = Mid(valid_list, 2)
    
    With ActiveSheet.Cells(1, 1).Validation
        .Delete
        .Add Type:=xlValidateList, Operator:=xlEqual, Formula1:=valid_list
        .ErrorMessage = "���͂ł��Ȃ��l�ł��B"
        .ErrorTitle = "���͋K���G���["
    End With
End Sub
