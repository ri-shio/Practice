Attribute VB_Name = "Module1"
Option Explicit
Public wsDF As Worksheet
Public wsDP As Worksheet

Sub Q67()
    Set wsDF = Worksheets.Add
    Set wsDP = Worksheets.Add
    wsDF.Move After:=Sheets(Sheets.Count)
    wsDP.Move After:=Sheets(Sheets.Count)
    
    Dim customerList As Variant
    customerList = Sheets("リスト").Cells(1, 1).CurrentRegion
    
    wsDF.Cells(1, 1).Resize(UBound(customerList, 1), UBound(customerList, 2)) = customerList
    
    UserForm1.Show
    
    Application.DisplayAlerts = False
    
    wsDF.Delete
    wsDP.Delete
    
    Application.DisplayAlerts = True
    
End Sub

