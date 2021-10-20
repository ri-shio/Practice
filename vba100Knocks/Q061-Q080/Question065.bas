Attribute VB_Name = "Module1"
Option Explicit

Sub Q65()
    '見出し列の列数
    Const headerRowsCnt As Integer = 1
    
    Dim dRowsCnt As Long: dRowsCnt = Sheets("data").Cells(2, 1).CurrentRegion.Rows.Count - headerRowsCnt
    Dim dColsCnt As Long: dColsCnt = Sheets("data").Cells(2, 1).CurrentRegion.Columns.Count

    Dim customerData As Variant
    ReDim customerData(1 To dRowsCnt, 1 To dColsCnt)
    
    Dim i As Long, j As Long
    
    For i = 1 To dRowsCnt
        For j = 1 To dColsCnt
            Select Case Sheets("フォーマット").Cells(j + 1, 2)
                Case Is = "C"
                    Select Case Sheets("フォーマット").Cells(j + 1, 3)
                        Case Is >= Len(Sheets("data").Cells(i + headerRowsCnt, j))
                            customerData(i, j) = Sheets("data").Cells(i + headerRowsCnt, j)
                            
                        Case Is < Len(Sheets("data").Cells(i + headerRowsCnt, j))
                            customerData(i, j) = "!文字数制限を超過しています"
                            
                        End Select
                        
                Case Is = "N"
                    Select Case Sheets("フォーマット").Cells(j + 1, 3)
                        Case Is > Len(Sheets("data").Cells(i + headerRowsCnt, j))
                            customerData(i, j) = _
                            String(Sheets("フォーマット").Cells(j + 1, 3) - Len(Sheets("data").Cells(i + headerRowsCnt, j)), "0") _
                            & Sheets("data").Cells(i + headerRowsCnt, j)

                        Case Is = Len(Sheets("data").Cells(i + headerRowsCnt, j))
                            customerData(i, j) = Sheets("data").Cells(i + headerRowsCnt, j)
                            
                        Case Is < Len(Sheets("data").Cells(i + headerRowsCnt, j))
                            customerData(i, j) = "!桁数制限を超過しています"
                            
                        End Select
                        
                Case Else
                    customerData(i, j) = "!フォーマット指定が不正です"
                    
            End Select
            
        Next
    Next
    
    Open ThisWorkbook.Path & "\customerData.txt" For Output As #1
    For i = 1 To UBound(customerData, 1)
        For j = 1 To UBound(customerData, 2) - 1
            Print #1, customerData(i, j) & ",";
        Next
        
        Print #1, customerData(i, j)
    Next
    
    Close #1
    
End Sub
