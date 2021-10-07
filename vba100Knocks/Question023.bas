Attribute VB_Name = "Module1"
Option Explicit

Sub Q23()
    '判定対象のワークブックは定数として定義。
    Const wb1_name As String = "Book_20201101.xlsx"
    Const wb2_name As String = "Book_20201102.xlsx"
    
    'wb1,wb2にそれぞれ判定するワークブックをセットする。
    Dim wb1 As Workbook, wb2 As Workbook
    Dim wb1_sheets() As String, wb2_sheets() As String

    Workbooks.Open ThisWorkbook.Path & "\" & wb1_name
    Set wb1 = Workbooks(wb1_name)

    Workbooks.Open ThisWorkbook.Path & "\" & wb2_name
    Set wb2 = Workbooks(wb2_name)
    
    'wb1_sheets,wb2_sheetsの配列数を動的に決定し、シート名を取得する。
    ReDim wb1_sheets(1 To wb1.Sheets.Count), wb2_sheets(1 To wb2.Sheets.Count)
    
    Dim i As Integer
    Dim ws As Object
    
    i = 1
    For Each ws In wb1.Sheets
        wb1_sheets(i) = ws.Name
        i = i + 1
    Next
    
    i = 1
    For Each ws In wb2.Sheets
        wb2_sheets(i) = ws.Name
        i = i + 1
    Next
    
    'それぞれの配列内でシート名の順番を整える。
    Dim j As Integer
    Dim swap As String
    
    For i = 1 To wb1.Sheets.Count
        For j = wb1.Sheets.Count To i Step -1
            If wb1_sheets(i) > wb1_sheets(j) Then
                swap = wb1_sheets(i)
                wb1_sheets(i) = wb1_sheets(j)
                wb1_sheets(j) = swap
            End If
        Next
    Next
    
    For i = 1 To wb2.Sheets.Count
        For j = wb2.Sheets.Count To i Step -1
            If wb2_sheets(i) > wb2_sheets(j) Then
                swap = wb2_sheets(i)
                wb2_sheets(i) = wb2_sheets(j)
                wb2_sheets(j) = swap
            End If
        Next
    Next
    
    'シート数およびシート名を比較し、不一致が発生したらFalseとする。どちらも一致の場合はTrueが残る。
    Dim isSame As Boolean: isSame = True
    
    If wb1.Sheets.Count <> wb2.Sheets.Count Then isSame = False
    
    For i = 1 To wb1.Sheets.Count
        If wb1_sheets(i) <> wb2_sheets(i) Then isSame = False
    Next
    
    'T/Fに従いメッセージを作成し、メッセージボックスに表示。判定対象のワークブックも閉じる。
    Dim message As String
    If isSame Then
        message = "一致"
    Else
        message = "不一致"
    End If
    
    wb1.Close
    wb2.Close
    
    MsgBox (message)
    
End Sub
