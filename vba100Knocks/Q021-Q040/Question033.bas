Attribute VB_Name = "Module2"
Option Explicit

Sub Q33()
    '二分探索を利用するため、マスタコードで昇順に並べ替える。
    Sheets("マスタ").Cells(1, 1).Sort key1:=Sheets("マスタ").Cells(1, 1), order1:=xlAscending, Header:=xlYes

    'データ・マスタの列数を取得し、また、値をそれぞれ配列に格納する。
    Dim masterTable As Variant
    Dim masterRowNum As Long
    Dim dataTable As Variant
    Dim dataRowNum As Long
    
    masterRowNum = Sheets("マスタ").Cells(1, 1).CurrentRegion.Rows.Count
    Sheets("マスタ").Activate
    masterTable = Sheets("マスタ").Range(Cells(2, 1), Cells(masterRowNum, 3)).Value
    
    dataRowNum = Sheets("データ").Cells(1, 1).CurrentRegion.Rows.Count
    Sheets("データ").Activate
    dataTable = Sheets("データ").Range(Cells(2, 1), Cells(dataRowNum, 3)).Value
    
    'マスタの探索結果を格納するための配列を定義する。
    Dim dataTable2 As Variant
    ReDim dataTable2(1 To dataRowNum, 1 To 2)
    
    '二分探索により高速に探索を行う。
    Dim lNum As Long, hNum As Long
    Dim i As Long
    
    For i = 1 To dataRowNum - 1
        lNum = 1: hNum = masterRowNum - 1
        Do
            If dataTable(i, 2) = masterTable(Int((lNum + hNum) / 2), 1) Then
                dataTable2(i, 1) = masterTable(Int((lNum + hNum) / 2), 2)
                dataTable2(i, 2) = masterTable(Int((lNum + hNum) / 2), 3)
                Exit Do
            ElseIf dataTable(i, 2) > masterTable(Int((lNum + hNum) / 2), 1) Then
                lNum = Int((lNum + hNum) / 2)
            ElseIf dataTable(i, 2) < masterTable(Int((lNum + hNum) / 2), 1) Then
                hNum = Int((lNum + hNum) / 2)
            End If
        Loop
    Next
    
    '配列から値を取り出し、セルに入力する。
    Sheets("データ").Activate
    Sheets("データ").Range(Cells(2, 4), Cells(masterRowNum, 5)).Value = dataTable2
    
    'データシートのF列には計算式を入力する。
    With Sheets("データ").Cells(2, 6)
        .FormulaR1C1 = "=RC[-3]*RC[-1]"
        .AutoFill Destination:=Range(Cells(2, 6), Cells(dataRowNum, 6))
    End With
    
    'データシートのA1を選択し、マクロは終了。
    Sheets("データ").Cells(1, 1).Activate
End Sub
