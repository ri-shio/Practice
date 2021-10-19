Attribute VB_Name = "Module1"
Option Explicit

Sub Q43()
    '後にExcel本体を終了させるため、まず保存をする。
    ThisWorkbook.Save
    
    'アクティブシートをデータシートとし、整形用に新しくワークシートを作成する。
    Dim data_ws As Worksheet
    Dim csvout_ws As Worksheet
    
    Set data_ws = ActiveSheet
    Set csvout_ws = Worksheets.Add
    
    data_ws.Cells(1, 1).CurrentRegion.Copy _
    Destination:=csvout_ws.Cells(1, 1)
    
    '要件に従い、A列をyyyy-mm-dd形式に、B列をカンマ無し整数、C列をカンマ無し小数2桁とする。
    csvout_ws.Columns(1).NumberFormatLocal = "yyyy-mm-dd"
    csvout_ws.Columns(2).NumberFormatLocal = "0"
    csvout_ws.Columns(3).NumberFormatLocal = "0.00"
    
    '要件に従い、D列内のダブルクォーテーションをエスケープできるようにする。
    Dim i As Long
    
    For i = 1 To csvout_ws.Cells(1, 4).CurrentRegion.Count
        If InStr(csvout_ws.Cells(i, 4), """") > 0 Then
            csvout_ws.Cells(i, 4) = "" & Replace(csvout_ws.Cells(i, 4), "", "" & "") & ""
        End If
    Next
    
    'メッセージ確認をFalseにし、CSVエクスポート・整形用シート削除、Excel本体の終了を行う。
    Application.DisplayAlerts = False

    ThisWorkbook.SaveAs _
    Filename:=ThisWorkbook.Path & "\CSVOutput.csv", _
    FileFormat:=xlCSV, _
    local:=True

    csvout_ws.Delete
    Application.Quit
    
End Sub
