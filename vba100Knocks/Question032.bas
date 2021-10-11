Attribute VB_Name = "Module1"
Option Explicit

Sub Q32()
    
    '動的配列にて各ブックのフルパスを保持させる。
    Dim wb As Workbook
    Dim wb_num As Integer
    Dim i As Integer: i = 1
    Dim outputLog() As String
    
    wb_num = Workbooks.Count
    ReDim outputLog(1 To wb_num)
    
    For Each wb In Workbooks
        wb.Save
        outputLog(i) = wb.Path & "\" & wb.Name
        i = i + 1
    Next
    
    'ログファイル用の定数を定義。
    Dim logTitle As String
    Dim logPath As String
    
    logTitle = "log_" & Format(Now, "yyyymmddhhnnss") & ".txt"
    logPath = ThisWorkbook.Path & "\" & logTitle
    
    'ログファイルを開くまたは作成し、各配列の内容を書き出す。
    Dim fso As Object
    Dim logFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Set logFile = fso.OpenTextFile(Filename:=logPath, IOMode:=8, Create:=True)
    For i = 1 To wb_num
        Call logFile.writeline(outputLog(i))
    Next
    
    'ログファイルを閉じ、Excelを終了させる。
    logFile.Close
    Application.Quit
    
End Sub
