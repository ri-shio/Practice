Attribute VB_Name = "Module1"
Option Explicit

Sub Q32()
    
    '���I�z��ɂĊe�u�b�N�̃t���p�X��ێ�������B
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
    
    '���O�t�@�C���p�̒萔���`�B
    Dim logTitle As String
    Dim logPath As String
    
    logTitle = "log_" & Format(Now, "yyyymmddhhnnss") & ".txt"
    logPath = ThisWorkbook.Path & "\" & logTitle
    
    '���O�t�@�C�����J���܂��͍쐬���A�e�z��̓��e�������o���B
    Dim fso As Object
    Dim logFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Set logFile = fso.OpenTextFile(Filename:=logPath, IOMode:=8, Create:=True)
    For i = 1 To wb_num
        Call logFile.writeline(outputLog(i))
    Next
    
    '���O�t�@�C������AExcel���I��������B
    logFile.Close
    Application.Quit
    
End Sub
