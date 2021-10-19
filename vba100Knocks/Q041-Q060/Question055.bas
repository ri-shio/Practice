Attribute VB_Name = "Module1"
Option Explicit

Sub Q55()
    '下記プロパティ2つは解答を見て追記。
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    MsgBox (Application.Run("'" & ThisWorkbook.Path & "\test.xlsm'!mult", 3, 4))

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
