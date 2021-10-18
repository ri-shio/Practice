Attribute VB_Name = "Module1"
Option Explicit

Sub Q56()

'記号を含むシート名の場合、手元の環境では題意を満たせず。
'また、上記のケースで模範解答をそのまま実行しても同様に設問通りに動作せず。
'なお、記号を含まないシート名では正常に動作する。

    Dim wsName As String
    Dim wsNameSym As String

    wsName = ActiveSheet.Name & "!"
    wsNameSym = "'" & wsName & "'!"

    Dim sht As Object

    For Each sht In ThisWorkbook.Sheets
        sht.UsedRange.Replace What:=wsName, Replacement:="", LookAt:=xlPart, MatchCase:=True
        sht.UsedRange.Replace What:=wsNameSym, Replacement:="", LookAt:=xlPart, MatchCase:=True
    Next

End Sub
