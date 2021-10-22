VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Workbookのイベントプロシージャとして利用する。
'なお、Application.Waitでは時刻カウント中のセル入力でエラーが出てしまうため、
'Windows APIのSleep関数を利用した。

Option Explicit
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Private Sub Workbook_Open()

    '冒頭にメッセージボックスを挟むことにより、ワークブックを正常に開けるようにする。
    MsgBox ("1分間ステータスバーに時刻を表示します。")

    Dim i As Integer
    
    For i = 1 To 60
    
        DoEvents
        
        Application.StatusBar = Format(Now, "yyyy/mm/dd hh:nn:ss")
        Sleep 1000
    
    Next
    
    Application.StatusBar = False

End Sub
