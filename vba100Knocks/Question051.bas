Attribute VB_Name = "Module1"
Option Explicit

Sub Q51()
    Dim ws As Worksheet
    Dim indexWs As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "目次" Then
            Set indexWs = ws
            Exit For
        End If
    Next

    If indexWs Is Nothing Then
        Set indexWs = Worksheets.Add
        indexWs.Name = "目次"
    End If

    indexWs.Move Before:=Sheets(1)
    indexWs.Cells.Clear
    indexWs.Cells(1, 1) = "シート名": indexWs.Cells(1, 2) = "印刷ページ数"

    Dim sht_cnt As Long: sht_cnt = 0
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "目次" Then
            sht_cnt = sht_cnt + 1
            indexWs.Cells(1 + sht_cnt, 1) = ws.Name
            If ws.Visible = xlSheetVisible Then

                'HyperlinksのSubAddressの指定は解答を見て作成。
                indexWs.Hyperlinks.Add anchor:=indexWs.Cells(1 + sht_cnt, 1), _
                Address:="", SubAddress:="'" & Replace(ws.Name, "'", "''") & "'!A1"

                indexWs.Cells(1 + sht_cnt, 2) = ws.PageSetup.Pages.Count

            End If
        End If
    Next

End Sub
