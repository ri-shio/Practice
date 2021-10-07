Attribute VB_Name = "Module1"
Option Explicit

Sub Q26()
    Cells.Clear
    
    Dim path_selected As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = "ファイル一覧を書き出すフォルダの選択"
        If .Show = False Then
            MsgBox ("中止されました。")
            Exit Sub
        Else
            path_selected = .SelectedItems(1)
        End If
    End With
    
    Cells(1, 1) = "ファイル一覧"
    Cells(1, 2) = "更新日時"
    Cells(1, 3) = "サイズ"
    
    Dim fileName As String
    Dim fso As Object
    Dim fileObj As Object
    Dim i As Integer: i = 2
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    fileName = Dir(path_selected & "\*.*", vbNormal)
    
    Do Until fileName = ""
        Set fileObj = fso.GetFile(path_selected & "\" & fileName)
        Cells(i, 1) = fileName
        If fileObj.Type Like "*Excel*" Then
            ActiveSheet.Hyperlinks.Add Anchor:=Cells(i, 1), Address:=fileObj.Path
        End If
        Cells(i, 2) = fileObj.DateLastModified
        Cells(i, 3) = fileObj.Size
        
        i = i + 1
        fileName = Dir()
    Loop
End Sub

