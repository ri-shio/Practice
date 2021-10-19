Attribute VB_Name = "Module1"
Option Explicit

Sub Q26()
    Cells.Clear
    
    Dim path_selected As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = "�t�@�C���ꗗ�������o���t�H���_�̑I��"
        If .Show = False Then
            MsgBox ("���~����܂����B")
            Exit Sub
        Else
            path_selected = .SelectedItems(1)
        End If
    End With
    
    Cells(1, 1) = "�t�@�C���ꗗ"
    Cells(1, 2) = "�X�V����"
    Cells(1, 3) = "�T�C�Y"
    
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

