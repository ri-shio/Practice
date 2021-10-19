Attribute VB_Name = "Module1"
Option Explicit

Sub Q29()

    '�_�C�A���O�ɂ���ă��[�U�ɉ摜�t�@�C����I��������B
    Dim path_selected As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Title = "�摜�̑I��"
        .Filters.Clear
        .Filters.Add "�摜�t�@�C��", "*.gif;*.jpg;*.jpeg;*.png"
        If .Show = False Then
            MsgBox ("���~����܂����B")
            Exit Sub
        Else
            path_selected = .SelectedItems(1)
        End If
    End With
    
    'Pictures.Insert�ő}�������摜�Ƀ����N���t������
    '�J�b�g&�y�[�X�g�ŉ����ł���H�Ƃ̂��ƁB
    '��Ҋ��ł̓����N��莩�̂��Č��ł����B
    Dim pct As Object
    Set pct = ActiveSheet.Pictures.Insert(path_selected)
    pct.Cut
    Set pct = ActiveSheet.Pictures.Paste
    
    '�A�N�e�B�u�Z���ɍ��킹�ĉ摜�̏k�ڂ�ς��A�Z���̒��S�ɉ摜������悤�ړ�����B
    With pct
        If .Width / ActiveCell.Width >= .Height / ActiveCell.Height Then
            .Width = ActiveCell.Width
            .Top = ActiveCell.Top + (ActiveCell.Height - .Height) / 2
            .Left = ActiveCell.Left
        Else
            .Height = ActiveCell.Height
            .Top = ActiveCell.Top
            .Left = ActiveCell.Left + (ActiveCell.Width - .Width) / 2
        End If
    End With
    
End Sub
