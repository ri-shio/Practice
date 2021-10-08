Attribute VB_Name = "Module1"
Option Explicit

Sub Q29()

    'ダイアログによってユーザに画像ファイルを選択させる。
    Dim path_selected As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Title = "画像の選択"
        .Filters.Clear
        .Filters.Add "画像ファイル", "*.gif;*.jpg;*.jpeg;*.png"
        If .Show = False Then
            MsgBox ("中止されました。")
            Exit Sub
        Else
            path_selected = .SelectedItems(1)
        End If
    End With
    
    'Pictures.Insertで挿入した画像にリンクが付く問題を
    'カット&ペーストで解決できる？とのこと。
    '作者環境ではリンク問題自体が再現できず。
    Dim pct As Object
    Set pct = ActiveSheet.Pictures.Insert(path_selected)
    pct.Cut
    Set pct = ActiveSheet.Pictures.Paste
    
    'アクティブセルに合わせて画像の縮尺を変え、セルの中心に画像が来るよう移動する。
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
