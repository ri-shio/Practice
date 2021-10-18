Attribute VB_Name = "Module1"
Option Explicit

Sub Q57()

    'FileSystemObjectの利用準備
    Dim fso As FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    'fileArrayにファイルのフルパスと最終更新日時を格納していく。
    Dim dirPath As String
    Dim fileArray() As Variant
    dirPath = ThisWorkbook.Path & "\BUCKUP\"

    Dim buFile As File
    Dim i As Long: i = 0
    ReDim fileArray(1 To fso.GetFolder(dirPath).Files.Count, 1)

    For Each buFile In fso.GetFolder(dirPath).Files
        i = i + 1
        fileArray(i, 0) = buFile.Path
        fileArray(i, 1) = buFile.DateLastModified

    Next

    '最終更新日時をバブルソートで降順に並び変える。その際、フルパスも同時に並び替える。
    Dim j As Long
    Dim swap As Variant

    For i = 1 To fso.GetFolder(dirPath).Files.Count
        For j = fso.GetFolder(dirPath).Files.Count To i Step -1
            If fileArray(i, 1) <= fileArray(j, 1) Then
                swap = Array(fileArray(i, 0), fileArray(i, 1))
                fileArray(i, 0) = fileArray(j, 0)
                fileArray(i, 1) = fileArray(j, 1)
                fileArray(j, 0) = swap(0)
                fileArray(j, 1) = swap(1)
            End If
        Next
    Next

    '最終更新日時が降順で並んでいるため、最終要素から見ていくと古いファイルから順に見ていくこととなる。
    'yyyymmddで一つ手前のファイルと比べ、一致した場合は自分がより古いファイルであるといえる。
    '今回は日付に加え拡張子も一致した場合、同一のバックアップファイルであるとみなし、
    '現在対象となっているファイルを削除する。
    'なお、最も先頭に来ているファイルは最も新しいファイルであるため削除するケースは存在せず、
    'よってi = 2までを見ればよい。

    For i = fso.GetFolder(dirPath).Files.Count To 2 Step -1
        If Format(fileArray(i, 1), "yyyymmdd") = Format(fileArray(i - 1, 1), "yyyymmdd") Then
            If Mid(fileArray(i, 0), InStrRev(fileArray(i, 0), ".")) = Mid(fileArray(i - 1, 0), InStrRev(fileArray(i - 1, 0), ".")) Then
                fso.GetFile(fileArray(i, 0)).Delete
            End If
        End If
    Next
End Sub
