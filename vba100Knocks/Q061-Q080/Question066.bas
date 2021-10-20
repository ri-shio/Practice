Attribute VB_Name = "Module1"
Option Explicit

Sub Q66()
    ActiveSheet.Cells.Clear
    Cells(1, 1) = "フルパス": Cells(1, 2) = "更新日時": Cells(1, 3) = "ファイルサイズ"
    
    Dim hitCnt As Integer: hitCnt = 0
    
    Call FileCheck(ThisWorkbook.path, hitCnt)
    
End Sub

Sub FileCheck(ByVal cPath As String, ByRef hitCnt As Integer)

    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    
    Dim tFile As File
    
    If cPath <> ThisWorkbook.path Then
        If fso.FileExists(cPath & "\" & ThisWorkbook.Name) Then
            hitCnt = hitCnt + 1
            Set tFile = fso.GetFile(cPath & "\" & ThisWorkbook.Name)
            Cells(hitCnt + 1, 1) = tFile.path
            Cells(hitCnt + 1, 2) = tFile.DateLastModified
            Cells(hitCnt + 1, 3) = tFile.Size
            
        End If
    End If
    
    Dim cDir As Object
    
    For Each cDir In fso.GetFolder(cPath).SubFolders

        Call FileCheck(cDir.path, hitCnt)
    Next
    
End Sub
