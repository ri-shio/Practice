Attribute VB_Name = "Module1"
Option Explicit

Sub Q46()
    Dim rng As Range
    Dim i As Long
    For i = 1 To Cells(1, 1).CurrentRegion.Columns.Count
        Set rng = Cells(1, i)

        On Error Resume Next
        rng.Name.Delete
        rng.Name = Rename(rng.Value)

        If rng.Name = "" Then
            Debug.Print ("名前の定義に使用できないか、セルが空白です：" & Cells(1, i).Value)
        End If
    Next
End Sub

Function Rename(ByVal cellValue As String) As String
    Dim ngCharList As Variant
    ngCharList = Array(" ", "(", ")", "!", """", "#", "$", "%", "&", "'", "=", "~", "\", "|", _
    "@", "`", "[", "{", "「", ";", "+", ":", "*", "]", "}", "」", ",", "<", ">", "/", "?")

    Dim j As Long
    For j = 0 To UBound(ngCharList)
        If InStr(cellValue, StrConv(ngCharList(j), vbNarrow)) > 0 Then
            cellValue = Replace(cellValue, StrConv(ngCharList(j), vbNarrow), "_")
        ElseIf InStr(cellValue, StrConv(ngCharList(j), vbWide)) > 0 Then
            cellValue = Replace(cellValue, StrConv(ngCharList(j), vbWide), "_")
        End If
    Next

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")

    Dim c_digit As String
    c_digit = "[" & ChrW(12881) & "-" & ChrW(12910) & "]"
    With re
        .Global = True

        .Pattern = ChrW(9450)
        cellValue = .Replace(cellValue, 0)

        .Pattern = "[①-⑳]"
        cellValue = .Replace(cellValue, "_")

        .Pattern = c_digit
        cellValue = .Replace(cellValue, "_")

        .IgnoreCase = True
        .Pattern = "^[a-z]{1,3}[0-9]{1,7}$"
        If .test(StrConv(cellValue, vbNarrow)) Then cellValue = "_" & cellValue
    End With

    If StrConv(Left(cellValue, 1), vbNarrow) Like "[0-9]*" Or StrConv(Left(cellValue, 1), vbNarrow) = ".*" Then
        cellValue = "_" & cellValue
    End If

    Rename = cellValue
End Function
