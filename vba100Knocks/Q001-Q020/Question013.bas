Attribute VB_Name = "Module2"
Option Explicit

Sub Q13()
Dim rng As Range
Dim shp As Object
Dim instr_count As Integer

Dim i As Integer

If TypeName(Selection) = "Range" Then
    For Each rng In Selection
        If rng.Value Like "*注意*" Then
            instr_count = 0
            Do
                instr_count = InStr(instr_count + 1, rng.Value, "注意")
                If instr_count = 0 Then Exit Do
                With rng.Characters(instr_count, 2).Font
                  .Bold = True
                  .Color = vbRed
                End With
            Loop
        End If
    Next
End If

End Sub
