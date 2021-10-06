Attribute VB_Name = "Module1"
Option Explicit

'Functionのテスト用に簡易的なSubを追記
Sub test()
    Dim rng As Range
    
    Set rng = ActiveCell
    Call Q16(rng)
    
End Sub

'解答を見てSubからFunctionに書き換え
Function Q16(ByVal rng As Range)

    Do
        If InStr(rng.Value, vbLf) = 1 Then
            rng.Characters(1, 1).Delete
        Else
            Exit Do
        End If
    Loop
    
    Do
        If InStr(rng.Value, vbLf & vbLf) <> 0 Then
            rng.Value = Replace(rng.Value, vbLf & vbLf, vbLf)
        Else
            Exit Do
        End If
    Loop
    
    Do
        If InStr(rng.Value, vbLf) = Len(rng.Value) Then
            rng.Characters(Len(rng.Value), 1).Delete
        Else
            Exit Do
        End If
    Loop
    
End Function

