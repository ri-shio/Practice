Attribute VB_Name = "Module1"
Option Explicit

'Function�̃e�X�g�p�ɊȈՓI��Sub��ǋL
Sub test()
    Dim rng As Range
    
    Set rng = ActiveCell
    Call Q16(rng)
    
End Sub

'�𓚂�����Sub����Function�ɏ�������
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

