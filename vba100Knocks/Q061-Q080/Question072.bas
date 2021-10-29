Attribute VB_Name = "Module1"
Option Explicit

Sub testSub()
    Dim testArr As Variant
    testArr = Array("ＩＴ", "とIT", "it は", "IT99", "ＧＩＴ", "site", "It's", "itはITでIt's ITでもITだ")
    
    Dim i As Integer
    
    For i = LBound(testArr) To UBound(testArr)
    
    MsgBox ("置換前：" & testArr(i) & vbCrLf _
    & "置換後：" & VARDX(testArr(i)))
    
    Next
End Sub

Function VARDX(ByVal str As String) As String
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    
    With reg
        .Pattern = "[iｉ][tｔ]"
        .IgnoreCase = True
        .Global = True
    End With
    
    '大文字小文字・半角全角を問わない"IT"を一度拾い上げた後、
    '前後にアルファベットがある場合は除外フラグを立てて置換しないようにする。
    'なお、スペースやアポストロフィーとしてのシングルクォーテーションは判定せず無視する。
    
    Dim matchReg As Variant
    
    Dim i As Long
    Dim except As Boolean
    
    For Each matchReg In reg.Execute(str)
        except = False
        
        If Not matchReg.firstindex = 0 Then
            i = matchReg.firstindex
            Do
                If i = 0 Then Exit Do
                Select Case Mid(str, i, 1)
                    Case "a" To "z", "A" To "Z", "ａ" To "ｚ", "Ａ" To "Ｚ"
                        except = True
                        Exit Do
                    Case Is = " ", "'", "　", "’"
                        i = i - 1
                    Case Else
                        Exit Do
                End Select
            Loop
        End If
        
        If Not matchReg.firstindex + 2 = Len(str) Then
            i = matchReg.firstindex + 3
            Do
                If i = Len(str) + 1 Then Exit Do
                Select Case Mid(str, i, 1)
                    Case "a" To "z", "A" To "Z", "ａ" To "ｚ", "Ａ" To "Ｚ"
                        except = True
                        Exit Do
                    Case Is = " ", "'"
                        i = i + 1
                    Case Else
                        Exit Do
                End Select
            Loop
        End If
        
        If except = False Then
            Mid(str, matchReg.firstindex + 1) = "DX"
        End If
    Next

    VARDX = str

End Function