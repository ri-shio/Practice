Attribute VB_Name = "Module1"
Option Explicit

'ŒŸØ—p‚ÉŠÈˆÕ“I‚ÈSub‚ğ—pˆÓB
Sub test()
    Dim str_fnc As String
    
    str_fnc = "‚Ä‚·‚ÆƒeƒXƒg‚”‚…‚“‚”‚s‚d‚r‚stestTEST123‚P‚Q‚R"
    Call Q24(str_fnc)
    Debug.Print str_fnc

End Sub

Function Q24(ByRef str_fnc As String) As String
    Dim letters() As String
    Dim i As Long
    Dim length_str_fnc
    ReDim letters(1 To Len(str_fnc))
    
    length_str_fnc = Len(str_fnc)
    For i = 1 To length_str_fnc
        letters(i) = Mid(str_fnc, i, 1)
        Select Case True
            Case letters(i) Like "[a-z]"
                letters(i) = UCase(letters(i))
            Case letters(i) Like "[‚-‚š]"
                letters(i) = StrConv(UCase(letters(i)), vbNarrow)
            Case letters(i) Like "[‚`-‚y]" Or letters(i) Like "[‚O-‚X]"
                letters(i) = StrConv(letters(i), vbNarrow)
        End Select
    Next
    
    str_fnc = ""
    For i = 1 To length_str_fnc
        str_fnc = str_fnc & letters(i)
    Next

End Function

