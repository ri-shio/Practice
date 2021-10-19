Attribute VB_Name = "Module1"
Option Explicit

Sub Q22()
    Const num_end As Integer = 30
    Dim i As Integer
    
    For i = 1 To num_end
        Select Case True
            Case i Mod 3 = 0 And i Mod 5 = 0
                Cells(i, 4) = "FizzBuzz"
            Case i Mod 5 = 0
                Cells(i, 3) = "Buzz"
            Case i Mod 3 = 0
                Cells(i, 2) = "Fizz"
            Case Else
                Cells(i, 1) = i
        End Select
    Next
End Sub

