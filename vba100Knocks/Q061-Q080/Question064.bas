Attribute VB_Name = "Module1"
Option Explicit

Sub Q64()
    Dim pasteRng1 As Range, pasteRng2 As Range
    Set pasteRng1 = Sheets("まとめ").Range("A1:J20")
    Set pasteRng2 = Sheets("まとめ").Range("A21:J40")
    
    Dim pct1 As Object, pct2 As Object
    
    Sheets("まとめ").Pictures.Delete
    
    Sheets("元表1").Cells(1, 1).CurrentRegion.Copy
    Sheets("まとめ").Activate
    pasteRng1.Select
    Set pct1 = ActiveSheet.Pictures.Paste(Link:=True)
    
    With pct1
        If .Width / pasteRng1.Width >= .Height / pasteRng1.Height Then
            .Width = pasteRng1.Width
            .Top = pasteRng1.Top
            .Left = pasteRng1.Left
        Else
            .Height = pasteRng1.Height
            .Top = pasteRng1.Top
            .Left = pasteRng1.Left + (pasteRng1.Width - .Width) / 2
        End If
    End With

    Sheets("元表2").Cells(1, 1).CurrentRegion.Copy
    Sheets("まとめ").Activate
    pasteRng2.Select
    Set pct2 = ActiveSheet.Pictures.Paste(Link:=True)
    
    With pct2
        If .Width / pasteRng2.Width >= .Height / pasteRng2.Height Then
            .Width = pasteRng2.Width
            .Top = pasteRng2.Top + (pasteRng2.Height - .Height) / 2
            .Left = pasteRng2.Left
        Else
            .Height = pasteRng2.Height
            .Top = pasteRng2.Top
            .Left = pasteRng2.Left + (pasteRng2.Width - .Width) / 2
        End If
    End With
    
End Sub
