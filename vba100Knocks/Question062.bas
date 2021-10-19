Attribute VB_Name = "Module1"
Option Explicit


Function ZLOOKUP(luValue As String, tblArr As Range, colIndNum As Integer, order As Integer) As Variant
    Dim hitRange As Range
    Dim startRange As Range
    Dim nextRange As Range
    
    Set hitRange = tblArr.Columns(1).Find(what:=luValue, lookat:=xlWhole, searchdirection:=xlNext)
    Set startRange = hitRange
    Set nextRange = hitRange
    
    If hitRange Is Nothing Then
        ZLOOKUP = CVErr(xlErrNA)
        Exit Function
    Else
        Do
        
            If tblArr.Columns(1).FindNext(nextRange).Address = startRange.Address Then
                Exit Do
            Else
                Set hitRange = Union(hitRange, tblArr.Columns(1).FindNext(nextRange))
                Set nextRange = tblArr.Columns(1).FindNext(nextRange)
            End If
        Loop
        
        Select Case order
            Case Is = -1
                ZLOOKUP = hitRange(hitRange.Count).Offset(0, colIndNum).Value
                
            Case Is = 0
                ZLOOKUP = hitRange(1).Offset(0, colIndNum).Value
                
            Case Is >= 1
                If order <= hitRange.Count Then
                    ZLOOKUP = hitRange(order).Offset(0, colIndNum).Value
                Else
                    ZLOOKUP = ""
                End If
                
            Case Else
                    ZLOOKUP = CVErr(xlErrValue)
                    
        End Select
        
    End If
    
End Function
