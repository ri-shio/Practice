Attribute VB_Name = "Module1"
Option Explicit

Sub Q49()
    Dim dispForm As Object
    Dim isFcApplied As Boolean
    Dim intCI As Integer, fntCI As Integer
    Dim copyCnt As Long: copyCnt = 0
    Dim i As Long

    Sheets("49Out").Cells.Clear
    Sheets("49Out").Cells(1, 1) = "科目": Sheets("49Out").Cells(1, 1) = "項目1"
    Sheets("49Out").Cells(1, 1) = "項目2": Sheets("49Out").Cells(1, 1) = "項目3"

    For i = 1 To Sheets("49In").Cells(1, 4).CurrentRegion.Rows.Count
        isFcApplied = False
        Set dispForm = Cells(i, 4).DisplayFormat

        With dispForm
            If .Interior.ColorIndex = 3 Or .Interior.ColorIndex = 6 Then isFcApplied = True
            If .Font.ColorIndex = 3 Then isFcApplied = True

            If isFcApplied = True Then
                intCI = .Interior.ColorIndex: fntCI = .Font.ColorIndex

                Sheets("49In").Range(Cells(i, 1), Cells(i, 4)).Copy _
                Destination:=Sheets("49Out").Cells(copyCnt + 2, 1)

                copyCnt = copyCnt + 1
            End If
        End With

    Next
End Sub
