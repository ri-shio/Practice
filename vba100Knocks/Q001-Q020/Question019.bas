Attribute VB_Name = "Module1"
Option Explicit

Sub Q19()
    Dim shp As Shape
    Dim pasted_shp As Object
    
    '解答を見て重複しない仕組みを追記
    For Each shp In ActiveSheet.Shapes
        If shp.Name Like "*_Pasted" Then
            shp.Delete
        End If
    Next
    
    For Each shp In ActiveSheet.Shapes
            If Not shp.Name Like "*Drop Down*" Then
                shp.Copy
                ActiveSheet.Paste
                Set pasted_shp = Selection
                pasted_shp.Name = shp.Name & "_Pasted"
                pasted_shp.Top = shp.Top
                pasted_shp.Left = shp.Left + shp.Width
            End If
    Next

End Sub
