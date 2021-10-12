Attribute VB_Name = "Module1"
Option Explicit

Sub Q37()
    Dim graphLocation As Range
    Dim graphRange As Range
    
    '#######################################
    '#
    '#  グラフのデータ領域、描画位置を決定。
    '#
    '#######################################
    Set graphLocation = Cells(1, 3)
    Set graphRange = Range(Cells(2, 1), Cells(Cells(2, 2).CurrentRegion.Rows.Count, 2))
    
    
    
    '対象シートのグラフをすべて削除する。
    Dim shp As Object
    
    For Each shp In ActiveSheet.Shapes
        If shp.Type = msoChart Then shp.Delete
    Next
    
    'グラフのデータ領域、描画位置に従い、棒グラフを作成する。
    With ActiveSheet.Shapes.AddChart(xlColumnClustered, graphLocation.Left + 5, graphLocation.Top + 3)
        .Chart.SetSourceData Source:=graphRange
    End With
    
    'グラフのデータ領域内から最大値・最小値を取得する。
    Dim maxValue As Long
    Dim minValue As Long
    
    With WorksheetFunction
        maxValue = .Max(graphRange)
        minValue = .Min(graphRange)
    End With
    
    '最大値・最小値を持つ項目の行数から、何個目の項目にあたるかを取得する。
    '最大値・最小値を持つ項目が複数ある場合に備え、配列に格納する。
    Dim rng As Range
    Dim maxRange() As Long, minRange() As Long
    Dim ma_cnt As Long, mi_cnt As Long
    
    ma_cnt = 0: mi_cnt = 0
    For Each rng In graphRange
        If rng.Value = maxValue Then
            ReDim Preserve maxRange(ma_cnt)
            maxRange(ma_cnt) = rng.Row - 1
            ma_cnt = ma_cnt + 1
        ElseIf rng.Value = minValue Then
            ReDim Preserve minRange(mi_cnt)
            minRange(mi_cnt) = rng.Row - 1
            mi_cnt = mi_cnt + 1
        End If
    Next
    
    '最大値・最小値の項目情報を持つ配列から情報を取り出し、書式を変更する。
    Dim i As Long
    
    With ActiveSheet.ChartObjects(1).Chart.SeriesCollection(1)
        For i = 0 To ma_cnt - 1
            .Points(maxRange(i)).Format.Fill.ForeColor.RGB = RGB(124, 252, 0)
            .Points(maxRange(i)).HasDataLabel = True
        Next
        For i = 0 To mi_cnt - 1
            .Points(minRange(i)).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
            .Points(minRange(i)).HasDataLabel = True
        Next
        
    End With
    
    '判例を削除し、マクロは終了。
    ActiveSheet.ChartObjects(1).Chart.Legend.Delete
End Sub
