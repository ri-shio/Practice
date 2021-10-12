Attribute VB_Name = "Module1"
Option Explicit

Sub Q37()
    Dim graphLocation As Range
    Dim graphRange As Range
    
    '#######################################
    '#
    '#  �O���t�̃f�[�^�̈�A�`��ʒu������B
    '#
    '#######################################
    Set graphLocation = Cells(1, 3)
    Set graphRange = Range(Cells(2, 1), Cells(Cells(2, 2).CurrentRegion.Rows.Count, 2))
    
    
    
    '�ΏۃV�[�g�̃O���t�����ׂč폜����B
    Dim shp As Object
    
    For Each shp In ActiveSheet.Shapes
        If shp.Type = msoChart Then shp.Delete
    Next
    
    '�O���t�̃f�[�^�̈�A�`��ʒu�ɏ]���A�_�O���t���쐬����B
    With ActiveSheet.Shapes.AddChart(xlColumnClustered, graphLocation.Left + 5, graphLocation.Top + 3)
        .Chart.SetSourceData Source:=graphRange
    End With
    
    '�O���t�̃f�[�^�̈������ő�l�E�ŏ��l���擾����B
    Dim maxValue As Long
    Dim minValue As Long
    
    With WorksheetFunction
        maxValue = .Max(graphRange)
        minValue = .Min(graphRange)
    End With
    
    '�ő�l�E�ŏ��l�������ڂ̍s������A���ڂ̍��ڂɂ����邩���擾����B
    '�ő�l�E�ŏ��l�������ڂ���������ꍇ�ɔ����A�z��Ɋi�[����B
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
    
    '�ő�l�E�ŏ��l�̍��ڏ������z�񂩂�������o���A������ύX����B
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
    
    '������폜���A�}�N���͏I���B
    ActiveSheet.ChartObjects(1).Chart.Legend.Delete
End Sub
