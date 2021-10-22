VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4740
   ClientLeft      =   110
   ClientTop       =   460
   ClientWidth     =   4980
   OleObjectBlob   =   "Question067.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb都道府県_Change()
    Dim Pref As String
    Dim pRange As Range
    Pref = cmb都道府県.Text
    
    'リストシートがアクティブのままだと、
    '何故かリストボックスがおかしくなる。
    'データ加工用シートをアクティブにすることで上記の状態は回避できる。
    wsDP.Activate
    wsDP.Cells(1, 1).Select
    wsDP.Cells.ClearContents
    
    wsDF.Cells(1, 1).CurrentRegion.AutoFilter _
    field:=1, Criteria1:=Pref
    
    wsDF.Cells(1, 1).CurrentRegion.Copy _
    Destination:=wsDP.Cells(1, 1)
    
    '都道府県でフィルターをかけた際に該当者が一人もいない場合に備え、
    'エラー回避、および空欄のSetを行う。
    '設問の場合は高知県が該当する。
    
    With wsDP.Cells(1, 1).CurrentRegion
        On Error Resume Next
        Set pRange = .Offset(1, 0).Resize(.Rows.Count - 1)
        On Error GoTo 0
    End With
    
    If pRange Is Nothing Then Set pRange = wsDP.Cells(1, 1).CurrentRegion.Offset(1, 0)
    
    With Me.lst個人
        .ColumnCount = pRange.Columns.Count
        .ColumnHeads = True
        .RowSource = pRange.Address
    End With
    
    Sheets("リスト").Activate
    Sheets("リスト").Cells(1, 1).Select
End Sub
