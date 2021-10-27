VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   2700
   ClientLeft      =   110
   ClientTop       =   460
   ClientWidth     =   5780
   OleObjectBlob   =   "Question068.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btn登録_Click()
    Dim dAddRow As Long
    Dim dAddCol As Long
    
    dAddRow = Sheets("Sheet1").Cells(1, 1).CurrentRegion.Rows.Count
    dAddCol = Sheets("Sheet1").Cells(1, 1).End(xlToRight).Column

    Dim i As Integer
    Dim tgt As Object
    Dim fixedText As String
    Dim headerRng As Range
    
    'TextBoxが1から9まであることは既知とする。
    For i = 1 To 9
        Set tgt = UserForm1.Controls("TextBox" & i)
        
        Set headerRng = Sheets("Sheet1").Rows(1).Find(What:=tgt.Name, LookAt:=xlWhole, LookIn:=xlValues)
        If headerRng Is Nothing Then
            Set headerRng = Sheets("Sheet1").Cells(1, 1).Offset(0, dAddCol)
        End If
        
        fixedText = Replace(tgt.Text, vbCrLf, "")
        headerRng.Offset(dAddRow, 0) = fixedText
    Next
    
End Sub
