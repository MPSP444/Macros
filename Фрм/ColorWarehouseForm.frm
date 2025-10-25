VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ColorWarehouseForm 
   Caption         =   "MPSP-444"
   ClientHeight    =   8130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7230
   OleObjectBlob   =   "ColorWarehouseForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ColorWarehouseForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Me.Width = 500
    Me.Height = 400
    
    Me.Caption = "Покраска текста ячеек склада"
    
    With TextBox1
        .MultiLine = True
        .ScrollBars = fmScrollBarsVertical
        .Width = 460
        .Height = 280
        .Top = 30
        .Left = 20
        .text = "Вводите координаты ячеек для покраски текста с новой строки:" & vbNewLine & _
                "Формат: номер_ангара-ряд-ячейка-уровень" & vbNewLine & _
                "Примеры:" & vbNewLine & _
                "7-1-А-1        (А/ниж.ряд)" & vbNewLine & _
                "7-1-ПРЗ-2      (ПРЗ/2ряд)" & vbNewLine & _
                "7-2-В-3        (В/3ряд)"
    End With
    
    With CommandButton1
        .Caption = "Покрасить"
        .Width = 100
        .Height = 30
        .Top = 320
        .Left = 380
    End With
    
    With CommandButton2
        .Caption = "Очистить"
        .Width = 100
        .Height = 30
        .Top = 320
        .Left = 270
    End With
End Sub

Private Sub CommandButton1_Click()
    Dim inputText As String
    inputText = TextBox1.text
    
    If Trim(inputText) = "" Then
        MsgBox "Введите координаты ячеек для покраски!", vbExclamation
        Exit Sub
    End If
    
    Dim locations() As WarehouseLocation
    locations = ModuleColor.ParseColorInput(inputText)
    
    Dim locationCount As Integer
    locationCount = UBound(locations) + 1
    
    If locationCount <= 0 Then
        MsgBox "Не найдено корректных координат для покраски!", vbExclamation
        Exit Sub
    End If
    
    Dim confirmMsg As String
    confirmMsg = "Подтвердите покраску следующих ячеек:" & vbNewLine & vbNewLine
    
    Dim i As Integer
    For i = 0 To locationCount - 1
        confirmMsg = confirmMsg & _
                    "Ангар " & locations(i).warehouse & _
                    ", Ряд " & locations(i).row & _
                    ", Ячейка " & locations(i).Letter & _
                    ", Ярус " & locations(i).level & vbNewLine
    Next i
    
    If MsgBox(confirmMsg & vbNewLine & "Выполнить покраску?", _
              vbYesNo + vbQuestion, "Подтверждение") = vbYes Then
        
        Dim result As String
        result = ModuleColor.ColorCells(locations, locationCount)
        
        If result = "" Then
            MsgBox "Все ячейки успешно покрашены!", vbInformation
            TextBox1.text = ""
        Else
            MsgBox "Произошли ошибки:" & vbNewLine & result, vbExclamation
        End If
    End If
End Sub

Private Sub CommandButton2_Click()
    TextBox1.text = ""
End Sub

