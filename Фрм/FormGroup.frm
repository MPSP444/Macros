VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormGroup 
   Caption         =   "UserForm4"
   ClientHeight    =   4680
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7680
   OleObjectBlob   =   "FormGroup.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Автоматическая настройка формы при загрузке:
Private Sub UserForm_Initialize()
    ' Настройки самой формы
    Me.Caption = "Группировка товаров по ангарам"
    Me.Width = 420
    Me.Height = 320
    
    ' Настройки TextBox
    With Me.txtInput
        .MultiLine = True
        .ScrollBars = fmScrollBarsVertical
        .Width = 380
        .Height = 220
        .Top = 10
        .Left = 10
        .Font.Size = 10
    End With
    
    ' Настройки кнопки "Обработать"
    With Me.btnProcess
        .Caption = "Обработать"
        .Width = 110
        .Height = 35
        .Top = 240
        .Left = 80
        .Font.Size = 10
        .Font.Bold = True
    End With
    
    ' Настройки кнопки "Отмена"
    With Me.btnCancel
        .Caption = "Отмена"
        .Width = 110
        .Height = 35
        .Top = 240
        .Left = 210
        .Font.Size = 10
    End With
End Sub

' Код для кнопки "Обработать":
Private Sub btnProcess_Click()
    Dim inputText As String
    inputText = Me.txtInput.value
    
    If Trim(inputText) = "" Then
        MsgBox "Пожалуйста, вставьте данные!", vbExclamation
        Exit Sub
    End If
    
    ' Закрыть форму и обработать данные
    Me.Hide
    ProcessWarehouseData inputText
    Unload Me
End Sub

' Код для кнопки "Отмена":
Private Sub btnCancel_Click()
    Unload Me
End Sub
