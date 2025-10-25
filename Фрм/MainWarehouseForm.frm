VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainWarehouseForm 
   ClientHeight    =   11535
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15090
   OleObjectBlob   =   "MainWarehouseForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainWarehouseForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ===== MainWarehouseForm (UserForm) =====
Option Explicit

Private Sub UserForm_Initialize()
    ' Настраиваем размеры формы
    Me.Width = 500
    Me.Height = 400
    
    ' Настраиваем форму
    Me.Caption = "Ввод данных для складского учета"
    
    ' Настраиваем текстовое поле
    With TextBox1
        .MultiLine = True
        .ScrollBars = fmScrollBarsVertical
        .Width = 460
        .Height = 280
        .Top = 30
        .Left = 20
        .value = "Вводите каждую операцию с новой строки:" & vbNewLine & _
                "Формат: номер_ангара-ряд-ячейка-уровень значение [партия]" & vbNewLine & _
                "Примеры:" & vbNewLine & _
                "7-1-А-1 100        (А/ниж.ряд)" & vbNewLine & _
                "7-1-ПРЗ-2 200      (ПРЗ/2ряд)" & vbNewLine & _
                "7-2-В-3 150        (В/3ряд)" & vbNewLine & _
                "7-1-А-1 100 пар01  (добавление с партией)"
    End With
    
    ' Настраиваем кнопки
    With CommandButton1
        .Caption = "Выполнить"
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
    inputText = TextBox1.value
    
    If Trim(inputText) = "" Then
        MsgBox "Введите данные для обработки!", vbExclamation
        Exit Sub
    End If
    
     Dim operations() As WarehouseOperation
     operations = ModuleOperations.ParseInput(inputText)

    
    Dim operationCount As Integer
    operationCount = UBound(operations) + 1
    
    If operationCount <= 0 Then
        MsgBox "Неверный формат ввода!", vbExclamation
        Exit Sub
    End If
    
    Dim confirmMsg As String
    confirmMsg = "Подтвердите операции:" & vbNewLine & vbNewLine
    
    Dim i As Integer
    For i = 0 To operationCount - 1
        confirmMsg = confirmMsg & _
                    "Ангар " & operations(i).warehouse & _
                    ", Ряд " & operations(i).row & _
                    ", Ячейка " & operations(i).Letter & _
                    ", Ярус " & operations(i).level & _
                    ", Значение " & operations(i).value
        
        If operations(i).Batch <> "" Then
            confirmMsg = confirmMsg & ", Партия " & operations(i).Batch & _
                        " (Добавление)" & vbNewLine
        Else
            confirmMsg = confirmMsg & " (Вычитание)" & vbNewLine
        End If
    Next i
    
    If MsgBox(confirmMsg & vbNewLine & "Выполнить операции?", _
              vbYesNo + vbQuestion, "Подтверждение") = vbYes Then
        
        Dim result As String
        result = ModuleOperations.ProcessOperations(operations, operationCount)
        
        If result = "" Then
            MsgBox "Все операции выполнены успешно!", vbInformation
            TextBox1.value = ""
        Else
            MsgBox "Произошли ошибки:" & vbNewLine & result, vbExclamation
        End If
    End If
End Sub

Private Sub CommandButton2_Click()
    TextBox1.value = ""
End Sub

