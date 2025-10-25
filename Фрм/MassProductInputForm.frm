VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MassProductInputForm 
   Caption         =   "Массовый ввод товаров на склад"
   ClientHeight    =   7230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8430.001
   OleObjectBlob   =   "MassProductInputForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MassProductInputForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub UserForm_Initialize()
    ' Настройка размеров формы
    Me.Width = 800
    Me.Height = 600
    
    ' Заголовок формы
    Me.Caption = "Массовый ввод товаров на склад"
    
    ' Настройка метки с инструкциями
    With lblInstructions
        .Caption = "Введите данные в формате:" & vbNewLine & _
                  "6-5-Е-2 - Лерашанс - пар13 - 720" & vbNewLine & _
                  "Ангар-Ряд-Стеллаж-Уровень - Препарат - Партия - Количество"
        .Top = 10
        .Left = 20
        .Width = 750
        .Height = 60
        .Font.Size = 10
        .ForeColor = RGB(0, 100, 0)
    End With
    
    ' Настройка большого текстового поля
    With txtMassInput
        .MultiLine = True
        .ScrollBars = fmScrollBarsVertical
        .Width = 750
        .Height = 350
        .Top = 80
        .Left = 20
        .Font.Name = "Consolas"
        .Font.Size = 10
        
        ' Разбиваем на части, чтобы не превышать лимит в 24 продолжения
        Dim text1 As String, text2 As String, text3 As String
        
        ' Первая часть - Хранение
        text1 = "12-2-а-1 - Хранение - пар01 - 1" & vbNewLine & _
                "12-3-а-1 - Хранение - пар01 - 1" & vbNewLine & _
                "12-4-а-1 - Хранение - пар01 - 1"
        
        ' Вторая часть - Семена
        text2 = "9-1-ф-1 - Семена - пар01 - 1" & vbNewLine & _
                "9-2-ф-1 - Семена - пар01 - 1" & vbNewLine & _
                "9-3-ф-1 - Семена - пар01 - 1" & vbNewLine & _
                "9-4-ф-1 - Семена - пар01 - 1" & vbNewLine & _
                "9-5-ф-1 - Семена - пар01 - 1" & vbNewLine & _
                "9-6-ф-1 - Семена - пар01 - 1" & vbNewLine & _
                "9-7-ф-1 - Семена - пар01 - 1" & vbNewLine & _
                "9-8-ф-1 - Семена - пар01 - 1" & vbNewLine & _
                "9-9-ф-1 - Семена - пар01 - 1" & vbNewLine & _
                "9-10-ф-1 - Семена - пар01 - 1" & vbNewLine & _
                "9-11-ф-1 - Семена - пар01 - 1" & vbNewLine & _
                "9-12-ф-1 - Семена - пар01 - 1" & vbNewLine & _
                "9-13-ф-1 - Семена - пар01 - 1" & vbNewLine & _
                "9-14-ф-1 - Семена - пар01 - 1" & vbNewLine & _
                "9-15-ф-1 - Семена - пар01 - 1" & vbNewLine & _
                "9-16-ф-1 - Семена - пар01 - 1"
        
        ' Третья часть - Просрочка
        text3 = "11-1-а-1 - Просрочка - пар01 - 1" & vbNewLine & _
                "11-2-а-1 - Просрочка - пар01 - 1" & vbNewLine & _
                "11-5-а-1 - Просрочка - пар01 - 1" & vbNewLine & _
                "11-6-а-1 - Просрочка - пар01 - 1" & vbNewLine & _
                "11-7-а-1 - Просрочка - пар01 - 1" & vbNewLine & _
                "11-8-а-1 - Просрочка - пар01 - 1" & vbNewLine & _
                "11-39-а-1 - Луч - пар01 - 1" & vbNewLine & _
                "11-40-а-1 - Луч - пар01 - 1" & vbNewLine & _
                "11-41-а-1 - Луч - пар01 - 1" & vbNewLine & _
                "12-1-о-1 - Хранение - пар01 - 1" & vbNewLine & _
                "12-2-о-1 - Хранение - пар01 - 1" & vbNewLine & _
                "12-3-о-1 - Хранение - пар01 - 1" & vbNewLine & _
                "12-4-о-1 - Хранение - пар01 - 1" & vbNewLine & _
                "12-5-о-1 - Хранение - пар01 - 1" & vbNewLine & _
                "12-6-о-1 - Хранение - пар01 - 1" & vbNewLine & _
                ""
        
        ' Объединяем все части
        .text = text1 & vbNewLine & text2 & vbNewLine & text3
    End With
    
    ' Настройка кнопки "Обработать данные"
    With btnProcess
        .Caption = "? Обработать данные"
        .Width = 150
        .Height = 35
        .Top = 450
        .Left = 600
        .BackColor = RGB(0, 150, 0)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    ' Настройка кнопки "Очистить все ангары"
    With btnClearAll
        .Caption = "? Очистить все ангары"
        .Width = 150
        .Height = 35
        .Top = 450
        .Left = 430
        .BackColor = RGB(220, 100, 0)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    ' Настройка кнопки "Отмена"
    With btnCancel
        .Caption = "< Отмена"
        .Width = 100
        .Height = 35
        .Top = 450
        .Left = 320
        .BackColor = RGB(150, 150, 150)
        .ForeColor = RGB(255, 255, 255)
    End With
    
    ' Настройка метки статуса
    With lblStatus
        .Caption = "Готов к работе..."
        .Top = 500
        .Left = 20
        .Width = 750
        .Height = 20
        .Font.Size = 9
        .ForeColor = RGB(0, 0, 150)
    End With
End Sub

Private Sub btnProcess_Click()
    ' Проверяем наличие данных
    If Trim(txtMassInput.text) = "" Or InStr(txtMassInput.text, "Пример:") > 0 Then
        MsgBox "Пожалуйста, введите данные для обработки!", vbExclamation, "Нет данных"
        txtMassInput.SetFocus
        Exit Sub
    End If
    
    ' Проверяем наличие активной книги
    If ActiveWorkbook Is Nothing Then
        MsgBox "Откройте книгу Excel с ангарами перед обработкой!", vbExclamation, "Нет активной книги"
        Exit Sub
    End If
    
    ' Подтверждение обработки
    Dim confirmMsg As String
    confirmMsg = "Обработать введенные данные?" & vbNewLine & vbNewLine & _
                "? Внимание: Данные будут записаны в таблицы ангаров!" & vbNewLine & _
                "Заголовки товаров будут обновлены автоматически."
    
    If MsgBox(confirmMsg, vbYesNo + vbQuestion, "Подтверждение обработки") = vbNo Then
        Exit Sub
    End If
    
    ' Обновляем статус
    lblStatus.Caption = "Обработка данных..."
    lblStatus.ForeColor = RGB(200, 100, 0)
    DoEvents
    
    ' Отключаем кнопки во время обработки
    btnProcess.Enabled = False
    btnClearAll.Enabled = False
    btnCancel.Enabled = False
    
    ' Вызываем модуль обработки данных
    Dim result As String
    result = ModuleMassInput.ProcessMassInput(txtMassInput.text)
    
    ' Включаем кнопки обратно
    btnProcess.Enabled = True
    btnClearAll.Enabled = True
    btnCancel.Enabled = True
    
    ' Показываем результат
    If result = "" Then
        lblStatus.Caption = "? Обработка завершена успешно!"
        lblStatus.ForeColor = RGB(0, 150, 0)
        MsgBox "Все данные успешно обработаны и записаны в ангары!", vbInformation, "Успех"
        txtMassInput.text = ""
    Else
        lblStatus.Caption = "? Обработка завершена с ошибками"
        lblStatus.ForeColor = RGB(200, 0, 0)
        
        ' Показываем окно с ошибками
        Call ShowErrorReport(result)
    End If
End Sub

Private Sub btnClearAll_Click()
    ' Подтверждение очистки
    Dim confirmMsg As String
    confirmMsg = "??? ВНИМАНИЕ! Вы хотите очистить ВСЕ АНГАРЫ?" & vbNewLine & vbNewLine & _
                "Будут удалены:" & vbNewLine & _
                "• Все заголовки товаров" & vbNewLine & _
                "• Все количества" & vbNewLine & _
                "• Все партии" & vbNewLine & vbNewLine & _
                "Сохранятся только структура таблиц и названия стеллажей." & vbNewLine & vbNewLine & _
                "? Продолжить очистку?"
    
    If MsgBox(confirmMsg, vbYesNo + vbCritical, "Подтверждение очистки") = vbNo Then
        Exit Sub
    End If
    
    ' Дополнительное подтверждение
    If MsgBox("Вы ДЕЙСТВИТЕЛЬНО уверены?", vbYesNo + vbQuestion, "Последнее предупреждение") = vbNo Then
        Exit Sub
    End If
    
    ' Обновляем статус
    lblStatus.Caption = "Очистка всех ангаров..."
    lblStatus.ForeColor = RGB(200, 0, 0)
    DoEvents
    
    ' Отключаем кнопки
    btnProcess.Enabled = False
    btnClearAll.Enabled = False
    btnCancel.Enabled = False
    
    ' Вызываем модуль очистки
    Dim result As String
    result = ModuleDataCleaner.ClearAllWarehouses()
    
    ' Включаем кнопки обратно
    btnProcess.Enabled = True
    btnClearAll.Enabled = True
    btnCancel.Enabled = True
    
    ' Показываем результат
    If result = "" Then
        lblStatus.Caption = "? Все ангары успешно очищены!"
        lblStatus.ForeColor = RGB(0, 150, 0)
        MsgBox "Все ангары успешно очищены!" & vbNewLine & "Таблицы готовы для новых данных.", vbInformation, "Очистка завершена"
    Else
        lblStatus.Caption = "? Ошибка при очистке"
        lblStatus.ForeColor = RGB(200, 0, 0)
        MsgBox "Произошли ошибки при очистке:" & vbNewLine & result, vbExclamation, "Ошибка"
    End If
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub ShowErrorReport(errorText As String)
    ' Создаем и показываем форму с отчетом об ошибках
    Dim errorForm As ErrorReportForm
    Set errorForm = New ErrorReportForm
    errorForm.SetErrorText errorText
    errorForm.Show
End Sub

