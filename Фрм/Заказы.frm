VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Заказы 
   Caption         =   "Заказы"
   ClientHeight    =   10575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15660
   OleObjectBlob   =   "Заказы.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Заказы"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Код для формы Заказы
Private Sub UserForm_Initialize()
    ' Заполняем комбобокс сроками
    With Me.СРОК
        .AddItem "1 день"
        .AddItem "2 день"
        .AddItem "3 дня"
        .AddItem "5 день"
        .AddItem "1 неделя"
        .AddItem "2 недели"
        .AddItem "1 месяц"
    End With
    
    ' Заполняем комбобокс услугами
    With Me.УСЛУГИ
        .AddItem "ВУЗ"
        .AddItem "Диплом"
        .AddItem "Документ"
        .AddItem "Задачи и ответ на вопросы"
        .AddItem "Контрольная"
        .AddItem "Курсовой"
        .AddItem "Лабораторная"
        .AddItem "Отчет/План лагеря"
        .AddItem "Отчет/Практика"
        .AddItem "Переделать-Курсовой"
        .AddItem "Практика"
        .AddItem "Презентация"
        .AddItem "Программа"
        .AddItem "Редактировать"
        .AddItem "реферат"
        .AddItem "Самостоятельная работа"
        .AddItem "Семестровая работа"
        .AddItem "С-Курс"
        .AddItem "Тест"
        .AddItem "Экзамен"
    End With
End Sub

' Очистка текстового поля при входе в него и восстановление при выходе
Private Sub ФИО_Enter()
    If ФИО.text = "ФИО" Then
        ФИО.text = ""
    End If
End Sub

Private Sub ФИО_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If ФИО.text = "" Then
        ФИО.text = "ФИО"
    End If
End Sub

Private Sub ГРУППА_Enter()
    If ГРУППА.text = "ГРУППА" Then
        ГРУППА.text = ""
    End If
End Sub

Private Sub ГРУППА_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If ГРУППА.text = "" Then
        ГРУППА.text = "ГРУППА"
    End If
End Sub

Private Sub НОМЕР_Enter()
    If номер.text = "НОМЕР" Then
        номер.text = ""
    End If
End Sub

Private Sub НОМЕР_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If номер.text = "" Then
        номер.text = "НОМЕР"
    End If
End Sub

Private Sub УСЛУГИ_Enter()
    If УСЛУГИ.text = "УСЛУГИ" Then
        УСЛУГИ.text = ""
    End If
End Sub

Private Sub УСЛУГИ_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If УСЛУГИ.text = "" Then
        УСЛУГИ.text = "УСЛУГИ"
    End If
End Sub

Private Sub НАПРАВЛЕНИЯ_Enter()
    If НАПРАВЛЕНИЯ.text = "НАПРАВЛЕНИЯ" Then
        НАПРАВЛЕНИЯ.text = ""
    End If
End Sub

Private Sub НАПРАВЛЕНИЯ_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If НАПРАВЛЕНИЯ.text = "" Then
        НАПРАВЛЕНИЯ.text = "НАПРАВЛЕНИЯ"
    End If
End Sub

Private Sub КОЛИЧЕСТВО_Enter()
    If КОЛИЧЕСТВО.text = "КОЛИЧЕСТВО" Then
        КОЛИЧЕСТВО.text = ""
    End If
End Sub

Private Sub КОЛИЧЕСТВО_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If КОЛИЧЕСТВО.text = "" Then
        КОЛИЧЕСТВО.text = "КОЛИЧЕСТВО"
    End If
End Sub

Private Sub СТОИМОСТЬ_Enter()
    If СТОИМОСТЬ.text = "СТОИМОСТЬ" Then
        СТОИМОСТЬ.text = ""
    End If
End Sub

Private Sub СТОИМОСТЬ_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If СТОИМОСТЬ.text = "" Then
        СТОИМОСТЬ.text = "СТОИМОСТЬ"
    End If
End Sub

Private Sub ЛОГИН_Enter()
    If ЛОГИН.text = "ЛОГИН" Then
        ЛОГИН.text = ""
    End If
End Sub

Private Sub ЛОГИН_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If ЛОГИН.text = "" Then
        ЛОГИН.text = "ЛОГИН"
    End If
End Sub

Private Sub ПАРОЛЬ_Enter()
    If ПАРОЛЬ.text = "ПАРОЛЬ" Then
        ПАРОЛЬ.text = ""
    End If
End Sub

Private Sub ПАРОЛЬ_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If ПАРОЛЬ.text = "" Then
        ПАРОЛЬ.text = "ПАРОЛЬ"
    End If
End Sub

Private Sub САЙТ_Enter()
    If САЙТ.text = "САЙТ" Then
        САЙТ.text = ""
    End If
End Sub

Private Sub САЙТ_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If САЙТ.text = "" Then
        САЙТ.text = "САЙТ"
    End If
End Sub

Private Sub ФИОПРЕПОДАВАТЕЛЯ_Enter()
    If ФИОПРЕПОДАВАТЕЛЯ.text = "ФИО ПРЕПОДАВАТЕЛЯ" Then
        ФИОПРЕПОДАВАТЕЛЯ.text = ""
    End If
End Sub

Private Sub ФИОПРЕПОДАВАТЕЛЯ_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If ФИОПРЕПОДАВАТЕЛЯ.text = "" Then
        ФИОПРЕПОДАВАТЕЛЯ.text = "ФИО ПРЕПОДАВАТЕЛЯ"
    End If
End Sub

Private Sub ПОЧТАПРЕПОДАВАТЕЛЯ_Enter()
    If ПОЧТАПРЕПОДАВАТЕЛЯ.text = "ПОЧТА ПРЕПОДАВАТЕЛЯ" Then
        ПОЧТАПРЕПОДАВАТЕЛЯ.text = ""
    End If
End Sub

Private Sub ПОЧТАПРЕПОДАВАТЕЛЯ_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If ПОЧТАПРЕПОДАВАТЕЛЯ.text = "" Then
        ПОЧТАПРЕПОДАВАТЕЛЯ.text = "ПОЧТА ПРЕПОДАВАТЕЛЯ"
    End If
End Sub

Private Sub ТЕМА_Enter()
    If ТЕМА.text = "ТЕМА" Then
        ТЕМА.text = ""
    End If
End Sub

Private Sub ТЕМА_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If ТЕМА.text = "" Then
        ТЕМА.text = "ТЕМА"
    End If
End Sub
Private Sub ДОБАВИТЬ_Click()
    On Error GoTo ErrorHandler
    
    ' Проверяем заполнение основных полей
    If ФИО.text = "ФИО" Or ГРУППА.text = "ГРУППА" Or УСЛУГИ.value = "" Then
        MsgBox "Пожалуйста, заполните основные поля (ФИО, Группа, Услуги)", vbExclamation
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim emptyRow As Long
    Const MAX_ROWS As Long = 800  ' Максимальное количество строк в таблице
    
    ' Получаем лист и таблицу
    Set ws = activeSheet
    Set tbl = ws.ListObjects("Таблица2")
    
    ' Ищем первую пустую строку в пределах MAX_ROWS
    For emptyRow = 1 To MAX_ROWS
        If isEmpty(tbl.ListColumns(1).Range(emptyRow + 1)) Then  ' +1 для пропуска заголовка
            Exit For
        End If
    Next emptyRow
    
    ' Проверяем, не вышли ли за пределы допустимого диапазона
    If emptyRow > MAX_ROWS Then
        MsgBox "Таблица заполнена полностью (максимум " & MAX_ROWS & " строк).", vbExclamation
        Exit Sub
    End If
    
    ' Добавляем данные в найденную пустую строку
    With tbl.Range.rows(emptyRow + 1)  ' +1 для учета заголовка
        .Cells(1).value = IIf(ФИО.text = "ФИО", "", ФИО.text)
        .Cells(2).value = IIf(ГРУППА.text = "ГРУППА", "", ГРУППА.text)
        .Cells(3).value = УСЛУГИ.value
        .Cells(4).value = IIf(НАПРАВЛЕНИЯ.text = "НАПРАВЛЕНИЯ", "", НАПРАВЛЕНИЯ.text)
        .Cells(5).value = IIf(ТЕМА.text = "ТЕМА", "", ТЕМА.text)
        .Cells(6).value = IIf(КОЛИЧЕСТВО.text = "КОЛИЧЕСТВО", "", КОЛИЧЕСТВО.text)
        .Cells(7).value = IIf(ЛОГИН.text = "ЛОГИН", "", ЛОГИН.text)
        .Cells(8).value = IIf(ПАРОЛЬ.text = "ПАРОЛЬ", "", ПАРОЛЬ.text)
        .Cells(9).value = IIf(САЙТ.text = "САЙТ", "", САЙТ.text)
        .Cells(10).value = IIf(СТОИМОСТЬ.text = "СТОИМОСТЬ", "", СТОИМОСТЬ.text)
        .Cells(11).value = ""  ' ОПЛАТИЛ
        .Cells(12).value = ""  ' ОСТАЛОСЬ
        .Cells(13).value = ""  ' ОБЩАЯ СУММА
        .Cells(14).value = IIf(ФИОПРЕПОДАВАТЕЛЯ.text = "ФИО ПРЕПОДАВАТЕЛЯ", "", ФИОПРЕПОДАВАТЕЛЯ.text)
        .Cells(15).value = IIf(ПОЧТАПРЕПОДАВАТЕЛЯ.text = "ПОЧТА ПРЕПОДАВАТЕЛЯ", "", ПОЧТАПРЕПОДАВАТЕЛЯ.text)
        .Cells(16).value = IIf(номер.text = "НОМЕР", "", номер.text)
        .Cells(17).value = Format(Date, "dd.mm.yyyy")
        .Cells(18).value = СРОК.value
    End With
    
    ' Показываем добавленную строку
    tbl.Range.rows(emptyRow + 1).Select
    Application.GoTo tbl.Range.rows(emptyRow + 1), True
    
    ' Сообщение об успешном добавлении
    MsgBox "Данные добавлены в строку " & emptyRow & " таблицы.", vbInformation
    
    ' Очищаем форму
    ClearForm
    
ExitSub:
    Exit Sub

ErrorHandler:
    MsgBox "Произошла ошибка:" & vbNewLine & _
           "Номер: " & Err.Number & vbNewLine & _
           "Описание: " & Err.description, vbCritical
    Resume ExitSub
End Sub
Private Sub ОЧИСТИТЬ_Click()
    ClearForm
End Sub

Private Sub ClearForm()
    ' Очистка всех полей формы
    ФИО.text = "ФИО"
    ГРУППА.text = "ГРУППА"
    номер.text = "НОМЕР"
    УСЛУГИ.text = "УСЛУГИ"
    НАПРАВЛЕНИЯ.text = "НАПРАВЛЕНИЯ"
    КОЛИЧЕСТВО.text = "КОЛИЧЕСТВО"
    СРОК.value = ""
    СТОИМОСТЬ.text = "СТОИМОСТЬ"
    ЛОГИН.text = "ЛОГИН"
    ПАРОЛЬ.text = "ПАРОЛЬ"
    САЙТ.text = "САЙТ"
    ФИОПРЕПОДАВАТЕЛЯ.text = "ФИО ПРЕПОДАВАТЕЛЯ"
    ПОЧТАПРЕПОДАВАТЕЛЯ.text = "ПОЧТА ПРЕПОДАВАТЕЛЯ"
End Sub

