VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ErrorReportForm 
   Caption         =   "UserForm1"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10800
   OleObjectBlob   =   "ErrorReportForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ErrorReportForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Private Sub UserForm_Initialize()
    ' Настройка размеров формы
    Me.Width = 700
    Me.Height = 450
    
    ' Заголовок формы
    Me.Caption = "Отчет об ошибках обработки"
    
    ' Настройка метки заголовка
    With lblErrorTitle
        .Caption = "? Обнаружены ошибки при обработке данных:"
        .Top = 10
        .Left = 20
        .Width = 650
        .Height = 25
        .Font.Size = 12
        .Font.Bold = True
        .ForeColor = RGB(200, 0, 0)
    End With
    
    ' Настройка текстового поля с ошибками
    With txtErrorDetails
        .MultiLine = True
        .ScrollBars = fmScrollBarsVertical
        .Width = 650
        .Height = 280
        .Top = 45
        .Left = 20
        .Font.Name = "Consolas"
        .Font.Size = 9
        .BackColor = RGB(255, 250, 250)
        .Locked = True
    End With
    
    ' Настройка метки с инструкциями
    With lblInstructions
        .Caption = "?? Исправьте ошибки и повторите обработку данных"
        .Top = 340
        .Left = 20
        .Width = 650
        .Height = 20
        .Font.Size = 10
        .ForeColor = RGB(0, 100, 0)
    End With
    
    ' Настройка кнопки "Скопировать в буфер"
    With btnCopyToClipboard
        .Caption = "?? Скопировать отчет"
        .Width = 130
        .Height = 30
        .Top = 370
        .Left = 430
        .BackColor = RGB(0, 120, 215)
        .ForeColor = RGB(255, 255, 255)
    End With
    
    ' Настройка кнопки "Сохранить в файл"
    With btnSaveToFile
        .Caption = "?? Сохранить в файл"
        .Width = 130
        .Height = 30
        .Top = 370
        .Left = 570
        .BackColor = RGB(0, 150, 0)
        .ForeColor = RGB(255, 255, 255)
    End With
    
    ' Настройка кнопки "Закрыть"
    With btnClose
        .Caption = "? Закрыть"
        .Width = 100
        .Height = 30
        .Top = 370
        .Left = 20
        .BackColor = RGB(150, 150, 150)
        .ForeColor = RGB(255, 255, 255)
    End With
End Sub

Public Sub SetErrorText(errorText As String)
    ' Устанавливаем текст ошибок
    txtErrorDetails.text = errorText
    
    ' Подсчитываем количество ошибок
    Dim errorCount As Integer
    errorCount = UBound(Split(errorText, vbNewLine)) + 1
    
    ' Обновляем заголовок с количеством ошибок
    lblErrorTitle.Caption = "? Обнаружено ошибок: " & errorCount
End Sub

Private Sub btnCopyToClipboard_Click()
    On Error GoTo ErrorHandler
    
    ' Создаем объект для работы с буфером обмена
    Dim dataObj As Object
    Set dataObj = CreateObject("MSForms.DataObject")
    
    ' Формируем полный отчет
    Dim fullReport As String
    fullReport = "=== ОТЧЕТ ОБ ОШИБКАХ ОБРАБОТКИ ДАННЫХ ===" & vbNewLine & vbNewLine & _
                Format(Now, "dd.mm.yyyy hh:mm:ss") & vbNewLine & vbNewLine & _
                txtErrorDetails.text
    
    ' Копируем в буфер
    dataObj.SetText fullReport
    dataObj.PutInClipboard
    
    MsgBox "Отчет об ошибках скопирован в буфер обмена!", vbInformation, "Копирование завершено"
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при копировании в буфер: " & Err.description, vbExclamation, "Ошибка"
End Sub

Private Sub btnSaveToFile_Click()
    On Error GoTo ErrorHandler
    
    ' Диалог сохранения файла
    Dim fileName As String
    fileName = Application.GetSaveAsFilename( _
        InitialFileName:="Отчет_об_ошибках_" & Format(Now, "yyyy-mm-dd_hh-mm") & ".txt", _
        FileFilter:="Текстовые файлы (*.txt), *.txt", _
        Title:="Сохранить отчет об ошибках")
    
    If fileName = "False" Then Exit Sub ' Пользователь отменил
    
    ' Формируем полный отчет
    Dim fullReport As String
    fullReport = "=== ОТЧЕТ ОБ ОШИБКАХ ОБРАБОТКИ ДАННЫХ ===" & vbNewLine & vbNewLine & _
                "Дата и время: " & Format(Now, "dd.mm.yyyy hh:mm:ss") & vbNewLine & _
                "Файл: " & ActiveWorkbook.Name & vbNewLine & vbNewLine & _
                "ДЕТАЛИ ОШИБОК:" & vbNewLine & _
                String(50, "=") & vbNewLine & vbNewLine & _
                txtErrorDetails.text
    
    ' Сохраняем в файл
    Dim fileNum As Integer
    fileNum = FreeFile
    Open fileName For Output As fileNum
    Print #fileNum, fullReport
    Close fileNum
    
    MsgBox "Отчет сохранен в файл:" & vbNewLine & fileName, vbInformation, "Сохранение завершено"
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при сохранении файла: " & Err.description, vbExclamation, "Ошибка"
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

