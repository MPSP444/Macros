VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ФормаОплаты 
   Caption         =   "UserForm1"
   ClientHeight    =   9615.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15120
   OleObjectBlob   =   "ФормаОплаты.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ФормаОплаты"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Глобальные переменные для хранения найденного ФИО
Private foundFIO As String
Private foundRows As String

' Инициализация формы
Private Sub UserForm_Initialize()
    ' Очищаем все поля
    txtФИО.text = ""
    txtИсторияЗаказа.text = ""
    txtОплата.text = ""
    cmbУслуги.Clear
    cmbГРУППЫ.Clear
End Sub
Private Function searchFIO(dbFIO As String, searchFIO As String) As Boolean
    ' Проверка на пустые значения
    If dbFIO = "" Or searchFIO = "" Then
        searchFIO = False
        Exit Function
    End If
    
    ' Объявляем все переменные в начале
    Dim dbParts() As String, searchParts() As String
    Dim i As Long, j As Long, matchCount As Long
    
    ' Разбиваем строки на части и удаляем лишние пробелы
    dbParts = Split(LCase(Trim(dbFIO)))
    searchParts = Split(LCase(Trim(searchFIO)))
    
    ' Инициализируем счетчик совпадений
    matchCount = 0
    
    ' Проверяем совпадение в любом порядке
    For i = 0 To UBound(searchParts)
        If Len(Trim(searchParts(i))) > 0 Then
            For j = 0 To UBound(dbParts)
                If Len(Trim(dbParts(j))) > 0 Then
                    If InStr(1, dbParts(j), searchParts(i)) > 0 Then
                        matchCount = matchCount + 1
                        Exit For
                    End If
                End If
            Next j
        End If
    Next i
    
    ' Возвращаем результат
    searchFIO = (matchCount >= UBound(searchParts) + 1)
End Function
Private Sub btnПроверить_Click()
    ' Проверка ввода
    If Trim(txtФИО.text) = "" Then
        MsgBox "Введите ФИО!", vbExclamation
        txtФИО.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(txtФИО.text)) < 2 Then
        MsgBox "Введите минимум 2 символа ФИО", vbExclamation
        txtФИО.SetFocus
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Заказы")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).row
    
    ' Очищаем все поля
    txtИсторияЗаказа.text = ""
    cmbУслуги.Clear
    cmbГРУППЫ.Clear
    foundRows = ""
    foundFIO = ""
    
    ' История всех заказов
    Dim history As String
    history = "ИСТОРИЯ ЗАКАЗОВ:" & vbNewLine & vbNewLine
    
    ' Поиск заказов
    Dim i As Long
    Dim found As Boolean
    found = False
    
    ' Очищаем предыдущее форматирование
    ws.Range("A2:A" & lastRow).Interior.ColorIndex = xlNone
    
    For i = 2 To lastRow
        Dim currentFIO As String
        currentFIO = ws.Cells(i, 1).value
        
        ' Проверяем совпадение в обоих направлениях
        If searchFIO(currentFIO, txtФИО.text) Then
            found = True
            foundFIO = currentFIO ' Сохраняем найденное ФИО
            
            ' Форматируем ячейку с ФИО
            Dim paid As Double
            Dim total As Double
            total = CDbl(IIf(ws.Cells(i, 10).value = "", 0, ws.Cells(i, 10).value))
            paid = CDbl(IIf(ws.Cells(i, 11).value = "", 0, ws.Cells(i, 11).value))
            
            ' Если заказ полностью оплачен - зеленый цвет
            If paid >= total Then
                ws.Cells(i, 1).Interior.Color = RGB(198, 239, 206) ' Светло-зеленый
            End If
            
            ' Добавляем в историю
            history = history & "ФИО: " & currentFIO & vbNewLine & _
                     "Группа: " & ws.Cells(i, 2).value & vbNewLine & _
                     "Услуга: " & ws.Cells(i, 3).value & vbNewLine & _
                     "Стоимость: " & FormatNumber(total, 2) & " руб." & vbNewLine & _
                     "Оплачено: " & FormatNumber(paid, 2) & " руб." & vbNewLine & _
                     "Остаток: " & FormatNumber(total - paid, 2) & " руб." & vbNewLine & _
                     "----------------------------------------" & vbNewLine
            
            ' Если есть долг, добавляем в комбобоксы
            If paid < total Then
                Dim serviceInfo As String
                serviceInfo = ws.Cells(i, 3).value & " (Долг: " & FormatNumber(total - paid, 2) & " руб.)"
                
                ' Проверяем на дубликаты перед добавлением
                If Not IsInComboBox(cmbУслуги, serviceInfo) Then
                    cmbУслуги.AddItem serviceInfo
                    cmbГРУППЫ.AddItem ws.Cells(i, 2).value
                End If
                
                ' Сохраняем номера строк с долгами
                foundRows = foundRows & i & ","
            End If
        End If
    Next i
    
    ' Выводим результаты
    If found Then
        txtИсторияЗаказа.text = history
    Else
        txtИсторияЗаказа.text = "Заказы не найдены для: " & txtФИО.text
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "Ошибка при поиске: " & vbNewLine & _
           "Описание: " & Err.description & vbNewLine & _
           "Код ошибки: " & Err.Number, vbCritical
End Sub

Private Sub btnОплатить_Click()
    ' Проверки перед оплатой
    If foundFIO = "" Then
        MsgBox "Сначала выполните поиск клиента!", vbExclamation
        Exit Sub
    End If
    
    If cmbУслуги.text = "" Then
        MsgBox "Выберите услугу для оплаты!", vbExclamation
        cmbУслуги.SetFocus
        Exit Sub
    End If
    
    If Trim(txtОплата.text) = "" Then
        MsgBox "Введите сумму оплаты!", vbExclamation
        txtОплата.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtОплата.text) Then
        MsgBox "Сумма оплаты должна быть числом!", vbExclamation
        txtОплата.SetFocus
        Exit Sub
    End If
    
    If CDbl(txtОплата.text) <= 0 Then
        MsgBox "Сумма оплаты должна быть больше нуля!", vbExclamation
        txtОплата.SetFocus
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Заказы")
    
    ' Получаем выбранную услугу
    Dim selectedService As String
    selectedService = Split(cmbУслуги.text, " (")(0)
    
    ' Находим соответствующую строку
    Dim rowNumbers() As String
    rowNumbers = Split(foundRows, ",")
    Dim i As Long
    Dim found As Boolean
    found = False
    
    For i = 0 To UBound(rowNumbers) - 1
        Dim rowNum As Long
        rowNum = CLng(rowNumbers(i))
        
        If ws.Cells(rowNum, 3).value = selectedService Then
            ' Проверяем сумму оплаты
            Dim currentTotal As Double
            Dim currentPaid As Double
            Dim newPayment As Double
            
            currentTotal = CDbl(ws.Cells(rowNum, 10).value)
            currentPaid = IIf(ws.Cells(rowNum, 11).value = "", 0, CDbl(ws.Cells(rowNum, 11).value))
            newPayment = CDbl(txtОплата.text)
            
            ' Проверяем, не превышает ли оплата остаток
            If (currentPaid + newPayment) > currentTotal Then
                MsgBox "Сумма оплаты превышает остаток!" & vbNewLine & _
                       "Максимальная сумма к оплате: " & FormatNumber(currentTotal - currentPaid, 2) & " руб.", vbExclamation
                Exit Sub
            End If
            
            ' Обновляем данные
            ws.Cells(rowNum, 11).value = currentPaid + newPayment
            ws.Cells(rowNum, 12).value = currentTotal - (currentPaid + newPayment)
            
            ' Если заказ полностью оплачен, выделяем зеленым
            If (currentPaid + newPayment) >= currentTotal Then
                ws.Cells(rowNum, 1).Interior.Color = RGB(198, 239, 206)
            End If
            
            found = True
            
            ' Показываем сообщение об успешной оплате
            MsgBox "Оплата успешно добавлена!" & vbNewLine & _
                   "Клиент: " & ws.Cells(rowNum, 1).value & vbNewLine & _
                   "Услуга: " & selectedService & vbNewLine & _
                   "Сумма оплаты: " & FormatNumber(newPayment, 2) & " руб." & vbNewLine & _
                   "Остаток: " & FormatNumber(ws.Cells(rowNum, 12).value, 2) & " руб.", vbInformation
            
            ' Обновляем данные на форме
            btnПроверить_Click
            txtОплата.text = ""
            Exit For
        End If
    Next i
    
    If Not found Then
        MsgBox "Не удалось найти выбранную услугу!", vbExclamation
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "Ошибка при добавлении оплаты:" & vbNewLine & _
           "Описание: " & Err.description & vbNewLine & _
           "Код ошибки: " & Err.Number, vbCritical
End Sub



' Функция проверки наличия значения в комбобоксе
Private Function IsInComboBox(cmb As MSForms.ComboBox, value As String) As Boolean
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If cmb.list(i) = value Then
            IsInComboBox = True
            Exit Function
        End If
    Next i
    IsInComboBox = False
End Function

' Проверка ввода только цифр и точки в поле оплаты
Private Sub txtОПЛАТА_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 48 To 57  ' Цифры
        Case 46       ' Точка
            ' Проверяем, есть ли уже точка
            If InStr(1, txtОплата.text, ".") > 0 Then KeyAscii = 0
        Case 8        ' Backspace
        Case Else
            KeyAscii = 0
    End Select
End Sub

