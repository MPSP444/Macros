Attribute VB_Name = "СоздатьСводкуПрепаратов"
Sub СводкаСДетальнойПроверкой()
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim препараты As String
    Dim значение As Double
    Dim найденоВсего As Integer
    Dim отчет As String
    Dim листНайден As Boolean
    
    ' ИСПОЛЬЗУЕМ АКТИВНУЮ КНИГУ И АКТИВНЫЙ ЛИСТ
    Set wsSource = activeSheet
    
    If wsSource.Name = "Сводка" Then
        MsgBox "Сначала перейдите на лист с данными!", vbExclamation
        Exit Sub
    End If
    
    ' Ищем или создаем лист Сводка
    листНайден = False
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name = "Сводка" Then
            Set wsDest = ws
            листНайден = True
            Exit For
        End If
    Next ws
    
    If Not листНайден Then
        Set wsDest = ActiveWorkbook.Worksheets.Add
        wsDest.Name = "Сводка"
        
        ' Заголовки
        wsDest.Cells(1, 1).value = "Ангар"
        wsDest.Cells(1, 2).value = "Препараты в наличии"
        wsDest.Cells(1, 3).value = "Количество"
        wsDest.Cells(1, 4).value = "Детали (название: количество)"
        wsDest.Range("A1:D1").Font.Bold = True
        
        wsDest.Cells(2, 1).value = "5 Ангар ШТ+ШЗ"
        wsDest.Cells(3, 1).value = "6 Ангар ШТ+ШЗ"
        wsDest.Cells(4, 1).value = "7 Ангар ШТ+ШЗ"
        wsDest.Cells(5, 1).value = "8 Ангар ШТ+ШЗ"
        wsDest.Cells(6, 1).value = "9 Ангар ШТ+ШЗ"
        wsDest.Cells(7, 1).value = "10 Ангар ШТ+ШЗ"
        wsDest.Cells(8, 1).value = "11 Ангар ШТ+ШЗ"
        wsDest.Cells(9, 1).value = "12 Ангар ШТ+ШЗ"
        wsDest.Cells(10, 1).value = "10 Ангар маленький"
    End If
    
    ' Находим последнюю строку с данными в столбце A
    lastRow = wsSource.Cells(wsSource.rows.count, 1).End(xlUp).row
    
    ' Проверка диапазона
    MsgBox "Будем проверять строки с 2 по " & lastRow & vbCrLf & _
           "Всего строк для анализа: " & (lastRow - 1), vbInformation
    
    ' Очищаем старые данные
    wsDest.Range("B2:D10").ClearContents
    
    найденоВсего = 0
    отчет = "ДЕТАЛЬНЫЙ АНАЛИЗ:" & vbCrLf & vbCrLf
    
    ' Перебираем каждый ангар (столбцы B-J)
    For j = 2 To 10
        препараты = ""
        Dim детали As String
        детали = ""
        Dim счетчик As Integer
        счетчик = 0
        
        ' Название ангара
        Dim ангарНазвание As String
        Select Case j
            Case 2: ангарНазвание = "5 Ангар"
            Case 3: ангарНазвание = "6 Ангар"
            Case 4: ангарНазвание = "7 Ангар"
            Case 5: ангарНазвание = "8 Ангар"
            Case 6: ангарНазвание = "9 Ангар"
            Case 7: ангарНазвание = "10 Ангар"
            Case 8: ангарНазвание = "11 Ангар"
            Case 9: ангарНазвание = "12 Ангар"
            Case 10: ангарНазвание = "10 маленький"
        End Select
        
        ' Перебираем ВСЕ строки
        For i = 2 To lastRow
            ' Получаем название препарата
            Dim названиеПрепарата As String
            названиеПрепарата = Trim(CStr(wsSource.Cells(i, 1).value))
            
            ' Пропускаем пустые названия
            If названиеПрепарата <> "" And названиеПрепарата <> "0" Then
                ' Получаем значение из столбца ангара
                Dim cellValue As Variant
                cellValue = wsSource.Cells(i, j).value
                
                ' Проверяем что ячейка не пустая
                If Not isEmpty(cellValue) And CStr(cellValue) <> "" Then
                    ' Преобразуем в строку и чистим
                    Dim cleanValue As String
                    cleanValue = Replace(Replace(CStr(cellValue), " ", ""), ",", ".")
                    
                    ' Проверяем что это число
                    If IsNumeric(cleanValue) Then
                        значение = CDbl(cleanValue)
                        
                        ' Если значение больше 0
                        If значение > 0 Then
                            счетчик = счетчик + 1
                            найденоВсего = найденоВсего + 1
                            
                            ' Добавляем в список препаратов
                            If препараты = "" Then
                                препараты = названиеПрепарата
                            Else
                                препараты = препараты & ", " & названиеПрепарата
                            End If
                            
                            ' Добавляем в детали
                            If детали = "" Then
                                детали = названиеПрепарата & ": " & Format(значение, "#,##0.00")
                            Else
                                детали = детали & " | " & названиеПрепарата & ": " & Format(значение, "#,##0.00")
                            End If
                        End If
                    End If
                End If
            End If
        Next i
        
        ' Записываем результаты
        If препараты = "" Then
            wsDest.Cells(j, 2).value = "Пусто"
            wsDest.Cells(j, 3).value = 0
            wsDest.Cells(j, 4).value = "-"
            отчет = отчет & ангарНазвание & ": ПУСТО" & vbCrLf
        Else
            wsDest.Cells(j, 2).value = препараты
            wsDest.Cells(j, 3).value = счетчик
            wsDest.Cells(j, 4).value = детали
            отчет = отчет & ангарНазвание & ": " & счетчик & " препаратов" & vbCrLf
        End If
    Next j
    
    ' Форматируем столбцы
    With wsDest
        .columns("A").AutoFit
        .columns("B").ColumnWidth = 50
        .columns("B").WrapText = True
        .columns("C").AutoFit
        .columns("D").ColumnWidth = 60
        .columns("D").WrapText = True
    End With
    
    ' Показываем отчет
    отчет = отчет & vbCrLf & "Всего найдено препаратов: " & найденоВсего
    MsgBox отчет, vbInformation, "Результаты анализа"
    
    ' Переходим на лист Сводка
    wsDest.Activate
    
End Sub

