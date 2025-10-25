Attribute VB_Name = "formGroupModule2"
' ===== MODULE2 - С ОТЛАДКОЙ В КОНСОЛИ =====
' Вставьте этот код в Module2
' ВАЖНО: Откройте консоль (Ctrl+G) чтобы видеть отладку!

Sub ProcessWarehouseData(inputText As String)
    ' Словарь для хранения данных по ангарам
    Dim warehouses As Object
    Set warehouses = CreateObject("Scripting.Dictionary")
    
    ' Счетчики для статистики
    Dim totalLines As Long
    Dim processedLines As Long
    Dim skippedLines As Long
    
    Debug.Print "===== НАЧАЛО ОБРАБОТКИ ====="
    Debug.Print ""
    
    ' Разбираем входные данные
    Dim lines() As String
    lines = Split(inputText, vbCrLf)
    totalLines = UBound(lines) - LBound(lines) + 1
    
    Debug.Print "Всего строк для обработки: " & totalLines
    Debug.Print ""
    
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        On Error Resume Next ' Продолжаем даже при ошибках
        
        Dim line As String
        line = lines(i)
        
        Debug.Print "--- Строка " & (i + 1) & " ---"
        Debug.Print "Содержимое: [" & line & "]"
        Debug.Print "Длина: " & Len(line)
        
        ' Проверяем пустую строку
        If Len(Trim(line)) = 0 Then
            Debug.Print "ПРОПУЩЕНО: Пустая строка"
            Debug.Print ""
            skippedLines = skippedLines + 1
            GoTo NextLine
        End If
        
        ' Пропускаем заголовки
        If InStr(line, "Адрес") > 0 Or _
           InStr(line, "Номенклатура") > 0 Or _
           InStr(line, "Количество") > 0 Or _
           InStr(line, "Конечный остаток") > 0 Then
            Debug.Print "ПРОПУЩЕНО: Заголовок"
            Debug.Print ""
            skippedLines = skippedLines + 1
            GoTo NextLine
        End If
        
        ' Парсим строку по табуляции
        Dim parts() As String
        parts = Split(line, vbTab)
        
        Debug.Print "Частей после split по TAB: " & (UBound(parts) + 1)
        
        Dim j As Integer
        For j = 0 To UBound(parts)
            Debug.Print "  Часть " & (j + 1) & ": [" & Trim(parts(j)) & "]"
        Next j
        
        ' Проверяем минимум 2 части
        If UBound(parts) < 1 Then
            Debug.Print "ПРОПУЩЕНО: Недостаточно колонок"
            Debug.Print ""
            skippedLines = skippedLines + 1
            GoTo NextLine
        End If
        
        Dim address As String
        Dim warehouse As String
        Dim product As String
        Dim quantity As Double
        Dim foundQty As Boolean
        
        address = Trim(parts(0))
        
        ' ВАЖНО: Пропускаем строки с "Приемка" в адресе
        If InStr(1, address, "Приемка", vbTextCompare) > 0 Or _
           InStr(1, address, "приемка", vbTextCompare) > 0 Or _
           InStr(1, address, "ПРИЕМКА", vbTextCompare) > 0 Then
            Debug.Print "ПРОПУЩЕНО: Адрес содержит 'Приемка' - [" & address & "]"
            Debug.Print ""
            skippedLines = skippedLines + 1
            GoTo NextLine
        End If
        
        ' Проверяем наличие адреса
        If address = "" Then
            Debug.Print "ПРОПУЩЕНО: Нет адреса"
            Debug.Print ""
            skippedLines = skippedLines + 1
            GoTo NextLine
        End If
        
        ' Извлекаем номер ангара
        Dim dashPos As Integer
        dashPos = InStr(address, "-")
        If dashPos > 0 Then
            warehouse = Left(address, dashPos - 1)
        Else
            warehouse = address
        End If
        
        ' ВАЖНО: Пропускаем "Приемка"
        If InStr(1, warehouse, "Приемка", vbTextCompare) > 0 Or _
           InStr(1, warehouse, "приемка", vbTextCompare) > 0 Then
            Debug.Print "ПРОПУЩЕНО: Адрес содержит 'Приемка'"
            Debug.Print ""
            skippedLines = skippedLines + 1
            GoTo NextLine
        End If
        
        Debug.Print "Адрес: " & address
        Debug.Print "Ангар: " & warehouse
        
        ' Получаем товар
        If UBound(parts) >= 1 Then
            product = Trim(parts(1))
        Else
            product = ""
        End If
        
        If product = "" Then
            Debug.Print "ПРОПУЩЕНО: Нет названия товара"
            Debug.Print ""
            skippedLines = skippedLines + 1
            GoTo NextLine
        End If
        
        Debug.Print "Товар: " & product
        
        ' Ищем количество в любой из оставшихся колонок
        foundQty = False
        For j = UBound(parts) To 2 Step -1
            Dim qtyStr As String
            qtyStr = Trim(parts(j))
            
            ' ВАЖНО: Убираем ВСЕ типы пробелов (обычные, неразрывные и т.д.)
            qtyStr = Replace(qtyStr, " ", "")           ' Обычный пробел (32)
            qtyStr = Replace(qtyStr, Chr(160), "")      ' Неразрывный пробел (NBSP)
            qtyStr = Replace(qtyStr, Chr(9), "")        ' Табуляция
            qtyStr = Replace(qtyStr, vbTab, "")         ' Табуляция VBA
            ' Заменяем запятую на точку
            qtyStr = Replace(qtyStr, ",", ".")
            
            ' Пробуем преобразовать Val() - он более толерантен
            On Error Resume Next
            Dim testVal As Double
            testVal = Val(qtyStr)
            On Error GoTo 0
            
            Debug.Print "  Проверяем колонку " & (j + 1) & ": исходное=[" & Trim(parts(j)) & "], после обработки=[" & qtyStr & "], Val()=" & testVal
            
            If qtyStr <> "" And testVal > 0 Then
                quantity = testVal
                foundQty = True
                Debug.Print "? Количество найдено в колонке " & (j + 1) & ": " & quantity
                Exit For
            End If
        Next j
        
        ' Если не нашли в конце, проверяем 3-ю колонку
        If Not foundQty And UBound(parts) >= 2 Then
            qtyStr = Trim(parts(2))
            qtyStr = Replace(qtyStr, " ", "")
            qtyStr = Replace(qtyStr, Chr(160), "")
            qtyStr = Replace(qtyStr, Chr(9), "")
            qtyStr = Replace(qtyStr, vbTab, "")
            qtyStr = Replace(qtyStr, ",", ".")
            
            On Error Resume Next
            testVal = Val(qtyStr)
            On Error GoTo 0
            
            Debug.Print "  Проверяем колонку 3 (запасной вариант): [" & qtyStr & "], Val()=" & testVal
            If qtyStr <> "" And testVal > 0 Then
                quantity = testVal
                foundQty = True
                Debug.Print "? Количество найдено в колонке 3: " & quantity
            End If
        End If
        
        If Not foundQty Then
            Debug.Print "ПРОПУЩЕНО: Не найдено количество"
            Debug.Print ""
            skippedLines = skippedLines + 1
            GoTo NextLine
        End If
        
        ' Добавляем в словарь
        If Not warehouses.exists(warehouse) Then
            Set warehouses(warehouse) = CreateObject("Scripting.Dictionary")
            Debug.Print "Создан новый ангар: " & warehouse
        End If
        
        If warehouses(warehouse).exists(product) Then
            Dim oldQty As Double
            oldQty = warehouses(warehouse)(product)
            warehouses(warehouse)(product) = oldQty + quantity
            Debug.Print "Товар уже есть, добавляем: " & oldQty & " + " & quantity & " = " & warehouses(warehouse)(product)
        Else
            warehouses(warehouse)(product) = quantity
            Debug.Print "Добавлен новый товар"
        End If
        
        Debug.Print "? ОБРАБОТАНО УСПЕШНО"
        Debug.Print ""
        processedLines = processedLines + 1
        
NextLine:
        On Error GoTo 0
    Next i
    
    Debug.Print "===== ИТОГИ ====="
    Debug.Print "Всего строк: " & totalLines
    Debug.Print "Обработано: " & processedLines
    Debug.Print "Пропущено: " & skippedLines
    Debug.Print "Ангаров создано: " & warehouses.count
    Debug.Print ""
    
    ' Проверяем результат
    If warehouses.count = 0 Then
        MsgBox "Не найдено ни одного товара для обработки!" & vbCrLf & _
               "Откройте консоль (Ctrl+G) для просмотра отладки.", vbExclamation
        Exit Sub
    End If
    
    ' Создаем листы
    Dim whKey As Variant
    For Each whKey In warehouses.keys
        Debug.Print "Создаем лист для ангара: " & whKey
        CreateWarehouseSheet CStr(whKey), warehouses(whKey)
    Next whKey
    
    Debug.Print ""
    Debug.Print "===== ЗАВЕРШЕНО ====="
    
    MsgBox "Обработка завершена!" & vbCrLf & _
           "Обработано строк: " & processedLines & vbCrLf & _
           "Пропущено строк: " & skippedLines & vbCrLf & _
           "Создано листов: " & warehouses.count & vbCrLf & vbCrLf & _
           "Подробности в консоли (Ctrl+G)", vbInformation
End Sub

Sub CreateWarehouseSheet(warehouseNum As String, products As Object)
    Dim ws As Worksheet
    Dim sheetName As String
    sheetName = "Ангар " & warehouseNum
    
    ' Проверяем, существует ли лист
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    ' Если лист существует, очищаем его, иначе создаем новый
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If
    
    ' Создаем заголовок
    With ws.Range("A1:F1")
        .Merge
        .value = sheetName
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(200, 200, 200)
    End With
    
    ' Создаем заголовки колонок
    ws.Range("A2").value = "Наименовании"
    ws.Range("B2").value = "Общее кол-во"
    ws.Range("C2").value = "Хранении"
    ws.Range("D2").value = "Луч"
    ws.Range("E2").value = "ФЭС"
    ws.Range("F2").value = "ФАКТ"
    
    With ws.Range("A2:F2")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(220, 220, 220)
        .Borders.Weight = xlThin
    End With
    
    ' Заполняем данные
    Dim row As Long
    row = 3
    
    Dim prodKey As Variant
    For Each prodKey In products.keys
        ws.Cells(row, 1).value = prodKey
        ws.Cells(row, 2).value = products(prodKey)
        ws.Cells(row, 2).NumberFormat = "0.000"
        row = row + 1
    Next prodKey
    
    ' Форматируем таблицу
    Dim lastRow As Long
    lastRow = row - 1
    
    If lastRow >= 3 Then
        With ws.Range("A2:F" & lastRow)
            .Borders.Weight = xlThin
        End With
    End If
    
    ' Настраиваем ширину колонок
    ws.columns("A:A").ColumnWidth = 35
    ws.columns("B:F").ColumnWidth = 15
    
    ' Автоподбор высоты строк
    ws.rows("1:" & lastRow).AutoFit
    
    ' Активируем созданный лист
    ws.Activate
End Sub

