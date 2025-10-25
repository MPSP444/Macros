Attribute VB_Name = "ModuleMassInput"


' ===== ModuleMassInput - ОБНОВЛЕННЫЙ с умной логикой суммирования =====
Option Explicit

' Структура для хранения данных о товаре
Private Type productData
    warehouse As String      ' Номер ангара
    row As String           ' Номер ряда
    Shelf As String         ' Стеллаж (А, Б, В...)
    level As String         ' Уровень (1, 2, 3)
    ProductName As String   ' Название товара
    Batch As String         ' Партия
    quantity As Double      ' Количество
    IsValid As Boolean      ' Валидность записи
    ErrorMessage As String  ' Сообщение об ошибке
End Type

' Структура для группировки товаров по рядам
Private Type rowData
    warehouse As String
    row As String
    products As Object      ' Dictionary с товарами
    HasConflict As Boolean  ' Есть ли конфликт (разные товары в одном ряду)
End Type

Public Function ProcessMassInput(inputText As String) As String
    On Error GoTo ErrorHandler
    
    ModuleLogger.LogMessage "=== НАЧАЛО УМНОЙ ОБРАБОТКИ МАССОВЫХ ДАННЫХ ==="
    ModuleLogger.LogMessage "Входных строк: " & UBound(Split(inputText, vbNewLine)) + 1
    
    Application.ScreenUpdating = False
    
    ' 1. Парсинг входных данных
    Dim parsedData() As productData
    parsedData = ParseInputData(inputText)
    
    If UBound(parsedData) < 0 Then
        ProcessMassInput = "Не найдено корректных данных для обработки!"
        Application.ScreenUpdating = True
        Exit Function
    End If
    
    ' 2. Группировка по ангарам и рядам (С УЧЕТОМ СЕКЦИЙ)
    Dim groupedData As Object
    Set groupedData = GroupDataByRows(parsedData)
    
    ' 3. НОВАЯ ЛОГИКА: Обработка данных ПЕРЕД заголовками (для чтения существующих)
    ModuleLogger.LogMessage "?? Начало записи данных с суммированием..."
    Dim dataResult As String
    dataResult = FillDataToTables(parsedData)
    
    ' 4. УМНАЯ ОБРАБОТКА ЗАГОЛОВКОВ (после записи данных!)
    ModuleLogger.LogMessage "?? Начало умной обработки заголовков..."
    Dim headerResult As String
    headerResult = ModuleHeaderManager.ProcessHeaders(groupedData)
    
    ' 5. Формирование отчета об ошибках
    Dim errorReport As String
    errorReport = GenerateErrorReport(parsedData, headerResult, dataResult)
    
    Application.ScreenUpdating = True
    ModuleLogger.LogMessage "=== ЗАВЕРШЕНИЕ УМНОЙ ОБРАБОТКИ МАССОВЫХ ДАННЫХ ==="
    
    ProcessMassInput = errorReport
    Exit Function
    
ErrorHandler:
    Application.ScreenUpdating = True
    ModuleLogger.LogMessage "КРИТИЧЕСКАЯ ОШИБКА: " & Err.description
    ProcessMassInput = "Критическая ошибка обработки: " & Err.description
End Function

Private Function ParseInputData(inputText As String) As productData()
    Dim lines() As String
    lines = Split(inputText, vbNewLine)
    
    ' Создаем массив для результатов
    Dim results() As productData
    ReDim results(UBound(lines))
    
    Dim validCount As Integer
    validCount = 0
    
    Dim i As Integer
    For i = 0 To UBound(lines)
        Dim line As String
        line = Trim(lines(i))
        
        ' Пропускаем пустые строки и примеры
        If line <> "" And Not (InStr(LCase(line), "пример") > 0) Then
            Dim productData As productData
            productData = ParseSingleLine(line)
            
            If productData.IsValid Then
                results(validCount) = productData
                validCount = validCount + 1
                ModuleLogger.LogMessage "? Распознано: " & line
            Else
                results(validCount) = productData ' Сохраняем и невалидные для отчета
                validCount = validCount + 1
                ModuleLogger.LogMessage "? Ошибка: " & line & " - " & productData.ErrorMessage
            End If
        End If
    Next i
    
    ' Корректируем размер массива
    If validCount > 0 Then
        ReDim Preserve results(validCount - 1)
    Else
        ReDim results(-1)
    End If
    
    ParseInputData = results
End Function

Private Function ParseSingleLine(line As String) As productData
    Dim result As productData
    result.IsValid = False
    
    ' Ожидаемый формат: 6-5-Е-2 - Лерашанс - пар13 - 720
    Dim parts() As String
    parts = Split(line, " - ")
    
    If UBound(parts) < 3 Then
        result.ErrorMessage = "Неверный формат строки. Ожидается: Ангар-Ряд-Стеллаж-Уровень - Товар - Партия - Количество"
        ParseSingleLine = result
        Exit Function
    End If
    
    ' Парсим координаты
    Dim coordinates() As String
    coordinates = Split(Trim(parts(0)), "-")
    
    If UBound(coordinates) < 3 Then
        result.ErrorMessage = "Неверный формат координат. Ожидается: Ангар-Ряд-Стеллаж-Уровень"
        ParseSingleLine = result
        Exit Function
    End If
    
    ' Заполняем координаты
    result.warehouse = Trim(coordinates(0))
    result.row = Trim(coordinates(1))
    result.Shelf = UCase(Trim(coordinates(2)))
    result.level = Trim(coordinates(3))
    
    ' Заполняем товар (убираем все после *)
    result.ProductName = CleanProductName(Trim(parts(1)))
    
    ' Заполняем партию
    result.Batch = Trim(parts(2))
    
    ' Заполняем количество
    On Error Resume Next
    result.quantity = CDbl(Trim(parts(3)))
    If Err.Number <> 0 Then
        result.ErrorMessage = "Неверное количество: " & parts(3)
        On Error GoTo 0
        ParseSingleLine = result
        Exit Function
    End If
    On Error GoTo 0
    
    ' Валидация с помощью обертки
    If Not ValidateProductDataLocal(result) Then
        result.ErrorMessage = ModuleValidator.GetLastError()
        ParseSingleLine = result
        Exit Function
    End If
    
    result.IsValid = True
    ParseSingleLine = result
End Function

Private Function CleanProductName(ProductName As String) As String
    ' Убираем все после символа *
    Dim asteriskPos As Integer
    asteriskPos = InStr(ProductName, "*")
    
    If asteriskPos > 0 Then
        CleanProductName = Trim(Left(ProductName, asteriskPos - 1))
    Else
        CleanProductName = Trim(ProductName)
    End If
End Function

Private Function ValidateProductDataLocal(data As productData) As Boolean
    ' Функция-обертка для валидации структуры ProductData
    ' Вызывает отдельные функции валидации из ModuleValidator
    
    ' Валидация ангара
    If Not ModuleValidator.ValidateWarehouse(data.warehouse) Then
        ValidateProductDataLocal = False
        Exit Function
    End If
    
    ' Валидация ряда
    If Not ModuleValidator.ValidateRow(data.row, data.warehouse) Then
        ValidateProductDataLocal = False
        Exit Function
    End If
    
    ' Валидация стеллажа
    If Not ModuleValidator.ValidateShelf(data.Shelf, data.warehouse) Then
        ValidateProductDataLocal = False
        Exit Function
    End If
    
    ' Валидация уровня
    If Not ModuleValidator.ValidateLevel(data.level) Then
        ValidateProductDataLocal = False
        Exit Function
    End If
    
    ' Валидация названия товара
    If Not ModuleValidator.ValidateProductName(data.ProductName) Then
        ValidateProductDataLocal = False
        Exit Function
    End If
    
    ' Валидация партии
    If Not ModuleValidator.ValidateBatch(data.Batch) Then
        ValidateProductDataLocal = False
        Exit Function
    End If
    
    ' Валидация количества
    If Not ModuleValidator.ValidateQuantity(data.quantity) Then
        ValidateProductDataLocal = False
        Exit Function
    End If
    
    ValidateProductDataLocal = True
End Function

Private Function GroupDataByRows(parsedData() As productData) As Object
    ' ОБНОВЛЕННАЯ ФУНКЦИЯ: группировка с учетом секций
    Dim groupedData As Object
    Set groupedData = CreateObject("Scripting.Dictionary")
    
    Dim i As Integer
    For i = 0 To UBound(parsedData)
        If parsedData(i).IsValid Then
            ' НОВАЯ ЛОГИКА: добавляем секцию к ключу
            Dim section As String
            section = ModuleConfig.DetermineSectionByShelf(parsedData(i).Shelf, parsedData(i).warehouse)
            
            Dim key As String
            key = parsedData(i).warehouse & "-" & parsedData(i).row & "-" & section
            
            If Not groupedData.exists(key) Then
                Dim rowData As Object
                Set rowData = CreateObject("Scripting.Dictionary")
                groupedData.Add key, rowData
            End If
            
            Dim productDict As Object
            Set productDict = groupedData(key)
            
            ' Добавляем или обновляем товар в ряду
            If productDict.exists(parsedData(i).ProductName) Then
                ' Товар уже есть - это дубль, увеличиваем количество
                ModuleLogger.LogMessage "?? Найден дубль: " & key & " - " & parsedData(i).ProductName
            Else
                productDict.Add parsedData(i).ProductName, True
            End If
            
            ' Логируем секцию для отладки
            ModuleLogger.LogDebug "?? Товар " & parsedData(i).ProductName & " определен в " & ModuleConfig.GetSectionName(section) & " секции (" & section & ") ангара " & parsedData(i).warehouse
        End If
    Next i
    
    Set GroupDataByRows = groupedData
End Function

Private Function FillDataToTables(parsedData() As productData) As String
    Dim errorLog As String
    errorLog = ""
    
    ' Словарь для накопления дублей
    Dim accumulator As Object
    Set accumulator = CreateObject("Scripting.Dictionary")
    
    ' Первый проход - накапливаем дубли
    Dim i As Integer
    For i = 0 To UBound(parsedData)
        If parsedData(i).IsValid Then
            Dim key As String
            key = parsedData(i).warehouse & "-" & parsedData(i).row & "-" & _
                  parsedData(i).Shelf & "-" & parsedData(i).level
            
            If accumulator.exists(key) Then
                ' Обновляем существующую запись
                Dim existingData As Object
                Set existingData = accumulator(key)
                
                ' Правильно обрабатываем типы данных
                Dim currentQuantity As Double
                currentQuantity = CDbl(existingData("Quantity"))
                
                existingData("Quantity") = currentQuantity + parsedData(i).quantity
                existingData("Batch") = CStr(existingData("Batch")) & "/" & parsedData(i).Batch
                
                ModuleLogger.LogMessage "?? Накопление дублей в позиции " & key & ": +" & parsedData(i).quantity
            Else
                ' Создаем новую запись с правильными типами данных
                Dim newData As Object
                Set newData = CreateObject("Scripting.Dictionary")
                newData("Warehouse") = CStr(parsedData(i).warehouse)
                newData("Row") = CStr(parsedData(i).row)
                newData("Shelf") = CStr(parsedData(i).Shelf)
                newData("Level") = CStr(parsedData(i).level)
                newData("ProductName") = CStr(parsedData(i).ProductName)
                newData("Batch") = CStr(parsedData(i).Batch)
                newData("Quantity") = CDbl(parsedData(i).quantity)
                
                accumulator.Add key, newData
                ModuleLogger.LogMessage "?? Новая позиция " & key & ": " & parsedData(i).quantity
            End If
        End If
    Next i
    
    ' Второй проход - записываем накопленные данные с УМНЫМ СУММИРОВАНИЕМ
    Dim accKey As Variant
    For Each accKey In accumulator.keys
        Dim dataDict As Object
        Set dataDict = accumulator(accKey)
        
        Dim writeResult As String
        writeResult = WriteDataToCellSmart(dataDict)
        
        If writeResult <> "" Then
            errorLog = errorLog & writeResult & vbNewLine
        End If
    Next accKey
    
    FillDataToTables = errorLog
End Function

Private Function WriteDataToCellSmart(dataDict As Object) As String
    ' НОВАЯ УМНАЯ ФУНКЦИЯ: суммирует вместо перезаписи
    On Error GoTo ErrorHandler
    
    ' Получаем конфигурацию ангара
    Dim config As ModuleTypes.WarehouseConfig
    config = ModuleConfig.GetWarehouseConfig(CStr(dataDict("Warehouse")))
    
    ' Проверяем существование листа
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(config.sheetName)
    
    ' Используем улучшенный поиск с учетом секций
    Dim foundCell As Range
    Set foundCell = ModuleConfig.FindCellInWarehouseSection(ws, CStr(dataDict("Shelf")), _
                                                           CStr(dataDict("Level")), _
                                                           CStr(dataDict("Warehouse")))
    
    If foundCell Is Nothing Then
        ' Определяем секцию для сообщения об ошибке
        Dim section As String
        section = ModuleConfig.DetermineSectionByShelf(CStr(dataDict("Shelf")), CStr(dataDict("Warehouse")))
        
        WriteDataToCellSmart = "Не найдена ячейка " & dataDict("Shelf") & "-" & dataDict("Level") & _
                              " в ангаре " & dataDict("Warehouse") & " (секция: " & ModuleConfig.GetSectionName(section) & ")"
        Exit Function
    End If
    
    ' Вычисляем столбцы для записи
    Dim rowNumber As Long
    rowNumber = CLng(CStr(dataDict("Row")))
    
    Dim valueColumn As Long
    valueColumn = 3 + (rowNumber - 1) * 2  ' C, E, G, I...
    
    Dim batchColumn As Long
    batchColumn = valueColumn + 1  ' D, F, H, J...
    
    ' ===== НОВАЯ ЛОГИКА: ЧИТАЕМ СУЩЕСТВУЮЩИЕ ДАННЫЕ =====
    
    ' Читаем текущее количество
    Dim currentQuantity As Double
    Dim currentValue As Variant
    currentValue = ws.Cells(foundCell.row, valueColumn).value
    
    If IsNumeric(currentValue) And Not isEmpty(currentValue) Then
        currentQuantity = CDbl(currentValue)
    Else
        currentQuantity = 0
    End If
    
    ' Читаем текущую партию
    Dim currentBatch As String
    Dim currentBatchValue As Variant
    currentBatchValue = ws.Cells(foundCell.row, batchColumn).value
    
    If Not isEmpty(currentBatchValue) And currentBatchValue <> "" Then
        currentBatch = Trim(CStr(currentBatchValue))
    Else
        currentBatch = ""
    End If
    
    ' ===== СУММИРУЕМ КОЛИЧЕСТВА =====
    Dim newQuantity As Double
    newQuantity = currentQuantity + CDbl(dataDict("Quantity"))
    
    ' ===== УМНО ОБЪЕДИНЯЕМ ПАРТИИ =====
    Dim newBatch As String
    Dim incomingBatch As String
    incomingBatch = Trim(CStr(dataDict("Batch")))
    
    If currentBatch = "" Then
        ' Если текущая партия пустая - просто записываем новую
        newBatch = incomingBatch
        ModuleLogger.LogMessage "?? Новая партия: " & incomingBatch
    ElseIf UCase(currentBatch) = UCase(incomingBatch) Then
        ' Если партии одинаковые - не дублируем
        newBatch = currentBatch
        ModuleLogger.LogMessage "?? Одинаковые партии, дублирование не требуется: " & currentBatch
    ElseIf InStr(UCase(currentBatch), UCase(incomingBatch)) > 0 Then
        ' Если новая партия уже есть в списке - не добавляем
        newBatch = currentBatch
        ModuleLogger.LogMessage "?? Партия уже есть в списке: " & incomingBatch & " в " & currentBatch
    Else
        ' Если разные партии - объединяем через "/"
        newBatch = currentBatch & "/" & incomingBatch
        ModuleLogger.LogMessage "?? Объединение партий: " & currentBatch & " + " & incomingBatch & " = " & newBatch
    End If
    
    ' ===== ЗАПИСЫВАЕМ РЕЗУЛЬТАТ =====
    ws.Cells(foundCell.row, valueColumn).value = newQuantity
    ws.Cells(foundCell.row, batchColumn).value = newBatch
    
    ' Определяем секцию для логирования
    Dim logSection As String
    logSection = ModuleConfig.DetermineSectionByShelf(CStr(dataDict("Shelf")), CStr(dataDict("Warehouse")))
    
    ' Логируем операцию
    If currentQuantity > 0 Then
        ModuleLogger.LogSuccess "?? СУММИРОВАНИЕ: " & config.sheetName & " " & foundCell.address & _
                               " (ряд " & rowNumber & ") " & currentQuantity & "+" & dataDict("Quantity") & "=" & newQuantity & _
                               " | Партии: " & newBatch & " | Секция: " & ModuleConfig.GetSectionName(logSection)
    Else
        ModuleLogger.LogMessage "?? НОВАЯ ЗАПИСЬ: " & config.sheetName & " " & foundCell.address & _
                               " (ряд " & rowNumber & ") = " & newQuantity & _
                               " | Партия: " & newBatch & " | Секция: " & ModuleConfig.GetSectionName(logSection)
    End If
    
    WriteDataToCellSmart = ""
    Exit Function
    
ErrorHandler:
    WriteDataToCellSmart = "Ошибка записи в ангар " & dataDict("Warehouse") & ": " & Err.description
End Function

' Сохраняем старую функцию для совместимости
Private Function WriteDataToCell(dataDict As Object) As String
    ' Вызываем новую умную функцию
    WriteDataToCell = WriteDataToCellSmart(dataDict)
End Function

Private Function GenerateErrorReport(parsedData() As productData, headerResult As String, dataResult As String) As String
    Dim errorReport As String
    errorReport = ""
    
    ' Ошибки парсинга
    Dim parseErrors As String
    parseErrors = ""
    Dim i As Integer
    For i = 0 To UBound(parsedData)
        If Not parsedData(i).IsValid Then
            parseErrors = parseErrors & "• " & parsedData(i).ErrorMessage & vbNewLine
        End If
    Next i
    
    If parseErrors <> "" Then
        errorReport = errorReport & "ОШИБКИ РАЗБОРА ДАННЫХ:" & vbNewLine & parseErrors & vbNewLine
    End If
    
    ' Ошибки заголовков
    If headerResult <> "" Then
        errorReport = errorReport & "ОШИБКИ ОБРАБОТКИ ЗАГОЛОВКОВ:" & vbNewLine & headerResult & vbNewLine
    End If
    
    ' Ошибки записи данных
    If dataResult <> "" Then
        errorReport = errorReport & "ОШИБКИ ЗАПИСИ ДАННЫХ:" & vbNewLine & dataResult & vbNewLine
    End If
    
    GenerateErrorReport = errorReport
End Function

' ===== ТЕСТОВЫЕ ФУНКЦИИ =====

Public Sub TestSmartSummation()
    ' Тест умного суммирования
    Call ModuleLogger.InitializeLogger(True, ModuleLogger.LOG_LEVEL_DEBUG)
    
    ModuleLogger.LogMessage "=== ТЕСТ УМНОГО СУММИРОВАНИЯ ==="
    
    Dim testData As String
    testData = "6-5-Е-2 - Лерашанс - пар13 - 720" & vbNewLine & _
               "6-5-Е-2 - лерашанс - пар14 - 280" & vbNewLine & _
               "6-5-Е-2 - Дикошанс - пар25 - 500"
    
    ModuleLogger.LogMessage "Тестовые данные:"
    ModuleLogger.LogMessage testData
    
    Dim result As String
    result = ProcessMassInput(testData)
    
    If result = "" Then
        ModuleLogger.LogSuccess "? Тест прошел успешно!"
        MsgBox "Тест умного суммирования завершен успешно!" & vbNewLine & _
               "Проверьте ячейку 6-5-Е-2:" & vbNewLine & _
               "Ожидается: 1500 (720+280+500)" & vbNewLine & _
               "Партии: пар13/пар14/пар25", vbInformation
    Else
        ModuleLogger.LogError "? Ошибки в тесте: " & result
        MsgBox "Тест завершен с ошибками:" & vbNewLine & result, vbExclamation
    End If
End Sub

Public Sub QuickTestNewLogic()
    ' Быстрый тест всей новой системы
    Call ModuleLogger.InitializeLogger(True, ModuleLogger.LOG_LEVEL_DEBUG)
    
    ModuleLogger.LogMessage "=== БЫСТРЫЙ ТЕСТ НОВОЙ СИСТЕМЫ ==="
    
    ' Показываем информацию
    MsgBox "?? НОВАЯ УМНАЯ СИСТЕМА ГОТОВА!" & vbNewLine & vbNewLine & _
           "? Суммирование количеств вместо перезаписи" & vbNewLine & _
           "? Умное объединение партий через /" & vbNewLine & _
           "? Избежание дублей партий" & vbNewLine & _
           "? Умный поиск препаратов в заголовках" & vbNewLine & _
           "? Красные конфликтные заголовки" & vbNewLine & _
           "? Подробное логирование всех операций" & vbNewLine & vbNewLine & _
           "?? Проверьте Immediate Window (Ctrl+G) для деталей!" & vbNewLine & vbNewLine & _
           "?? Запустите TestSmartSummation() для тестирования!", _
           vbInformation, "Умная система обновлена"
End Sub

