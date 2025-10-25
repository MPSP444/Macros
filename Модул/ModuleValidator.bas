Attribute VB_Name = "ModuleValidator"

' ===== ModuleValidator - Модуль валидации данных =====
Option Explicit

Private lastError As String

' Списки допустимых значений
Private Const VALID_WAREHOUSES As String = "5,6,7,8,9,10,11,12"
Private Const VALID_LEVELS As String = "1,2,3"

' ===== ОСНОВНЫЕ ФУНКЦИИ ВАЛИДАЦИИ =====

Public Function ValidateProductData(data As Object) As Boolean
    ' Основная функция валидации структуры данных о товаре
    lastError = ""
    
    ' Получаем данные из структуры
    Dim warehouse As String
    Dim row As String
    Dim Shelf As String
    Dim level As String
    Dim ProductName As String
    Dim Batch As String
    Dim quantity As Double
    
    ' Извлекаем данные в зависимости от типа параметра
    On Error GoTo ErrorHandler
    
    ' Если это структура ProductData (VBA Type), обращаемся к полям напрямую
    warehouse = data.warehouse
    row = data.row
    Shelf = data.Shelf
    level = data.level
    ProductName = data.ProductName
    Batch = data.Batch
    quantity = data.quantity
    
    ' Валидация ангара
    If Not ValidateWarehouse(warehouse) Then
        ValidateProductData = False
        Exit Function
    End If
    
    ' Валидация ряда
    If Not ValidateRow(row, warehouse) Then
        ValidateProductData = False
        Exit Function
    End If
    
    ' Валидация стеллажа
    If Not ValidateShelf(Shelf, warehouse) Then
        ValidateProductData = False
        Exit Function
    End If
    
    ' Валидация уровня
    If Not ValidateLevel(level) Then
        ValidateProductData = False
        Exit Function
    End If
    
    ' Валидация названия товара
    If Not ValidateProductName(ProductName) Then
        ValidateProductData = False
        Exit Function
    End If
    
    ' Валидация партии
    If Not ValidateBatch(Batch) Then
        ValidateProductData = False
        Exit Function
    End If
    
    ' Валидация количества
    If Not ValidateQuantity(quantity) Then
        ValidateProductData = False
        Exit Function
    End If
    
    ValidateProductData = True
    Exit Function
    
ErrorHandler:
    lastError = "Ошибка валидации структуры данных: " & Err.description
    ValidateProductData = False
End Function

Public Function GetLastError() As String
    ' Возвращает последнюю ошибку валидации
    GetLastError = lastError
End Function

' ===== ВАЛИДАЦИЯ ОТДЕЛЬНЫХ ПОЛЕЙ =====

Public Function ValidateWarehouse(warehouse As String) As Boolean
    ' Валидация номера ангара
    If warehouse = "" Then
        lastError = "Не указан номер ангара"
        ValidateWarehouse = False
        Exit Function
    End If
    
    If Not IsNumeric(warehouse) Then
        lastError = "Номер ангара должен быть числом: " & warehouse
        ValidateWarehouse = False
        Exit Function
    End If
    
    If InStr("," & VALID_WAREHOUSES & ",", "," & warehouse & ",") = 0 Then
        lastError = "Несуществующий ангар: " & warehouse & ". Допустимые: " & VALID_WAREHOUSES
        ValidateWarehouse = False
        Exit Function
    End If
    
    ' Дополнительная проверка - существует ли лист ангара
    If Not ValidateWarehouseExists(warehouse) Then
        ValidateWarehouse = False
        Exit Function
    End If
    
    ValidateWarehouse = True
End Function

Public Function ValidateRow(row As String, warehouse As String) As Boolean
    ' Валидация номера ряда
    If row = "" Then
        lastError = "Не указан номер ряда"
        ValidateRow = False
        Exit Function
    End If
    
    If Not IsNumeric(row) Then
        lastError = "Номер ряда должен быть числом: " & row
        ValidateRow = False
        Exit Function
    End If
    
    Dim rowNum As Integer
    rowNum = CInt(row)
    
    If rowNum < 1 Then
        lastError = "Номер ряда должен быть больше 0: " & row
        ValidateRow = False
        Exit Function
    End If
    
    ' Проверяем максимальное количество рядов для каждого ангара
    Dim maxRows As Integer
    maxRows = GetMaxRowsForWarehouse(warehouse)
    
    If rowNum > maxRows Then
        lastError = "Ряд " & row & " не существует в ангаре " & warehouse & ". Максимум рядов: " & maxRows
        ValidateRow = False
        Exit Function
    End If
    
    ValidateRow = True
End Function

Public Function ValidateShelf(Shelf As String, warehouse As String) As Boolean
    ' Валидация стеллажа
    If Shelf = "" Then
        lastError = "Не указан стеллаж"
        ValidateShelf = False
        Exit Function
    End If
    
    If Len(Shelf) > 3 Then
        lastError = "Слишком длинное название стеллажа: " & Shelf
        ValidateShelf = False
        Exit Function
    End If
    
    ' Проверяем, что стеллаж состоит из допустимых символов
    If Not IsValidShelfName(Shelf) Then
        lastError = "Недопустимое название стеллажа: " & Shelf & ". Используйте русские буквы А-Я или ПР-коды"
        ValidateShelf = False
        Exit Function
    End If
    
    ' Проверяем, что стеллаж существует в данном ангаре
    If Not ValidateShelfInWarehouse(Shelf, warehouse) Then
        ValidateShelf = False
        Exit Function
    End If
    
    ValidateShelf = True
End Function

Public Function ValidateLevel(level As String) As Boolean
    ' Валидация уровня стеллажа
    If level = "" Then
        lastError = "Не указан уровень"
        ValidateLevel = False
        Exit Function
    End If
    
    If Not IsNumeric(level) Then
        lastError = "Уровень должен быть числом: " & level
        ValidateLevel = False
        Exit Function
    End If
    
    If InStr("," & VALID_LEVELS & ",", "," & level & ",") = 0 Then
        lastError = "Несуществующий уровень: " & level & ". Допустимые: " & VALID_LEVELS
        ValidateLevel = False
        Exit Function
    End If
    
    ValidateLevel = True
End Function

Public Function ValidateProductName(ProductName As String) As Boolean
    ' Валидация названия товара
    If ProductName = "" Then
        lastError = "Не указано название товара"
        ValidateProductName = False
        Exit Function
    End If
    
    If Len(Trim(ProductName)) < 2 Then
        lastError = "Название товара слишком короткое: " & ProductName
        ValidateProductName = False
        Exit Function
    End If
    
    If Len(ProductName) > 50 Then
        lastError = "Слишком длинное название товара: " & ProductName
        ValidateProductName = False
        Exit Function
    End If
    
    ' Проверяем на недопустимые символы
    If InStr(ProductName, Chr(10)) > 0 Or InStr(ProductName, Chr(13)) > 0 Then
        lastError = "Название товара содержит недопустимые символы переноса строки"
        ValidateProductName = False
        Exit Function
    End If
    
    ' Проверяем на недопустимые символы
    If InStr(ProductName, """") > 0 Or InStr(ProductName, "'") > 0 Then
        lastError = "Название товара содержит недопустимые кавычки: " & ProductName
        ValidateProductName = False
        Exit Function
    End If
    
    ValidateProductName = True
End Function

Public Function ValidateBatch(Batch As String) As Boolean
    ' Валидация партии (ИСПРАВЛЕННАЯ ВЕРСИЯ - разрешает буквы и цифры после "пар")
    If Batch = "" Then
        lastError = "Не указана партия"
        ValidateBatch = False
        Exit Function
    End If
    
    If Len(Batch) > 20 Then
        lastError = "Слишком длинная партия: " & Batch
        ValidateBatch = False
        Exit Function
    End If
    
    ' Проверяем формат партии (должна начинаться с "пар")
    If Left(LCase(Trim(Batch)), 3) <> "пар" Then
        lastError = "Партия должна начинаться с 'пар': " & Batch
        ValidateBatch = False
        Exit Function
    End If
    
    ' ИСПРАВЛЕНО: разрешаем буквы, цифры и другие символы после "пар"
    Dim suffixPart As String
    suffixPart = Mid(Batch, 4)
    
    ' Проверяем, что после "пар" есть хоть что-то
    If suffixPart = "" Then
        lastError = "После 'пар' должно быть указание номера/кода партии: " & Batch
        ValidateBatch = False
        Exit Function
    End If
    
    ' Запрещаем только опасные символы (переносы строк, кавычки)
    If InStr(suffixPart, Chr(10)) > 0 Or InStr(suffixPart, Chr(13)) > 0 Then
        lastError = "Партия содержит недопустимые символы переноса: " & Batch
        ValidateBatch = False
        Exit Function
    End If
    
    If InStr(suffixPart, """") > 0 Or InStr(suffixPart, "'") > 0 Then
        lastError = "Партия содержит недопустимые кавычки: " & Batch
        ValidateBatch = False
        Exit Function
    End If
    
    ValidateBatch = True
End Function

Public Function ValidateQuantity(quantity As Double) As Boolean
    ' Валидация количества
    If quantity <= 0 Then
        lastError = "Количество должно быть больше нуля: " & quantity
        ValidateQuantity = False
        Exit Function
    End If
    
    If quantity > 999999 Then
        lastError = "Слишком большое количество: " & quantity
        ValidateQuantity = False
        Exit Function
    End If
    
    ' Проверяем, что количество не содержит слишком много дробных знаков
    Dim quantityStr As String
    quantityStr = CStr(quantity)
    
    If InStr(quantityStr, ".") > 0 Or InStr(quantityStr, ",") > 0 Then
        Dim decimalPart As String
        If InStr(quantityStr, ".") > 0 Then
            decimalPart = Mid(quantityStr, InStr(quantityStr, ".") + 1)
        Else
            decimalPart = Mid(quantityStr, InStr(quantityStr, ",") + 1)
        End If
        
        If Len(decimalPart) > 3 Then
            lastError = "Слишком много знаков после запятой в количестве: " & quantity
            ValidateQuantity = False
            Exit Function
        End If
    End If
    
    ValidateQuantity = True
End Function

' ===== СПЕЦИАЛИЗИРОВАННЫЕ ФУНКЦИИ ВАЛИДАЦИИ =====

Public Function ValidateWarehouseExists(warehouse As String) As Boolean
    ' Проверяем, существует ли лист ангара в активной книге
    On Error Resume Next
    Dim ws As Worksheet
    Dim sheetName As String
    
    Select Case warehouse
        Case "5": sheetName = "Ангар 5"
        Case "6": sheetName = "Ангар 6"
        Case "7": sheetName = "Ангар 7"
        Case "8": sheetName = "Ангар 8"
        Case "9": sheetName = "Ангар 9"
        Case "10": sheetName = "Ангар 10"
        Case "11": sheetName = "Ангар 11"
        Case "12": sheetName = "Ангар 12"
        Case Else
            lastError = "Неизвестный ангар: " & warehouse
            ValidateWarehouseExists = False
            Exit Function
    End Select
    
    If ActiveWorkbook Is Nothing Then
        lastError = "Нет активной книги Excel"
        ValidateWarehouseExists = False
        On Error GoTo 0
        Exit Function
    End If
    
    Set ws = ActiveWorkbook.Worksheets(sheetName)
    If Err.Number <> 0 Then
        lastError = "Лист '" & sheetName & "' не найден в активной книге"
        ValidateWarehouseExists = False
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    
    ValidateWarehouseExists = True
End Function

Private Function IsValidShelfName(shelfName As String) As Boolean
    ' Проверяет корректность названия стеллажа
    shelfName = UCase(Trim(shelfName))
    
    ' Проверяем ПР-коды
    If Left(shelfName, 2) = "ПР" Then
        If Len(shelfName) = 3 Then
            ' ПРА, ПРБ, ПРВ и т.д.
            Dim lastChar As String
            lastChar = Right(shelfName, 1)
            If (lastChar >= "А" And lastChar <= "Я") Or (lastChar >= "A" And lastChar <= "Z") Then
                IsValidShelfName = True
                Exit Function
            End If
        End If
    End If
    
    ' Проверяем обычные буквы
    If Len(shelfName) = 1 Then
        If (shelfName >= "А" And shelfName <= "Я") Or (shelfName >= "A" And shelfName <= "Z") Then
            IsValidShelfName = True
            Exit Function
        End If
    End If
    
    IsValidShelfName = False
End Function

Private Function ValidateShelfInWarehouse(Shelf As String, warehouse As String) As Boolean
    ' Проверяет, что стеллаж существует в данном ангаре
    ' Используем функцию определения секции из ModuleConfig
    On Error Resume Next
    
    Dim section As String
    section = ModuleConfig.DetermineSectionByShelf(Shelf, warehouse)
    
    If Err.Number <> 0 Then
        lastError = "Ошибка определения секции для стеллажа " & Shelf & " в ангаре " & warehouse
        ValidateShelfInWarehouse = False
        On Error GoTo 0
        Exit Function
    End If
    
    If section = "" Then
        lastError = "Стеллаж " & Shelf & " не найден в ангаре " & warehouse
        ValidateShelfInWarehouse = False
        On Error GoTo 0
        Exit Function
    End If
    
    On Error GoTo 0
    ValidateShelfInWarehouse = True
End Function

Private Function GetMaxRowsForWarehouse(warehouse As String) As Integer
    ' Определяем максимальное количество рядов для каждого ангара
    ' Основано на анализе конфигураций
    Select Case warehouse
        Case "5": GetMaxRowsForWarehouse = 60
        Case "6": GetMaxRowsForWarehouse = 60
        Case "7": GetMaxRowsForWarehouse = 60
        Case "8": GetMaxRowsForWarehouse = 60
        Case "9": GetMaxRowsForWarehouse = 60
        Case "10": GetMaxRowsForWarehouse = 60
        Case "11": GetMaxRowsForWarehouse = 60
        Case "12": GetMaxRowsForWarehouse = 60
        Case Else: GetMaxRowsForWarehouse = 60 ' По умолчанию
    End Select
End Function

' ===== ФУНКЦИИ ВАЛИДАЦИИ СТРОК ВВОДА =====

Public Function ValidateInputLine(inputLine As String) As Boolean
    ' Валидация строки ввода в формате "6-5-Е-2 - Лерашанс - пар13 - 720"
    lastError = ""
    
    If Trim(inputLine) = "" Then
        lastError = "Пустая строка ввода"
        ValidateInputLine = False
        Exit Function
    End If
    
    ' Проверяем общий формат
    If InStr(inputLine, " - ") = 0 Then
        lastError = "Неверный формат строки. Ожидается разделитель ' - '"
        ValidateInputLine = False
        Exit Function
    End If
    
    Dim parts() As String
    parts = Split(inputLine, " - ")
    
    If UBound(parts) < 3 Then
        lastError = "Недостаточно частей в строке. Ожидается: Координаты - Товар - Партия - Количество"
        ValidateInputLine = False
        Exit Function
    End If
    
    ' Валидация координат
    If Not ValidateCoordinatesString(Trim(parts(0))) Then
        ValidateInputLine = False
        Exit Function
    End If
    
    ' Валидация товара
    If Not ValidateProductName(Trim(parts(1))) Then
        ValidateInputLine = False
        Exit Function
    End If
    
    ' Валидация партии
    If Not ValidateBatch(Trim(parts(2))) Then
        ValidateInputLine = False
        Exit Function
    End If
    
    ' Валидация количества
    If Not IsNumeric(Trim(parts(3))) Then
        lastError = "Количество должно быть числом: " & parts(3)
        ValidateInputLine = False
        Exit Function
    End If
    
    If Not ValidateQuantity(CDbl(Trim(parts(3)))) Then
        ValidateInputLine = False
        Exit Function
    End If
    
    ValidateInputLine = True
End Function

Private Function ValidateCoordinatesString(coordinates As String) As Boolean
    ' Валидация строки координат "6-5-Е-2"
    If InStr(coordinates, "-") = 0 Then
        lastError = "Неверный формат координат. Ожидается: Ангар-Ряд-Стеллаж-Уровень"
        ValidateCoordinatesString = False
        Exit Function
    End If
    
    Dim coordParts() As String
    coordParts = Split(coordinates, "-")
    
    If UBound(coordParts) < 3 Then
        lastError = "Недостаточно частей в координатах. Ожидается: Ангар-Ряд-Стеллаж-Уровень"
        ValidateCoordinatesString = False
        Exit Function
    End If
    
    ' Валидируем каждую часть
    If Not ValidateWarehouse(Trim(coordParts(0))) Then
        ValidateCoordinatesString = False
        Exit Function
    End If
    
    If Not ValidateRow(Trim(coordParts(1)), Trim(coordParts(0))) Then
        ValidateCoordinatesString = False
        Exit Function
    End If
    
    If Not ValidateShelf(Trim(coordParts(2)), Trim(coordParts(0))) Then
        ValidateCoordinatesString = False
        Exit Function
    End If
    
    If Not ValidateLevel(Trim(coordParts(3))) Then
        ValidateCoordinatesString = False
        Exit Function
    End If
    
    ValidateCoordinatesString = True
End Function

' ===== ФУНКЦИИ ПАКЕТНОЙ ВАЛИДАЦИИ =====

Public Function ValidateMultipleLines(inputText As String, ByRef errorReport As String) As Integer
    ' Валидирует несколько строк ввода, возвращает количество валидных строк
    Dim lines() As String
    lines = Split(inputText, vbNewLine)
    
    Dim validCount As Integer
    validCount = 0
    
    errorReport = ""
    
    Dim i As Integer
    For i = 0 To UBound(lines)
        Dim line As String
        line = Trim(lines(i))
        
        ' Пропускаем пустые строки и примеры
        If line <> "" And Not (InStr(LCase(line), "пример") > 0) Then
            If ValidateInputLine(line) Then
                validCount = validCount + 1
            Else
                errorReport = errorReport & "Строка " & (i + 1) & ": " & GetLastError() & vbNewLine
            End If
        End If
    Next i
    
    ValidateMultipleLines = validCount
End Function

' ===== СПРАВОЧНЫЕ ФУНКЦИИ =====

Public Function GetValidationSummary() As String
    ' Возвращает справку по правилам валидации
    Dim summary As String
    summary = "=== ПРАВИЛА ВАЛИДАЦИИ ДАННЫХ ===" & vbNewLine & vbNewLine
    summary = summary & "?? ФОРМАТ ВВОДА:" & vbNewLine
    summary = summary & "Ангар-Ряд-Стеллаж-Уровень - Товар - Партия - Количество" & vbNewLine
    summary = summary & "Пример: 6-5-Е-2 - Лерашанс - пар13 - 720" & vbNewLine & vbNewLine
    
    summary = summary & "?? ДОПУСТИМЫЕ ЗНАЧЕНИЯ:" & vbNewLine
    summary = summary & "• Ангары: " & VALID_WAREHOUSES & vbNewLine
    summary = summary & "• Ряды: 1-10 (зависит от ангара)" & vbNewLine
    summary = summary & "• Стеллажи: А-Я, ПРА, ПРБ, ПРВ... (зависит от ангара)" & vbNewLine
    summary = summary & "• Уровни: " & VALID_LEVELS & " (1=ниж.ряд, 2=2ряд, 3=3ряд)" & vbNewLine
    summary = summary & "• Товары: 2-50 символов, без спецсимволов" & vbNewLine
    summary = summary & "• Партии: должны начинаться с 'пар' + цифры" & vbNewLine
    summary = summary & "• Количество: больше 0, не более 999999, до 3 знаков после запятой" & vbNewLine & vbNewLine
    
    summary = summary & "?? ЧАСТЫЕ ОШИБКИ:" & vbNewLine
    summary = summary & "• Неправильный разделитель (используйте ' - ')" & vbNewLine
    summary = summary & "• Несуществующий ангар или ряд" & vbNewLine
    summary = summary & "• Стеллаж не существует в данном ангаре" & vbNewLine
    summary = summary & "• Партия не начинается с 'пар'" & vbNewLine
    summary = summary & "• Отрицательное или нулевое количество" & vbNewLine
    
    GetValidationSummary = summary
End Function

Public Function GetWarehouseInfo(warehouse As String) As String
    ' Возвращает информацию о конкретном ангаре
    If Not ValidateWarehouse(warehouse) Then
        GetWarehouseInfo = "Ошибка: " & GetLastError()
        Exit Function
    End If
    
    Dim info As String
    info = "=== ИНФОРМАЦИЯ ОБ АНГАРЕ " & warehouse & " ===" & vbNewLine
    info = info & "Максимальное количество рядов: " & GetMaxRowsForWarehouse(warehouse) & vbNewLine
    
    ' Получаем информацию о секциях
    On Error Resume Next
    Dim sectionInfo As String
    sectionInfo = ModuleConfig.GetWarehouseSectionsList(warehouse)
    If Err.Number = 0 Then
        info = info & sectionInfo
    Else
        info = info & "Информация о секциях недоступна"
    End If
    On Error GoTo 0
    
    GetWarehouseInfo = info
End Function

' ===== ТЕСТОВЫЕ И ОТЛАДОЧНЫЕ ФУНКЦИИ =====

Public Function TestModuleValidator() As Boolean
    ' Публичная функция для тестирования модуля
    On Error GoTo ErrorHandler
    
    ' Тестируем основные функции
    If Not ValidateWarehouse("7") Then
        TestModuleValidator = False
        Exit Function
    End If
    
    If Not ValidateLevel("1") Then
        TestModuleValidator = False
        Exit Function
    End If
    
    If Not ValidateInputLine("6-5-Е-2 - Лерашанс - пар13 - 720") Then
        TestModuleValidator = False
        Exit Function
    End If
    
    TestModuleValidator = True
    Exit Function
    
ErrorHandler:
    TestModuleValidator = False
End Function

Public Sub RunValidationTests()
    ' Запускает полный набор тестов валидации
    Dim testResults As String
    testResults = "=== РЕЗУЛЬТАТЫ ТЕСТИРОВАНИЯ ВАЛИДАЦИИ ===" & vbNewLine & vbNewLine
    
    ' Тест 1: Валидные данные
    If ValidateInputLine("6-5-Е-2 - Лерашанс - пар13 - 720") Then
        testResults = testResults & "? Тест валидных данных: ПРОЙДЕН" & vbNewLine
    Else
        testResults = testResults & "? Тест валидных данных: ПРОВАЛЕН - " & GetLastError() & vbNewLine
    End If
    
    ' Тест 2: Неверный ангар
    If Not ValidateInputLine("15-5-Е-2 - Лерашанс - пар13 - 720") Then
        testResults = testResults & "? Тест неверного ангара: ПРОЙДЕН" & vbNewLine
    Else
        testResults = testResults & "? Тест неверного ангара: ПРОВАЛЕН" & vbNewLine
    End If
    
    ' Тест 3: Неверная партия
    If Not ValidateInputLine("6-5-Е-2 - Лерашанс - abc13 - 720") Then
        testResults = testResults & "? Тест неверной партии: ПРОЙДЕН" & vbNewLine
    Else
        testResults = testResults & "? Тест неверной партии: ПРОВАЛЕН" & vbNewLine
    End If
    
    ' Тест 4: Нулевое количество
    If Not ValidateInputLine("6-5-Е-2 - Лерашанс - пар13 - 0") Then
        testResults = testResults & "? Тест нулевого количества: ПРОЙДЕН" & vbNewLine
    Else
        testResults = testResults & "? Тест нулевого количества: ПРОВАЛЕН" & vbNewLine
    End If
    
    ' Тест 5: Модуль в целом
    If TestModuleValidator() Then
        testResults = testResults & "? Общий тест модуля: ПРОЙДЕН" & vbNewLine
    Else
        testResults = testResults & "? Общий тест модуля: ПРОВАЛЕН" & vbNewLine
    End If
    
    testResults = testResults & vbNewLine & "Тестирование завершено!"
    
    MsgBox testResults, vbInformation, "Результаты тестирования валидации"
End Sub
