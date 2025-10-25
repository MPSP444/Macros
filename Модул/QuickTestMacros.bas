Attribute VB_Name = "QuickTestMacros"


' ===== QuickTestMacros - Быстрые тестовые макросы для проверки системы =====
Option Explicit

' Быстрый запуск основной системы
Sub RunMassInput()
    Call ModuleLauncher.ЗапуститьСистему
End Sub

' Быстрая тестовая очистка одного ангара
Sub TestClearOne()
    Call TestClearSingleWarehouse
End Sub

' Быстрая полная очистка всех ангаров
Sub ClearAllWarehouses()
    Call ModuleLauncher.ОчиститьВсе
End Sub

' Быстрый тест системы
Sub TestSystemQuick()
    Call MainLauncher.TestSystem
End Sub
Sub TestCorrectLogic()
    Call TestOptimizeWithCorrectLogic
End Sub
' Тест только валидации (без GUI)
Sub TestValidationOnly()
    On Error GoTo ErrorHandler
    
    ' Инициализируем логирование
    Call ModuleLogger.InitializeLogger(True, ModuleLogger.LOG_LEVEL_INFO)
    
    ModuleLogger.LogMessage "=== ТЕСТ ВАЛИДАЦИИ ==="
    
    ' Тестируем различные форматы данных
    Dim testData(4) As String
    testData(0) = "6-5-Е-2 - Лерашанс - пар13 - 720"  ' Корректный
    testData(1) = "6-15-Е-2 - Лерашанс - пар13 - 720" ' Неверный ряд
    testData(2) = "15-5-Е-2 - Лерашанс - пар13 - 720" ' Неверный ангар
    testData(3) = "6-5-Я-2 - Лерашанс - пар13 - 720"  ' Может быть неверный стеллаж
    testData(4) = "6-5-Е-5 - Лерашанс - пар13 - 720"  ' Неверный уровень
    
    Dim i As Integer
    For i = 0 To 4
        ModuleLogger.LogMessage "Тест " & (i + 1) & ": " & testData(i)
        
        ' Тестируем валидацию строки
        If ModuleValidator.ValidateInputLine(testData(i)) Then
            ModuleLogger.LogSuccess "? Строка валидна"
        Else
            ModuleLogger.LogError "? Ошибка валидации: " & ModuleValidator.GetLastError()
        End If
    Next i
    
    ModuleLogger.LogMessage "=== КОНЕЦ ТЕСТА ВАЛИДАЦИИ ==="
    
    MsgBox "Тест валидации завершен! Проверьте Immediate Window (Ctrl+G) для деталей.", vbInformation, "Тест валидации"
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка теста валидации: " & Err.description, vbCritical, "Ошибка"
End Sub

' Показать справку по быстрым макросам
Sub ShowQuickHelp()
    Dim helpText As String
    helpText = "=== БЫСТРЫЕ ТЕСТОВЫЕ МАКРОСЫ ===" & vbNewLine & vbNewLine
    
    helpText = helpText & "?? ОСНОВНЫЕ ФУНКЦИИ:" & vbNewLine
    helpText = helpText & "• RunMassInput() - Запуск системы массового ввода" & vbNewLine
    helpText = helpText & "• TestClearOne() - Тестовая очистка одного ангара" & vbNewLine
    helpText = helpText & "• ClearAllWarehouses() - Полная очистка всех ангаров" & vbNewLine
    helpText = helpText & "• TestSystemQuick() - Быстрый тест системы" & vbNewLine
    helpText = helpText & "• TestValidationOnly() - Тест только валидации" & vbNewLine & vbNewLine
    
    helpText = helpText & "?? ОПТИМИЗАЦИЯ ЗАГОЛОВКОВ:" & vbNewLine
    helpText = helpText & "• OptimizeAllHeaders() - Объединить заголовки во всех ангарах" & vbNewLine
    helpText = helpText & "• OptimizeOneWarehouse() - Объединить заголовки в одном ангаре" & vbNewLine
    helpText = helpText & "• PreviewOptimization() - Превью объединения без изменений" & vbNewLine & vbNewLine
    
    helpText = helpText & "?? КАК ИСПОЛЬЗОВАТЬ:" & vbNewLine
    helpText = helpText & "1. Нажмите Alt+F8 для списка макросов" & vbNewLine
    helpText = helpText & "2. Выберите нужный макрос и нажмите 'Выполнить'" & vbNewLine
    helpText = helpText & "3. Или создайте кнопки на панели инструментов" & vbNewLine & vbNewLine
    
    helpText = helpText & "?? РЕКОМЕНДУЕМЫЙ ПОРЯДОК:" & vbNewLine
    helpText = helpText & "1. RunMassInput() - введите товары" & vbNewLine
    helpText = helpText & "2. OptimizeAllHeaders() - объедините заголовки" & vbNewLine
    helpText = helpText & "3. При необходимости - ClearAllWarehouses()" & vbNewLine
    
    MsgBox helpText, vbInformation, "Справка по быстрым макросам"
End Sub

' Создать тестовые данные для демонстрации
Sub CreateSampleData()
    Dim sampleData As String
    sampleData = "=== ОБРАЗЕЦ ДАННЫХ ДЛЯ ТЕСТИРОВАНИЯ ===" & vbNewLine & vbNewLine
    
    sampleData = sampleData & "Скопируйте эти строки в систему массового ввода:" & vbNewLine & vbNewLine
    
    sampleData = sampleData & "6-5-Е-2 - Лерашанс - пар13 - 720" & vbNewLine
    sampleData = sampleData & "6-5-П-2 - Дикошанс - пар25 - 50" & vbNewLine
    sampleData = sampleData & "6-6-А-1 - Лерашанс - пар13 - 720" & vbNewLine
    sampleData = sampleData & "6-6-А-2 - Лерашанс - пар13 - 720" & vbNewLine
    sampleData = sampleData & "6-6-Б-1 - Лерашанс - пар13 - 720" & vbNewLine
    sampleData = sampleData & "7-1-А-1 - Босфор - пар03 - 720" & vbNewLine
    sampleData = sampleData & "7-1-А-1 - Босфор - пар12 - 360" & vbNewLine
    sampleData = sampleData & "8-2-В-2 - Чистосад - пар07 - 480" & vbNewLine & vbNewLine
    
    sampleData = sampleData & "?? ЧТО ПОКАЖЕТ СИСТЕМА:" & vbNewLine
    sampleData = sampleData & "• Ряд 5 в Ангаре 6: разные секции > отдельные заголовки" & vbNewLine
    sampleData = sampleData & "• Ряд 6 в Ангаре 6: 'Лерашанс' (обычный заголовок)" & vbNewLine
    sampleData = sampleData & "• Ряд 1 в Ангаре 7: 'Босфор' (обычный заголовок)" & vbNewLine
    sampleData = sampleData & "• В 7-1-А-1: количество 1080, партия 'пар03/пар12'" & vbNewLine
    
    MsgBox sampleData, vbInformation, "Образец тестовых данных"
End Sub

' Тестовая очистка одного ангара (для проверки)
Sub TestClearSingleWarehouse()
    On Error GoTo ErrorHandler
    
    ' Инициализируем логирование
    Call ModuleLogger.InitializeLogger(True, ModuleLogger.LOG_LEVEL_INFO)
    
    ' Запрашиваем номер ангара для тестирования
    Dim warehouseNum As String
    warehouseNum = InputBox("Введите номер ангара для тестовой очистки (5-12):", "Тестовая очистка", "7")
    
    If warehouseNum = "" Then Exit Sub
    
    ' Проверяем корректность номера
    If Not IsNumeric(warehouseNum) Or CInt(warehouseNum) < 5 Or CInt(warehouseNum) > 12 Then
        MsgBox "Неверный номер ангара! Используйте числа от 5 до 12.", vbExclamation, "Ошибка"
        Exit Sub
    End If
    
    ' Подтверждение очистки
    Dim confirmMsg As String
    confirmMsg = "?? ТЕСТОВАЯ ОЧИСТКА АНГАРА " & warehouseNum & vbNewLine & vbNewLine & _
                "Будут удалены:" & vbNewLine & _
                "• Заголовки товаров в этом ангаре" & vbNewLine & _
                "• Все количества товаров" & vbNewLine & _
                "• Все партии" & vbNewLine & vbNewLine & _
                "Сохранятся:" & vbNewLine & _
                "• Названия стеллажей (А, Б, В...)" & vbNewLine & _
                "• Номера рядов (1, 2, 3...)" & vbNewLine & _
                "• Структура таблицы" & vbNewLine & vbNewLine & _
                "?? Продолжить тестовую очистку?"
    
    If MsgBox(confirmMsg, vbYesNo + vbQuestion, "Тестовая очистка") = vbNo Then
        Exit Sub
    End If
    
    ' Выполняем очистку
    ModuleLogger.LogMessage "=== ТЕСТОВАЯ ОЧИСТКА АНГАРА " & warehouseNum & " ==="
    
    Dim result As String
    result = ModuleDataCleaner.ClearSpecificWarehouse(warehouseNum)
    
    ' Показываем результат
    If result = "" Then
        MsgBox "? Тестовая очистка Ангара " & warehouseNum & " завершена успешно!" & vbNewLine & _
               "Проверьте результат и при необходимости запустите полную очистку всех ангаров.", _
               vbInformation, "Тестовая очистка завершена"
        ModuleLogger.LogSuccess "Тестовая очистка Ангара " & warehouseNum & " завершена успешно"
    Else
        MsgBox "?? Тестовая очистка завершена с ошибками:" & vbNewLine & result, vbExclamation, "Ошибки очистки"
        ModuleLogger.LogError "Ошибки при тестовой очистке: " & result
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка тестовой очистки: " & Err.description, vbCritical, "Ошибка"
    ModuleLogger.LogError "Ошибка тестовой очистки: " & Err.description
End Sub

' Демонстрация работы с секциями
Sub TestSectionDetection()
    On Error GoTo ErrorHandler
    
    ' Инициализируем логирование
    Call ModuleLogger.InitializeLogger(True, ModuleLogger.LOG_LEVEL_DEBUG)
    
    ModuleLogger.LogMessage "=== ТЕСТ ОПРЕДЕЛЕНИЯ СЕКЦИЙ ==="
    
    ' Тестовые данные для проверки секций
    Dim testItems(7, 2) As String  ' warehouse, shelf, expected_section
    testItems(0, 0) = "6": testItems(0, 1) = "Е": testItems(0, 2) = "UPPER"
    testItems(1, 0) = "6": testItems(1, 1) = "П": testItems(1, 2) = "LOWER"
    testItems(2, 0) = "7": testItems(2, 1) = "А": testItems(2, 2) = "UPPER"
    testItems(3, 0) = "7": testItems(3, 1) = "Ж": testItems(3, 2) = "LOWER"
    testItems(4, 0) = "9": testItems(4, 1) = "И": testItems(4, 2) = "UPPER"
    testItems(5, 0) = "9": testItems(5, 1) = "М": testItems(5, 2) = "LOWER"
    testItems(6, 0) = "12": testItems(6, 1) = "В": testItems(6, 2) = "UPPER"
    testItems(7, 0) = "12": testItems(7, 1) = "З": testItems(7, 2) = "LOWER"
    
    Dim i As Integer
    For i = 0 To 7
        Dim warehouse As String, Shelf As String, expectedSection As String
        warehouse = testItems(i, 0)
        Shelf = testItems(i, 1)
        expectedSection = testItems(i, 2)
        
        Dim actualSection As String
        actualSection = ModuleConfig.DetermineSectionByShelf(Shelf, warehouse)
        
        If actualSection = expectedSection Then
            ModuleLogger.LogSuccess "? Ангар " & warehouse & ", стеллаж " & Shelf & " > " & actualSection & " (ожидалось: " & expectedSection & ")"
        Else
            ModuleLogger.LogError "? Ангар " & warehouse & ", стеллаж " & Shelf & " > " & actualSection & " (ожидалось: " & expectedSection & ")"
        End If
    Next i
    
    ModuleLogger.LogMessage "=== КОНЕЦ ТЕСТА ОПРЕДЕЛЕНИЯ СЕКЦИЙ ==="
    
    MsgBox "Тест определения секций завершен! Проверьте Immediate Window (Ctrl+G) для деталей.", vbInformation, "Тест секций"
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка теста секций: " & Err.description, vbCritical, "Ошибка"
End Sub

' Оптимизация заголовков
Sub OptimizeAllHeaders()
    Call ModuleHeaderOptimizer.OptimizeAllWarehouseHeaders
End Sub

' Оптимизация одного ангара
Sub OptimizeOneWarehouse()
    Dim warehouse As String
    warehouse = InputBox("Введите номер ангара для оптимизации (5-12):", "Оптимизация ангара", "6")
    
    If warehouse <> "" And IsNumeric(warehouse) Then
        If CInt(warehouse) >= 5 And CInt(warehouse) <= 12 Then
            Call ModuleHeaderOptimizer.OptimizeSpecificWarehouse(warehouse)
        Else
            MsgBox "Неверный номер ангара! Используйте числа от 5 до 12.", vbExclamation, "Ошибка"
        End If
    End If
End Sub
' ВРЕМЕННЫЙ КОД ДЛЯ ОТЛАДКИ - добавьте в любой модуль для теста

Sub TestTypes()
    ' Проверяем, что все типы определены правильно
    
    ' Тест 1: WarehouseConfig
    Dim config As ModuleTypes.WarehouseConfig
    config.sheetName = "Тест"
    Debug.Print "WarehouseConfig: OK"
    
    ' Тест 2: WarehouseOperation
    Dim operation As ModuleTypes.WarehouseOperation
    operation.warehouse = "6"
    Debug.Print "WarehouseOperation: OK"
    
    ' Тест 3: WarehouseLocation
    Dim location As ModuleTypes.WarehouseLocation
    location.warehouse = "7"
    Debug.Print "WarehouseLocation: OK"
    
    MsgBox "Все типы данных работают корректно!", vbInformation, "Тест типов"
End Sub

' ДОБАВЬТЕ ЭТОТ КОД В ЛЮБОЙ МОДУЛЬ ДЛЯ ПОИСКА ОШИБКИ

Sub FindErrorLocation()
    ' Этот код поможет найти, где именно ошибка
    
    Dim i As Integer
    For i = 1 To 20
        Debug.Print "Проверяем модуль " & i
        
        ' Если здесь появится ошибка - значит проблема в этой строке
        On Error GoTo ErrorFound
        
        ' Попробуем создать переменную правильного типа
        Dim testVar As ModuleTypes.WarehouseOperation
        testVar.warehouse = "6"
        
        On Error GoTo 0
    Next i
    
    MsgBox "Ошибка не найдена в тесте", vbInformation
    Exit Sub
    
ErrorFound:
    MsgBox "Ошибка найдена в строке " & Erl & ": " & Err.description, vbCritical
End Sub
' Превью оптимизации
Sub PreviewOptimization()
    Dim warehouse As String
    warehouse = InputBox("Введите номер ангара для превью (5-12):", "Превью оптимизации", "6")
    
    If warehouse <> "" And IsNumeric(warehouse) Then
        If CInt(warehouse) >= 5 And CInt(warehouse) <= 12 Then
            Call ModuleHeaderOptimizer.ShowOptimizationPreview(warehouse)
        Else
            MsgBox "Неверный номер ангара! Используйте числа от 5 до 12.", vbExclamation, "Ошибка"
        End If
    End If
End Sub
    On Error Resume Next
    
    Dim info As String
    info = "=== ИНФОРМАЦИЯ О СИСТЕМЕ МАССОВОГО ВВОДА ===" & vbNewLine & vbNewLine
    
    ' Основная информация
    info = info & "?? Версия системы: 2.0 (с поддержкой секций)" & vbNewLine
    info = info & "?? Дата: " & Format(Date, "dd.mm.yyyy") & vbNewLine
    info = info & "? Время: " & Format(Time, "hh:mm:ss") & vbNewLine
    
    ' Информация о Excel
    info = info & "??? Версия Excel: " & Application.version & vbNewLine
    
    ' Информация о книге
    If ActiveWorkbook Is Nothing Then
        info = info & "?? Активная книга: НЕ НАЙДЕНА" & vbNewLine
    Else
        info = info & "?? Активная книга: " & ActiveWorkbook.Name & vbNewLine
        info = info & "?? Путь: " & Left(ActiveWorkbook.FullName, 50) & "..." & vbNewLine
        
        ' Подсчет ангаров
        Dim warehouseCount As Integer
        warehouseCount = 0
        
        Dim i As Integer
        For i = 5 To 12
            Dim ws As Worksheet
            Set ws = ActiveWorkbook.Worksheets("Ангар " & i)
            If Err.Number = 0 Then
                warehouseCount = warehouseCount + 1
            End If
        Next i
        
        info = info & "?? Доступных ангаров: " & warehouseCount & " из 8" & vbNewLine
    End If
    
    ' Статус модулей
    info = info & vbNewLine & "?? СТАТУС МОДУЛЕЙ:" & vbNewLine
    
    ' Проверяем основные модули
    info = info & "• ModuleConfig: "
    If Not ModuleConfig.GetWarehouseConfig("6").sheetName = "" Then
        info = info & "? OK" & vbNewLine
    Else
        info = info & "? Ошибка" & vbNewLine
    End If
    
    info = info & "• ModuleValidator: "
    If ModuleValidator.TestModuleValidator() Then
        info = info & "? OK" & vbNewLine
    Else
        info = info & "? Ошибка" & vbNewLine
    End If
    
    info = info & "• ModuleLogger: "
    If ModuleLogger.IsLoggerInitialized() Then
        info = info & "? Инициализирован" & vbNewLine
    Else
        info = info & "?? Не инициализирован" & vbNewLine
    End If
    
    MsgBox info, vbInformation, "Информация о системе"
End Sub

