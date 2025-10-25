Attribute VB_Name = "MainLauncher"


' ===== MainLauncher - Главный модуль запуска системы массового ввода =====
Option Explicit

Public Sub StartMassInputSystem()
    On Error GoTo ErrorHandler
    
    ' Инициализируем систему логирования
    Call ModuleLogger.InitializeLogger(True, ModuleLogger.LOG_LEVEL_INFO)
    
    ' Логируем информацию о системе
    Call ModuleLogger.LogSystemInfo
    
    ' Проверяем готовность системы
    If Not CheckSystemReadiness() Then
        Exit Sub
    End If
    
    ' Показываем форму массового ввода
    ModuleLogger.LogMessage "Запуск формы массового ввода товаров"
    MassProductInputForm.Show
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка запуска системы: " & Err.description, vbCritical, "Ошибка системы"
    ModuleLogger.LogError "Критическая ошибка запуска: " & Err.description
End Sub

Public Sub LaunchMassInput()
    ' Альтернативная функция запуска (для совместимости)
    Call StartMassInputSystem
End Sub

Public Sub QuickClearAllData()
    On Error GoTo ErrorHandler
    
    ' Инициализируем логирование
    Call ModuleLogger.InitializeLogger(True, ModuleLogger.LOG_LEVEL_INFO)
    
    ' Проверяем готовность
    If Not CheckSystemReadiness() Then
        Exit Sub
    End If
    
    ' Подтверждение очистки
    Dim confirmMsg As String
    confirmMsg = "?????? БЫСТРАЯ ОЧИСТКА ВСЕХ АНГАРОВ" & vbNewLine & vbNewLine & _
                "Будут удалены:" & vbNewLine & _
                "• Все заголовки товаров" & vbNewLine & _
                "• Все количества" & vbNewLine & _
                "• Все партии" & vbNewLine & _
                "• Все результаты подсчетов" & vbNewLine & vbNewLine & _
                "?? Продолжить очистку?"
    
    If MsgBox(confirmMsg, vbYesNo + vbQuestion, "Быстрая очистка") = vbNo Then
        Exit Sub
    End If
    
    ' Дополнительное подтверждение
    If MsgBox("Вы ДЕЙСТВИТЕЛЬНО уверены?", vbYesNo + vbCritical, "Последнее предупреждение") = vbNo Then
        Exit Sub
    End If
    
    ' Выполняем очистку
    ModuleLogger.LogMessage "=== ЗАПУСК БЫСТРОЙ ОЧИСТКИ ==="
    Dim result As String
    result = ModuleDataCleaner.ClearAllWarehouses()
    
    ' Показываем результат
    If result = "" Then
        MsgBox "? Быстрая очистка завершена успешно!" & vbNewLine & _
               "Все ангары готовы для новых данных.", vbInformation, "Очистка завершена"
        ModuleLogger.LogSuccess "Быстрая очистка завершена успешно"
    Else
        MsgBox "?? Очистка завершена с ошибками:" & vbNewLine & result, vbExclamation, "Ошибки очистки"
        ModuleLogger.LogError "Ошибки при быстрой очистке: " & result
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка быстрой очистки: " & Err.description, vbCritical, "Ошибка"
    ModuleLogger.LogError "Ошибка быстрой очистки: " & Err.description
End Sub

Public Sub ShowSystemHelp()
    ' Функция для отображения справки по системе
    Dim helpText As String
    helpText = "=== СИСТЕМА МАССОВОГО ВВОДА ТОВАРОВ НА СКЛАД ===" & vbNewLine & vbNewLine
    
    helpText = helpText & "?? ОСНОВНЫЕ ФУНКЦИИ:" & vbNewLine
    helpText = helpText & "• StartMassInputSystem() - Запуск основной системы" & vbNewLine
    helpText = helpText & "• QuickClearAllData() - Быстрая очистка всех ангаров" & vbNewLine
    helpText = helpText & "• ShowSystemHelp() - Эта справка" & vbNewLine & vbNewLine
    
    helpText = helpText & "?? ФОРМАТ ВВОДА ДАННЫХ:" & vbNewLine
    helpText = helpText & "Ангар-Ряд-Стеллаж-Уровень - Товар - Партия - Количество" & vbNewLine
    helpText = helpText & "Пример: 6-5-Е-2 - Лерашанс - пар13 - 720" & vbNewLine & vbNewLine
    
    helpText = helpText & "? ВОЗМОЖНОСТИ СИСТЕМЫ:" & vbNewLine
    helpText = helpText & "• Автоматическое создание заголовков товаров" & vbNewLine
    helpText = helpText & "• Правильная работа с секциями ангаров" & vbNewLine
    helpText = helpText & "• Выделение конфликтов (разные товары в одном ряду)" & vbNewLine
    helpText = helpText & "• Суммирование дублей (одинаковые позиции)" & vbNewLine
    helpText = helpText & "• Подробное логирование операций" & vbNewLine
    helpText = helpText & "• Валидация всех входных данных" & vbNewLine & vbNewLine
    
    helpText = helpText & "?? ПОДДЕРЖИВАЕМЫЕ АНГАРЫ: 5, 6, 7, 8, 9, 10, 11, 12" & vbNewLine
    helpText = helpText & "?? УРОВНИ СТЕЛЛАЖЕЙ: 1 (ниж.ряд), 2 (2ряд), 3 (3ряд)" & vbNewLine
    
    MsgBox helpText, vbInformation, "Справка по системе"
End Sub

' ===== НЕДОСТАЮЩИЕ ФУНКЦИИ (ВОТ ОНИ!) =====

Private Function CheckSystemReadiness() As Boolean
    ' Проверяет готовность системы к работе
    CheckSystemReadiness = False
    
    ' Проверяем наличие активной книги
    If ActiveWorkbook Is Nothing Then
        MsgBox "? Откройте книгу Excel с листами ангаров перед запуском системы!", _
               vbExclamation, "Нет активной книги"
        ModuleLogger.LogError "Нет активной книги Excel"
        Exit Function
    End If
    
    ' Проверяем наличие ключевых листов ангаров
    Dim missingSheets As String
    missingSheets = ""
    
    Dim i As Integer
    For i = 5 To 12
        On Error Resume Next
        Dim ws As Worksheet
        Set ws = ActiveWorkbook.Worksheets("Ангар " & i)
        If Err.Number <> 0 Then
            missingSheets = missingSheets & "Ангар " & i & ", "
        End If
        On Error GoTo 0
    Next i
    
    If missingSheets <> "" Then
        missingSheets = Left(missingSheets, Len(missingSheets) - 2) ' Убираем последнюю запятую
        Dim continueResult As VbMsgBoxResult
        continueResult = MsgBox("?? В активной книге не найдены следующие листы:" & vbNewLine & _
               missingSheets & vbNewLine & vbNewLine & _
               "Система может работать некорректно. Продолжить?", _
               vbYesNo + vbQuestion, "Отсутствуют листы ангаров")
        
        If continueResult = vbNo Then
            Exit Function
        End If
        
        ModuleLogger.LogWarning "Отсутствуют листы: " & missingSheets
    End If
    
    ' Проверяем доступность модулей
    If Not CheckModulesAvailability() Then
        Exit Function
    End If
    
    ModuleLogger.LogSuccess "Система готова к работе"
    CheckSystemReadiness = True
End Function

Private Function CheckModulesAvailability() As Boolean
    ' Проверяем доступность ключевых модулей
    On Error GoTo ErrorHandler
    
    ' Проверяем ModuleConfig
    Dim testConfig As ModuleTypes.WarehouseConfig
    testConfig = ModuleConfig.GetWarehouseConfig("7")
    If testConfig.sheetName = "" Then
        MsgBox "? Модуль ModuleConfig недоступен или поврежден!", vbCritical, "Ошибка модуля"
        Exit Function
    End If
    
    ' Проверяем ModuleValidator
    If Not ModuleValidator.TestModuleValidator() Then
        MsgBox "? Модуль ModuleValidator недоступен или поврежден!", vbCritical, "Ошибка модуля"
        Exit Function
    End If
    
    CheckModulesAvailability = True
    Exit Function
    
ErrorHandler:
    MsgBox "? Ошибка проверки модулей: " & Err.description, vbCritical, "Ошибка системы"
    ModuleLogger.LogError "Ошибка проверки модулей: " & Err.description
    CheckModulesAvailability = False
End Function

Public Sub TestSystem()
    ' Простой тест системы
    On Error GoTo ErrorHandler
    
    MsgBox "?? Запуск тестирования системы...", vbInformation, "Тест системы"
    
    ' Инициализируем логирование для теста
    Call ModuleLogger.InitializeLogger(True, ModuleLogger.LOG_LEVEL_DEBUG)
    
    ' Тестируем основные компоненты
    Dim testResults As String
    testResults = "=== РЕЗУЛЬТАТЫ ТЕСТИРОВАНИЯ ===" & vbNewLine & vbNewLine
    
    ' Тест 1: ModuleValidator
    If ModuleValidator.TestModuleValidator() Then
        testResults = testResults & "? ModuleValidator: OK" & vbNewLine
    Else
        testResults = testResults & "? ModuleValidator: ОШИБКА" & vbNewLine
    End If
    
    ' Тест 2: ModuleConfig
    On Error Resume Next
    Dim testConfig As ModuleTypes.WarehouseConfig
    testConfig = ModuleConfig.GetWarehouseConfig("7")
    If Err.Number = 0 And testConfig.sheetName <> "" Then
        testResults = testResults & "? ModuleConfig: OK" & vbNewLine
    Else
        testResults = testResults & "? ModuleConfig: ОШИБКА" & vbNewLine
    End If
    On Error GoTo ErrorHandler
    
    ' Тест 3: Активная книга
    If ActiveWorkbook Is Nothing Then
        testResults = testResults & "? Активная книга: НЕ НАЙДЕНА" & vbNewLine
    Else
        testResults = testResults & "? Активная книга: " & ActiveWorkbook.Name & vbNewLine
        
        ' Тест 4: Листы ангаров
        Dim foundSheets As Integer
        foundSheets = 0
        Dim i As Integer
        For i = 5 To 12
            On Error Resume Next
            Dim ws As Worksheet
            Set ws = ActiveWorkbook.Worksheets("Ангар " & i)
            If Err.Number = 0 Then
                foundSheets = foundSheets + 1
            End If
            On Error GoTo ErrorHandler
        Next i
        
        testResults = testResults & "?? Найдено листов ангаров: " & foundSheets & " из 8" & vbNewLine
    End If
    
    testResults = testResults & vbNewLine & "Система готова к использованию!"
    
    MsgBox testResults, vbInformation, "Результаты тестирования"
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка тестирования: " & Err.description, vbCritical, "Ошибка теста"
End Sub

Public Sub ExportCurrentLog()
    ' Функция для экспорта лога в файл (удобная обертка)
    On Error GoTo ErrorHandler
    
    If ModuleLogger.GetLogBuffer() = "" Then
        MsgBox "Лог пуст. Сначала выполните какие-либо операции.", vbInformation, "Пустой лог"
        Exit Sub
    End If
    
    Call ModuleLogger.SaveLogToFile
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка экспорта лога: " & Err.description, vbExclamation, "Ошибка"
End Sub

Public Sub ShowSystemStatistics()
    ' Функция для отображения статистики системы
    Dim stats As String
    stats = "=== СТАТИСТИКА СИСТЕМЫ ===" & vbNewLine & vbNewLine
    
    ' Подсчитываем доступные ангары
    Dim availableWarehouses As Integer
    availableWarehouses = 0
    Dim warehouseList As String
    warehouseList = ""
    
    Dim i As Integer
    For i = 5 To 12
        On Error Resume Next
        Dim ws As Worksheet
        Set ws = ActiveWorkbook.Worksheets("Ангар " & i)
        If Err.Number = 0 Then
            availableWarehouses = availableWarehouses + 1
            warehouseList = warehouseList & i & ", "
        End If
        On Error GoTo 0
    Next i
    
    If warehouseList <> "" Then
        warehouseList = Left(warehouseList, Len(warehouseList) - 2)
    End If
    
    If ActiveWorkbook Is Nothing Then
        stats = stats & "? Активная книга: НЕ НАЙДЕНА" & vbNewLine
    Else
        stats = stats & "?? Активная книга: " & ActiveWorkbook.Name & vbNewLine
    End If
    
    stats = stats & "?? Доступные ангары: " & availableWarehouses & " из 8" & vbNewLine
    stats = stats & "?? Список ангаров: " & warehouseList & vbNewLine
    stats = stats & "??? Версия Excel: " & Application.version & vbNewLine
    stats = stats & "?? Время: " & Format(Now, "dd.mm.yyyy hh:mm:ss") & vbNewLine
    
    MsgBox stats, vbInformation, "Статистика системы"
End Sub

