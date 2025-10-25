Attribute VB_Name = "ModuleLogger"


' ===== ModuleLogger - Система логирования для массового ввода =====
Option Explicit

' ===== КОНСТАНТЫ УРОВНЕЙ ЛОГИРОВАНИЯ =====
Public Const LOG_LEVEL_DEBUG As Integer = 0
Public Const LOG_LEVEL_INFO As Integer = 1
Public Const LOG_LEVEL_WARNING As Integer = 2
Public Const LOG_LEVEL_ERROR As Integer = 3
Public Const LOG_LEVEL_SUCCESS As Integer = 4

' ===== ГЛОБАЛЬНЫЕ ПЕРЕМЕННЫЕ =====
Private loggerInitialized As Boolean  ' ИСПРАВЛЕНО: переименовано из isInitialized
Private currentLogLevel As Integer
Private immediateWindowEnabled As Boolean  ' ИСПРАВЛЕНО: переименовано из enableImmediateWindow
Private logBuffer As String
Private maxBufferSize As Long

' ===== ИНИЦИАЛИЗАЦИЯ И НАСТРОЙКА =====

Public Sub InitializeLogger(Optional enableImmediate As Boolean = True, Optional logLevel As Integer = LOG_LEVEL_INFO)
    ' Инициализирует систему логирования
    On Error Resume Next
    
    loggerInitialized = True  ' ИСПРАВЛЕНО
    immediateWindowEnabled = enableImmediate  ' ИСПРАВЛЕНО
    currentLogLevel = logLevel
    logBuffer = ""
    maxBufferSize = 50000 ' Максимальный размер буфера лога
    
    ' Логируем успешную инициализацию
    LogInfo "=== СИСТЕМА ЛОГИРОВАНИЯ ИНИЦИАЛИЗИРОВАНА ==="
    LogInfo "Уровень логирования: " & GetLogLevelName(logLevel)
    LogInfo "Immediate Window: " & IIf(enableImmediate, "включен", "отключен")
    LogInfo "Время запуска: " & Format(Now, "dd.mm.yyyy hh:mm:ss")
End Sub

Public Sub SetLogLevel(logLevel As Integer)
    ' Изменяет уровень логирования
    If Not loggerInitialized Then Call InitializeLogger  ' ИСПРАВЛЕНО
    
    currentLogLevel = logLevel
    LogInfo "Уровень логирования изменен на: " & GetLogLevelName(logLevel)
End Sub

Public Sub EnableImmediateWindow(enable As Boolean)
    ' Включает/отключает вывод в Immediate Window
    immediateWindowEnabled = enable  ' ИСПРАВЛЕНО
    LogInfo "Immediate Window: " & IIf(enable, "включен", "отключен")
End Sub

' ===== ОСНОВНЫЕ ФУНКЦИИ ЛОГИРОВАНИЯ =====

Public Sub LogMessage(message As String)
    ' Общая функция логирования (уровень INFO)
    LogInfo message
End Sub

Public Sub LogDebug(message As String)
    ' Отладочные сообщения
    WriteLog "DEBUG", message, LOG_LEVEL_DEBUG
End Sub

Public Sub LogInfo(message As String)
    ' Информационные сообщения
    WriteLog "INFO", message, LOG_LEVEL_INFO
End Sub

Public Sub LogWarning(message As String)
    ' Предупреждения
    WriteLog "WARNING", message, LOG_LEVEL_WARNING
End Sub

Public Sub LogError(message As String)
    ' Ошибки
    WriteLog "ERROR", message, LOG_LEVEL_ERROR
End Sub

Public Sub LogSuccess(message As String)
    ' Успешные операции
    WriteLog "SUCCESS", message, LOG_LEVEL_SUCCESS
End Sub

' ===== СИСТЕМНЫЕ ФУНКЦИИ ЛОГИРОВАНИЯ =====

Public Sub LogSystemInfo()
    ' Логирует информацию о системе
    If Not loggerInitialized Then Call InitializeLogger  ' ИСПРАВЛЕНО
    
    LogInfo "=== ИНФОРМАЦИЯ О СИСТЕМЕ ==="
    LogInfo "Дата: " & Format(Date, "dd.mm.yyyy")
    LogInfo "Время: " & Format(Time, "hh:mm:ss")
    LogInfo "Версия Excel: " & Application.version
    
    If ActiveWorkbook Is Nothing Then
        LogWarning "Активная книга: НЕ НАЙДЕНА"
    Else
        LogInfo "Активная книга: " & ActiveWorkbook.Name
        LogInfo "Путь к книге: " & ActiveWorkbook.Path
    End If
    
    LogInfo "Количество листов: " & IIf(ActiveWorkbook Is Nothing, "неизвестно", ActiveWorkbook.Worksheets.count)
    LogInfo "=== КОНЕЦ СИСТЕМНОЙ ИНФОРМАЦИИ ==="
End Sub

Public Sub LogOperationStart(operationName As String)
    ' Логирует начало операции
    LogInfo ">>> НАЧАЛО ОПЕРАЦИИ: " & operationName & " <<<"
End Sub

Public Sub LogOperationEnd(operationName As String, Optional isSuccess As Boolean = True)
    ' Логирует завершение операции
    If isSuccess Then
        LogSuccess "<<< ОПЕРАЦИЯ ЗАВЕРШЕНА УСПЕШНО: " & operationName & " >>>"
    Else
        LogError "<<< ОПЕРАЦИЯ ЗАВЕРШЕНА С ОШИБКАМИ: " & operationName & " >>>"
    End If
End Sub

' ===== СПЕЦИАЛИЗИРОВАННЫЕ ФУНКЦИИ =====

Public Sub LogDataProcessing(recordCount As Integer, successCount As Integer, errorCount As Integer)
    ' Логирует результаты обработки данных
    LogInfo "=== РЕЗУЛЬТАТЫ ОБРАБОТКИ ДАННЫХ ==="
    LogInfo "Общее количество записей: " & recordCount
    LogSuccess "Успешно обработано: " & successCount
    
    If errorCount > 0 Then
        LogError "Ошибок при обработке: " & errorCount
    Else
        LogSuccess "Ошибок не найдено!"
    End If
    
    Dim successRate As Double
    If recordCount > 0 Then
        successRate = (successCount / recordCount) * 100
        LogInfo "Процент успешности: " & Format(successRate, "0.0") & "%"
    End If
    
    LogInfo "=== КОНЕЦ РЕЗУЛЬТАТОВ ОБРАБОТКИ ==="
End Sub

Public Sub LogWarehouseOperation(warehouse As String, operation As String, details As String)
    ' Логирует операции с ангарами
    LogInfo "[АНГАР " & warehouse & "] " & operation & ": " & details
End Sub

Public Sub LogValidationError(fieldName As String, value As String, errorDescription As String)
    ' Логирует ошибки валидации
    LogError "ВАЛИДАЦИЯ [" & fieldName & "]: значение '" & value & "' - " & errorDescription
End Sub

' ===== ВНУТРЕННИЕ ФУНКЦИИ =====

Private Sub WriteLog(logType As String, message As String, logLevel As Integer)
    ' Основная функция записи лога
    If Not loggerInitialized Then Call InitializeLogger  ' ИСПРАВЛЕНО
    
    ' Проверяем уровень логирования
    If logLevel < currentLogLevel Then Exit Sub
    
    ' Формируем сообщение
    Dim timestamp As String
    timestamp = Format(Now, "hh:mm:ss.000")
    
    Dim formattedMessage As String
    formattedMessage = "[" & timestamp & "] " & logType & ": " & message
    
    ' Записываем в буфер
    AddToBuffer formattedMessage
    
    ' Выводим в Immediate Window
    If immediateWindowEnabled Then  ' ИСПРАВЛЕНО
        Debug.Print formattedMessage
    End If
End Sub

Private Sub AddToBuffer(message As String)
    ' Добавляет сообщение в буфер лога
    logBuffer = logBuffer & message & vbNewLine
    
    ' Проверяем размер буфера
    If Len(logBuffer) > maxBufferSize Then
        ' Обрезаем буфер, оставляя последние записи
        Dim lines() As String
        lines = Split(logBuffer, vbNewLine)
        
        Dim newBuffer As String
        newBuffer = ""
        
        ' Берем последние 75% строк
        Dim startIndex As Integer
        startIndex = Int(UBound(lines) * 0.25)
        
        Dim i As Integer
        For i = startIndex To UBound(lines)
            newBuffer = newBuffer & lines(i) & vbNewLine
        Next i
        
        logBuffer = "... [БУФЕР ОБРЕЗАН] ..." & vbNewLine & newBuffer
    End If
End Sub

Private Function GetLogLevelName(logLevel As Integer) As String
    ' Возвращает название уровня логирования
    Select Case logLevel
        Case LOG_LEVEL_DEBUG: GetLogLevelName = "DEBUG"
        Case LOG_LEVEL_INFO: GetLogLevelName = "INFO"
        Case LOG_LEVEL_WARNING: GetLogLevelName = "WARNING"
        Case LOG_LEVEL_ERROR: GetLogLevelName = "ERROR"
        Case LOG_LEVEL_SUCCESS: GetLogLevelName = "SUCCESS"
        Case Else: GetLogLevelName = "UNKNOWN"
    End Select
End Function

' ===== ФУНКЦИИ РАБОТЫ С БУФЕРОМ =====

Public Function GetLogBuffer() As String
    ' Возвращает содержимое буфера лога
    GetLogBuffer = logBuffer
End Function

Public Sub ClearLogBuffer()
    ' Очищает буфер лога
    logBuffer = ""
    LogInfo "Буфер лога очищен"
End Sub

Public Function GetLogBufferSize() As Long
    ' Возвращает размер буфера в символах
    GetLogBufferSize = Len(logBuffer)
End Function

Public Sub ShowLogBuffer()
    ' Показывает содержимое буфера в окне
    If logBuffer = "" Then
        MsgBox "Буфер лога пуст", vbInformation, "Лог системы"
    Else
        ' Создаем форму для отображения лога (упрощенный вариант)
        Dim logLines() As String
        logLines = Split(logBuffer, vbNewLine)
        
        Dim displayText As String
        displayText = "=== ЛОГ СИСТЕМЫ ===" & vbNewLine & vbNewLine
        
        ' Показываем последние 20 строк
        Dim startIndex As Integer
        startIndex = IIf(UBound(logLines) > 20, UBound(logLines) - 20, 0)
        
        Dim i As Integer
        For i = startIndex To UBound(logLines)
            If Trim(logLines(i)) <> "" Then
                displayText = displayText & logLines(i) & vbNewLine
            End If
        Next i
        
        If UBound(logLines) > 20 Then
            displayText = "... показаны последние 20 записей ..." & vbNewLine & vbNewLine & displayText
        End If
        
        MsgBox displayText, vbInformation, "Лог системы"
    End If
End Sub

Public Sub SaveLogToFile(Optional filePath As String = "")
    ' Сохраняет лог в файл
    On Error GoTo ErrorHandler
    
    If logBuffer = "" Then
        MsgBox "Буфер лога пуст - нечего сохранять", vbInformation, "Сохранение лога"
        Exit Sub
    End If
    
    ' Определяем путь файла
    If filePath = "" Then
        filePath = Environ("USERPROFILE") & "\Desktop\MassInput_Log_" & _
                   Format(Now, "yyyy-mm-dd_hh-mm-ss") & ".txt"
    End If
    
    ' Создаем содержимое файла
    Dim fileContent As String
    fileContent = "=== ЛОГ СИСТЕМЫ МАССОВОГО ВВОДА ТОВАРОВ ===" & vbNewLine
    fileContent = fileContent & "Дата и время создания: " & Format(Now, "dd.mm.yyyy hh:mm:ss") & vbNewLine
    fileContent = fileContent & "Версия Excel: " & Application.version & vbNewLine
    
    If ActiveWorkbook Is Nothing Then
        fileContent = fileContent & "Активная книга: НЕ НАЙДЕНА" & vbNewLine
    Else
        fileContent = fileContent & "Активная книга: " & ActiveWorkbook.Name & vbNewLine
    End If
    
    fileContent = fileContent & "=" & String(50, "=") & vbNewLine & vbNewLine
    fileContent = fileContent & logBuffer
    
    ' Записываем в файл
    Dim fileNumber As Integer
    fileNumber = FreeFile
    Open filePath For Output As #fileNumber
    Print #fileNumber, fileContent
    Close #fileNumber
    
    LogSuccess "Лог сохранен в файл: " & filePath
    MsgBox "Лог успешно сохранен в файл:" & vbNewLine & filePath, vbInformation, "Сохранение лога"
    Exit Sub
    
ErrorHandler:
    LogError "Ошибка сохранения лога: " & Err.description
    MsgBox "Ошибка при сохранении лога: " & Err.description, vbCritical, "Ошибка сохранения"
End Sub

' ===== ОТЛАДОЧНЫЕ И СЛУЖЕБНЫЕ ФУНКЦИИ =====

Public Sub TestLogger()
    ' Тестирует систему логирования
    Call InitializeLogger(True, LOG_LEVEL_DEBUG)
    
    LogDebug "Это отладочное сообщение"
    LogInfo "Это информационное сообщение"
    LogWarning "Это предупреждение"
    LogError "Это сообщение об ошибке"
    LogSuccess "Это сообщение об успехе"
    
    LogOperationStart "Тестовая операция"
    LogWarehouseOperation "6", "Запись товара", "Лерашанс в ячейку Е/2ряд"
    LogValidationError "Количество", "-50", "не может быть отрицательным"
    LogDataProcessing 100, 95, 5
    LogOperationEnd "Тестовая операция", True
    
    MsgBox "Тест системы логирования завершен!" & vbNewLine & _
           "Проверьте Immediate Window (Ctrl+G) для просмотра результатов.", _
           vbInformation, "Тест логирования"
End Sub

Public Function GetLogStatistics() As String
    ' Возвращает статистику по логам
    If logBuffer = "" Then
        GetLogStatistics = "Буфер лога пуст"
        Exit Function
    End If
    
    Dim lines() As String
    lines = Split(logBuffer, vbNewLine)
    
    Dim debugCount, infoCount, warningCount, errorCount, successCount As Integer
    debugCount = 0: infoCount = 0: warningCount = 0: errorCount = 0: successCount = 0
    
    Dim i As Integer
    For i = 0 To UBound(lines)
        If InStr(lines(i), "DEBUG:") > 0 Then debugCount = debugCount + 1
        If InStr(lines(i), "INFO:") > 0 Then infoCount = infoCount + 1
        If InStr(lines(i), "WARNING:") > 0 Then warningCount = warningCount + 1
        If InStr(lines(i), "ERROR:") > 0 Then errorCount = errorCount + 1
        If InStr(lines(i), "SUCCESS:") > 0 Then successCount = successCount + 1
    Next i
    
    Dim stats As String
    stats = "=== СТАТИСТИКА ЛОГИРОВАНИЯ ===" & vbNewLine
    stats = stats & "Общее количество записей: " & (UBound(lines) + 1) & vbNewLine
    stats = stats & "DEBUG: " & debugCount & vbNewLine
    stats = stats & "INFO: " & infoCount & vbNewLine
    stats = stats & "WARNING: " & warningCount & vbNewLine
    stats = stats & "ERROR: " & errorCount & vbNewLine
    stats = stats & "SUCCESS: " & successCount & vbNewLine
    stats = stats & "Размер буфера: " & Format(Len(logBuffer), "#,##0") & " символов"
    
    GetLogStatistics = stats
End Function

Public Sub ShowLogStatistics()
    ' Показывает статистику логирования
    MsgBox GetLogStatistics(), vbInformation, "Статистика логирования"
End Sub

' ===== ФУНКЦИИ КОНФИГУРАЦИИ =====

Public Sub SetMaxBufferSize(newSize As Long)
    ' Устанавливает максимальный размер буфера
    If newSize < 1000 Then newSize = 1000
    If newSize > 500000 Then newSize = 500000
    
    maxBufferSize = newSize
    LogInfo "Максимальный размер буфера установлен: " & Format(newSize, "#,##0") & " символов"
End Sub

Public Function GetMaxBufferSize() As Long
    ' Возвращает максимальный размер буфера
    GetMaxBufferSize = maxBufferSize
End Function

Public Function IsLoggerInitialized() As Boolean  ' ИСПРАВЛЕНО: переименовано из IsInitialized
    ' Проверяет, инициализирована ли система логирования
    IsLoggerInitialized = loggerInitialized  ' ИСПРАВЛЕНО
End Function

Public Sub ResetLogger()
    ' Сбрасывает систему логирования
    loggerInitialized = False  ' ИСПРАВЛЕНО
    currentLogLevel = LOG_LEVEL_INFO
    immediateWindowEnabled = True  ' ИСПРАВЛЕНО
    logBuffer = ""
    maxBufferSize = 50000
    
    Debug.Print "[RESET] Система логирования сброшена"
End Sub

