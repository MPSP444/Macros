Attribute VB_Name = "ModuleHeaderOptimizer"


' ===== ModuleHeaderOptimizer - ПРОСТАЯ ПРАВИЛЬНАЯ ВЕРСИЯ С ПЕРЕНОСОМ ТЕКСТА =====
' Версия 6.1 - Только соседние заголовки, простая логика + WrapText
Option Explicit

Public Sub OptimizeAllWarehouseHeaders()
    ' Простая оптимизация - только соседние одинаковые заголовки
    On Error GoTo ErrorHandler
    
    If Not ModuleLogger.IsLoggerInitialized() Then
        Call ModuleLogger.InitializeLogger(True, ModuleLogger.LOG_LEVEL_INFO)
    End If
    
    ModuleLogger.LogOperationStart "Простая оптимизация заголовков v6.1 (только соседние + перенос текста)"
    
    Application.ScreenUpdating = False
    
    Dim totalOptimized As Integer
    totalOptimized = 0
    
    ' Обрабатываем каждый ангар
    Dim i As Integer
    For i = 5 To 12
        ModuleLogger.LogMessage "=== Оптимизация ангара " & i & " ==="
        
        Dim optimizedCount As Integer
        Call OptimizeWarehouseSimple(CStr(i), optimizedCount)
        totalOptimized = totalOptimized + optimizedCount
        
        ModuleLogger.LogSuccess "Ангар " & i & " оптимизирован (" & optimizedCount & " объединений)"
    Next i
    
    Application.ScreenUpdating = True
    
    MsgBox "? Простая оптимизация завершена!" & vbNewLine & _
           "Объединено групп: " & totalOptimized & vbNewLine & _
           "Объединялись только соседние одинаковые заголовки." & vbNewLine & _
           "? Перенос текста включен для всех заголовков!", _
           vbInformation, "Оптимизация завершена"
    
    ModuleLogger.LogOperationEnd "Простая оптимизация заголовков", True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    ModuleLogger.LogError "Ошибка оптимизации: " & Err.description
    MsgBox "Ошибка: " & Err.description, vbCritical, "Ошибка"
End Sub

Private Sub OptimizeWarehouseSimple(warehouse As String, ByRef optimizedCount As Integer)
    ' Простая оптимизация одного ангара
    On Error Resume Next
    
    optimizedCount = 0
    
    ' Получаем конфигурацию
    Dim config As ModuleTypes.WarehouseConfig
    config = ModuleConfig.GetWarehouseConfig(warehouse)
    
    ' Получаем лист
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(config.sheetName)
    If ws Is Nothing Then Exit Sub
    
    ' Оптимизируем верхнюю секцию
    If config.UpperHeaderRow > 0 Then
        Dim upperCount As Integer
        Call OptimizeRowSimple(ws, config.UpperHeaderRow, "верхняя", upperCount)
        optimizedCount = optimizedCount + upperCount
    End If
    
    ' Оптимизируем нижнюю секцию
    If config.LowerHeaderRow > 0 Then
        Dim lowerCount As Integer
        Call OptimizeRowSimple(ws, config.LowerHeaderRow, "нижняя", lowerCount)
        optimizedCount = optimizedCount + lowerCount
    End If
    
    On Error GoTo 0
End Sub

Private Sub OptimizeRowSimple(ws As Worksheet, headerRow As Long, sectionName As String, ByRef optimizedCount As Integer)
    ' ПРОСТАЯ ЛОГИКА: идем слева направо, объединяем только соседние
    On Error Resume Next
    
    optimizedCount = 0
    
    ModuleLogger.LogMessage "Анализ " & sectionName & " секции (строка " & headerRow & ")"
    
    ' Идем по нечетным столбцам (заголовки) до ряда 63
    Dim currentCol As Integer
    currentCol = 3 ' Начинаем с ряда 1 (столбец C)
    
    Do While currentCol <= 127 ' До ряда 63
        ' Получаем текущий заголовок
        Dim currentHeader As String
        currentHeader = GetHeaderText(ws, headerRow, currentCol)
        
        ' Если заголовок пустой или конфликтный - пропускаем
        If currentHeader = "" Or InStr(currentHeader, "+") > 0 Then
            currentCol = currentCol + 2 ' Переходим к следующему ряду
            GoTo NextRow
        End If
        
        ' Ищем конец последовательности СОСЕДНИХ одинаковых заголовков
        Dim sequenceEnd As Integer
        sequenceEnd = FindNeighborSequenceEnd(ws, headerRow, currentCol, currentHeader)
        
        ' Если нашли последовательность из 2+ рядов - объединяем
        If sequenceEnd > currentCol Then
            Call MergeSimpleRange(ws, headerRow, currentCol, sequenceEnd, currentHeader, sectionName)
            optimizedCount = optimizedCount + 1
            
            ' Переходим за объединенную область
            currentCol = sequenceEnd + 2
        Else
            ' Переходим к следующему ряду
            currentCol = currentCol + 2
        End If
        
NextRow:
    Loop
    
    ModuleLogger.LogMessage "Секция " & sectionName & ": " & optimizedCount & " объединений"
    
    On Error GoTo 0
End Sub

Private Function FindNeighborSequenceEnd(ws As Worksheet, headerRow As Long, startCol As Integer, targetHeader As String) As Integer
    ' Ищем конец последовательности ТОЛЬКО соседних одинаковых заголовков
    Dim currentCol As Integer
    currentCol = startCol
    
    ' Проверяем следующие ряды подряд
    Do While currentCol <= 125 ' Не выходим за границы
        Dim nextCol As Integer
        nextCol = currentCol + 2 ' Следующий ряд
        
        Dim nextHeader As String
        nextHeader = GetHeaderText(ws, headerRow, nextCol)
        
        ' Если следующий заголовок такой же - продолжаем
        If UCase(Trim(nextHeader)) = UCase(Trim(targetHeader)) And nextHeader <> "" Then
            currentCol = nextCol
            ModuleLogger.LogDebug "  Продлили: " & targetHeader & " до ряда " & ((nextCol - 3) / 2 + 1)
        Else
            ' Как только встретили другой заголовок или пустоту - останавливаемся
            Exit Do
        End If
    Loop
    
    FindNeighborSequenceEnd = currentCol
End Function

Private Sub MergeSimpleRange(ws As Worksheet, headerRow As Long, startCol As Integer, endCol As Integer, headerText As String, sectionName As String)
    ' Простое объединение соседних заголовков с ПЕРЕНОСОМ ТЕКСТА
    On Error Resume Next
    
    ' Добавляем +1 чтобы включить столбец партии последнего ряда
    Dim actualEndCol As Integer
    actualEndCol = endCol + 1
    
    ' Вычисляем номера рядов
    Dim startRow As Integer
    Dim endRow As Integer
    startRow = (startCol - 3) / 2 + 1
    endRow = (endCol - 3) / 2 + 1
    
    ModuleLogger.LogMessage "Объединяем '" & headerText & "' ряды " & startRow & "-" & endRow & " (столбцы " & startCol & "-" & actualEndCol & ")"
    
    ' Определяем диапазон
    Dim mergeRange As Range
    Set mergeRange = ws.Range(ws.Cells(headerRow, startCol), ws.Cells(headerRow, actualEndCol))
    
    ' Разъединяем существующие объединения
    mergeRange.UnMerge
    
    ' Очищаем содержимое
    mergeRange.ClearContents
    
    ' Объединяем ячейки
    mergeRange.Merge
    
    ' Форматируем с ПЕРЕНОСОМ ТЕКСТА
    With ws.Cells(headerRow, startCol)
        .value = headerText
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True                    ' ? ДОБАВЛЕНО: Перенос текста
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(144, 238, 144) ' Зеленый как было
        .Font.Color = RGB(0, 100, 0)
        
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(0, 150, 0)
        End With
    End With
    
    ModuleLogger.LogSuccess "? Объединен '" & headerText & "' ряды " & startRow & "-" & endRow & " + перенос текста"
    
    On Error GoTo 0
End Sub

Private Function GetHeaderText(ws As Worksheet, row As Long, col As Integer) As String
    ' Безопасно получаем текст заголовка
    On Error Resume Next
    
    Dim cellValue As Variant
    cellValue = ws.Cells(row, col).value
    
    If isEmpty(cellValue) Or IsNull(cellValue) Then
        GetHeaderText = ""
    Else
        GetHeaderText = Trim(CStr(cellValue))
    End If
    
    On Error GoTo 0
End Function

' ===== ДОПОЛНИТЕЛЬНАЯ ФУНКЦИЯ ДЛЯ ПРИМЕНЕНИЯ ПЕРЕНОСА К СУЩЕСТВУЮЩИМ ЗАГОЛОВКАМ =====

Public Sub ApplyWrapTextToAllExistingHeaders()
    ' Применяет перенос текста ко всем существующим заголовкам во всех ангарах
    On Error GoTo ErrorHandler
    
    If Not ModuleLogger.IsLoggerInitialized() Then
        Call ModuleLogger.InitializeLogger(True, ModuleLogger.LOG_LEVEL_INFO)
    End If
    
    ModuleLogger.LogOperationStart "Применение переноса текста к существующим заголовкам"
    
    Application.ScreenUpdating = False
    
    Dim processedHeaders As Integer
    processedHeaders = 0
    
    ' Обрабатываем каждый ангар
    Dim i As Integer
    For i = 5 To 12
        ModuleLogger.LogMessage "Обработка заголовков ангара " & i
        
        On Error Resume Next
        Dim config As ModuleTypes.WarehouseConfig
        config = ModuleConfig.GetWarehouseConfig(CStr(i))
        
        Dim ws As Worksheet
        Set ws = ActiveWorkbook.Worksheets(config.sheetName)
        
        If Not ws Is Nothing Then
            ' Обрабатываем верхнюю секцию
            If config.UpperHeaderRow > 0 Then
                Dim upperCount As Integer
                upperCount = ApplyWrapToHeaderRow(ws, config.UpperHeaderRow, "верхняя секция ангара " & i)
                processedHeaders = processedHeaders + upperCount
            End If
            
            ' Обрабатываем нижнюю секцию
            If config.LowerHeaderRow > 0 Then
                Dim lowerCount As Integer
                lowerCount = ApplyWrapToHeaderRow(ws, config.LowerHeaderRow, "нижняя секция ангара " & i)
                processedHeaders = processedHeaders + lowerCount
            End If
        End If
        On Error GoTo ErrorHandler
    Next i
    
    Application.ScreenUpdating = True
    
    MsgBox "? Перенос текста применен!" & vbNewLine & _
           "Обработано заголовков: " & processedHeaders & vbNewLine & _
           "Теперь все заголовки поддерживают перенос текста.", _
           vbInformation, "Применение переноса завершено"
    
    ModuleLogger.LogOperationEnd "Применение переноса текста", True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    ModuleLogger.LogError "Ошибка применения переноса: " & Err.description
    MsgBox "Ошибка: " & Err.description, vbCritical, "Ошибка"
End Sub

Private Function ApplyWrapToHeaderRow(ws As Worksheet, headerRow As Long, sectionName As String) As Integer
    ' Применяет перенос текста ко всем заголовкам в строке
    On Error Resume Next
    
    Dim processedCount As Integer
    processedCount = 0
    
    ' Проходим по всем столбцам заголовков (C, E, G, I... до ряда 63)
    Dim col As Integer
    For col = 3 To 127 Step 2 ' Столбцы C, E, G, I... (через 2)
        Dim cellValue As String
        cellValue = Trim(CStr(ws.Cells(headerRow, col).value))
        
        ' Если в ячейке есть текст - применяем перенос
        If cellValue <> "" Then
            ws.Cells(headerRow, col).WrapText = True
            processedCount = processedCount + 1
        End If
    Next col
    
    If processedCount > 0 Then
        ModuleLogger.LogInfo "Применен перенос к " & processedCount & " заголовкам в " & sectionName & " (строка " & headerRow & ")"
    End If
    
    ApplyWrapToHeaderRow = processedCount
    On Error GoTo 0
End Function

' ===== ФУНКЦИЯ ДЛЯ ОПТИМИЗАЦИИ ОДНОГО АНГАРА =====

Public Sub OptimizeSpecificWarehouse(warehouse As String)
    ' Оптимизация конкретного ангара
    On Error GoTo ErrorHandler
    
    If Not ModuleLogger.IsLoggerInitialized() Then
        Call ModuleLogger.InitializeLogger(True, ModuleLogger.LOG_LEVEL_INFO)
    End If
    
    ModuleLogger.LogOperationStart "Оптимизация ангара " & warehouse
    
    Application.ScreenUpdating = False
    
    Dim optimizedCount As Integer
    Call OptimizeWarehouseSimple(warehouse, optimizedCount)
    
    Application.ScreenUpdating = True
    
    MsgBox "? Ангар " & warehouse & " оптимизирован!" & vbNewLine & _
           "Объединено групп: " & optimizedCount & vbNewLine & _
           "? Перенос текста включен!", _
           vbInformation, "Оптимизация завершена"
    
    ModuleLogger.LogOperationEnd "Оптимизация ангара " & warehouse, True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    ModuleLogger.LogError "Ошибка оптимизации ангара " & warehouse & ": " & Err.description
    MsgBox "Ошибка: " & Err.description, vbCritical, "Ошибка"
End Sub

' ===== ТЕСТОВЫЕ ФУНКЦИИ =====

Public Sub TestSimpleOptimize()
    ' Тест простой оптимизации
    Dim warehouse As String
    warehouse = InputBox("Введите номер ангара для простого теста (5-12):", "Простой тест", "8")
    
    If warehouse = "" Then Exit Sub
    
    If IsNumeric(warehouse) And CInt(warehouse) >= 5 And CInt(warehouse) <= 12 Then
        Call ModuleLogger.InitializeLogger(True, ModuleLogger.LOG_LEVEL_DEBUG)
        
        ModuleLogger.LogMessage "=== ПРОСТОЙ ТЕСТ АНГАРА " & warehouse & " ==="
        
        Dim optimizedCount As Integer
        Call OptimizeWarehouseSimple(warehouse, optimizedCount)
        
        MsgBox "Простой тест завершен!" & vbNewLine & _
               "Ангар: " & warehouse & vbNewLine & _
               "Объединений: " & optimizedCount & vbNewLine & _
               "Только соседние одинаковые заголовки." & vbNewLine & _
               "? Перенос текста включен!", vbInformation
    Else
        MsgBox "Неверный номер ангара!", vbExclamation
    End If
End Sub

Public Sub ShowOptimizationPreview(warehouse As String)
    ' Превью оптимизации (без изменений)
    On Error GoTo ErrorHandler
    
    If Not ModuleLogger.IsLoggerInitialized() Then
        Call ModuleLogger.InitializeLogger(True, ModuleLogger.LOG_LEVEL_DEBUG)
    End If
    
    ' Получаем конфигурацию
    Dim config As ModuleTypes.WarehouseConfig
    config = ModuleConfig.GetWarehouseConfig(warehouse)
    
    ' Получаем лист
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(config.sheetName)
    If ws Is Nothing Then
        MsgBox "Лист ангара " & warehouse & " не найден!", vbExclamation
        Exit Sub
    End If
    
    Dim preview As String
    preview = "=== ПРЕВЬЮ ОПТИМИЗАЦИИ АНГАРА " & warehouse & " ===" & vbNewLine & vbNewLine
    
    ' Анализируем верхнюю секцию
    If config.UpperHeaderRow > 0 Then
        preview = preview & "?? ВЕРХНЯЯ СЕКЦИЯ (строка " & config.UpperHeaderRow & "):" & vbNewLine
        preview = preview & AnalyzeHeaderRowForPreview(ws, config.UpperHeaderRow, "верхняя")
        preview = preview & vbNewLine
    End If
    
    ' Анализируем нижнюю секцию
    If config.LowerHeaderRow > 0 Then
        preview = preview & "?? НИЖНЯЯ СЕКЦИЯ (строка " & config.LowerHeaderRow & "):" & vbNewLine
        preview = preview & AnalyzeHeaderRowForPreview(ws, config.LowerHeaderRow, "нижняя")
    End If
    
    preview = preview & vbNewLine & "? После оптимизации всем заголовкам будет добавлен перенос текста!"
    
    MsgBox preview, vbInformation, "Превью оптимизации"
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка превью: " & Err.description, vbCritical, "Ошибка"
End Sub

Private Function AnalyzeHeaderRowForPreview(ws As Worksheet, headerRow As Long, sectionName As String) As String
    ' Анализирует строку заголовков для превью
    On Error Resume Next
    
    Dim result As String
    result = ""
    
    Dim currentCol As Integer
    currentCol = 3 ' Начинаем с ряда 1 (столбец C)
    
    Dim groupsFound As Integer
    groupsFound = 0
    
    Do While currentCol <= 127 ' До ряда 63
        Dim currentHeader As String
        currentHeader = GetHeaderText(ws, headerRow, currentCol)
        
        If currentHeader <> "" And InStr(currentHeader, "+") = 0 Then
            ' Ищем последовательность
            Dim sequenceEnd As Integer
            sequenceEnd = FindNeighborSequenceEnd(ws, headerRow, currentCol, currentHeader)
            
            If sequenceEnd > currentCol Then
                ' Нашли группу для объединения
                Dim startRowNum As Integer
                Dim endRowNum As Integer
                startRowNum = (currentCol - 3) / 2 + 1
                endRowNum = (sequenceEnd - 3) / 2 + 1
                
                result = result & "  ? '" & currentHeader & "' ряды " & startRowNum & "-" & endRowNum & vbNewLine
                groupsFound = groupsFound + 1
                
                currentCol = sequenceEnd + 2
            Else
                currentCol = currentCol + 2
            End If
        Else
            currentCol = currentCol + 2
        End If
    Loop
    
    If groupsFound = 0 Then
        result = "  ??  Соседних одинаковых заголовков не найдено" & vbNewLine
    Else
        result = "  ?? Будет объединено групп: " & groupsFound & vbNewLine & result
    End If
    
    AnalyzeHeaderRowForPreview = result
    On Error GoTo 0
End Function

Public Sub ShowSimpleInfo()
    ' Информация о простой системе
    Dim info As String
    info = "=== ПРОСТАЯ СИСТЕМА ОПТИМИЗАЦИИ v6.1 ===" & vbNewLine & vbNewLine
    
    info = info & "?? ПРИНЦИП РАБОТЫ:" & vbNewLine
    info = info & "• Идет слева направо по заголовкам" & vbNewLine
    info = info & "• Объединяет ТОЛЬКО соседние одинаковые" & vbNewLine
    info = info & "• Пропускает пустые ряды" & vbNewLine
    info = info & "• Останавливается при встрече другого товара" & vbNewLine
    info = info & "• ? Устанавливает перенос текста для всех заголовков" & vbNewLine & vbNewLine
    
    info = info & "?? ПРИМЕР:" & vbNewLine
    info = info & "Ряд 1: Лерашанс  <" & vbNewLine
    info = info & "Ряд 2: Лерашанс  > Объединить + перенос текста" & vbNewLine
    info = info & "Ряд 3: пусто     < Пропустить" & vbNewLine
    info = info & "Ряд 4: Дикошанс  < НЕ объединять" & vbNewLine
    info = info & "Ряд 5: Лерашанс  < НЕ объединять" & vbNewLine & vbNewLine
    
    info = info & "?? ДИАПАЗОН: ряды 1-63" & vbNewLine
    info = info & "?? ФУНКЦИИ:" & vbNewLine
    info = info & "• OptimizeAllWarehouseHeaders() - все ангары" & vbNewLine
    info = info & "• OptimizeSpecificWarehouse() - один ангар" & vbNewLine
    info = info & "• ApplyWrapTextToAllExistingHeaders() - только перенос" & vbNewLine
    info = info & "• TestSimpleOptimize() - тест одного ангара" & vbNewLine
    
    MsgBox info, vbInformation, "Простая система v6.1"
End Sub

