Attribute VB_Name = "ModuleHeaderManager"


' ===== ModuleHeaderManager - Управление заголовками товаров С ПЕРЕНОСОМ ТЕКСТА =====
Option Explicit

Public Function ProcessHeaders(groupedData As Object) As String
    On Error GoTo ErrorHandler
    
    ModuleLogger.LogMessage "=== НАЧАЛО ОБРАБОТКИ ЗАГОЛОВКОВ ==="
    
    Dim errorLog As String
    errorLog = ""
    
    ' Группируем данные по ангарам
    Dim warehouseGroups As Object
    Set warehouseGroups = GroupByWarehouses(groupedData)
    
    ' Обрабатываем каждый ангар
    Dim warehouse As Variant
    For Each warehouse In warehouseGroups.keys
        ModuleLogger.LogMessage "Обработка заголовков для ангара " & warehouse
        
        Dim warehouseResult As String
        warehouseResult = ProcessWarehouseHeaders(CStr(warehouse), warehouseGroups(warehouse))
        
        If warehouseResult <> "" Then
            errorLog = errorLog & warehouseResult & vbNewLine
        End If
    Next warehouse
    
    ModuleLogger.LogMessage "=== ЗАВЕРШЕНИЕ ОБРАБОТКИ ЗАГОЛОВКОВ ==="
    ProcessHeaders = errorLog
    Exit Function
    
ErrorHandler:
    ModuleLogger.LogError "Критическая ошибка в ProcessHeaders: " & Err.description
    ProcessHeaders = "Критическая ошибка обработки заголовков: " & Err.description
End Function

Private Function GroupByWarehouses(groupedData As Object) As Object
    ' ОБНОВЛЕННАЯ ФУНКЦИЯ: учитывает новый формат ключей с секциями
    Dim warehouseGroups As Object
    Set warehouseGroups = CreateObject("Scripting.Dictionary")
    
    Dim key As Variant
    For Each key In groupedData.keys
        ' НОВАЯ ЛОГИКА: ключ имеет формат "warehouse-row-section"
        Dim parts() As String
        parts = Split(CStr(key), "-")
        
        If UBound(parts) >= 2 Then
            Dim warehouse As String
            warehouse = parts(0)
            
            If Not warehouseGroups.exists(warehouse) Then
                Set warehouseGroups(warehouse) = CreateObject("Scripting.Dictionary")
            End If
            
            warehouseGroups(warehouse).Add key, groupedData(key)
        End If
    Next key
    
    Set GroupByWarehouses = warehouseGroups
End Function

Private Function ProcessWarehouseHeaders(warehouse As String, warehouseData As Object) As String
    On Error GoTo ErrorHandler
    
    ' Получаем конфигурацию ангара
    Dim config As ModuleTypes.WarehouseConfig
    config = ModuleConfig.GetWarehouseConfig(warehouse)
    
    ' Проверяем существование листа
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(config.sheetName)
    
    ' Анализируем товары по рядам и секциям
    Dim rowAnalysis As Object
    Set rowAnalysis = AnalyzeRowProducts(warehouseData)
    
    ' Определяем, какие секции нужно обработать
    Dim upperSectionResult As String
    upperSectionResult = ProcessSection(ws, rowAnalysis, config.UpperHeaderRow, "верхняя", warehouse)
    
    Dim lowerSectionResult As String
    lowerSectionResult = ProcessSection(ws, rowAnalysis, config.LowerHeaderRow, "нижняя", warehouse)
    
    Dim result As String
    result = ""
    If upperSectionResult <> "" Then result = result & upperSectionResult & vbNewLine
    If lowerSectionResult <> "" Then result = result & lowerSectionResult & vbNewLine
    
    ProcessWarehouseHeaders = result
    Exit Function
    
ErrorHandler:
    ModuleLogger.LogError "Ошибка в ProcessWarehouseHeaders для ангара " & warehouse & ": " & Err.description
    ProcessWarehouseHeaders = "Ошибка обработки ангара " & warehouse & ": " & Err.description
End Function

Private Function AnalyzeRowProducts(warehouseData As Object) As Object
    ' ОБНОВЛЕННАЯ ФУНКЦИЯ: учитывает секции в ключах
    Dim analysis As Object
    Set analysis = CreateObject("Scripting.Dictionary")
    
    Dim key As Variant
    For Each key In warehouseData.keys
        ' НОВАЯ ЛОГИКА: ключ имеет формат "warehouse-row-section"
        Dim parts() As String
        parts = Split(CStr(key), "-")
        
        If UBound(parts) >= 2 Then
            Dim rowSectionKey As String
            rowSectionKey = parts(1) & "-" & parts(2)  ' "row-section"
            
            ' Получаем товары в этом ряду и секции
            Dim productDict As Object
            Set productDict = warehouseData(key)
            
            ' Создаем список товаров
            Dim products As String
            products = ""
            Dim HasConflict As Boolean
            HasConflict = False
            
            Dim product As Variant
            For Each product In productDict.keys
                If products <> "" Then
                    products = products & " + "
                    HasConflict = True ' Если больше одного товара, то есть конфликт
                End If
                products = products & CStr(product)
            Next product
            
            ' Сохраняем анализ ряда-секции
            Dim rowInfo As Object
            Set rowInfo = CreateObject("Scripting.Dictionary")
            rowInfo("Products") = products
            rowInfo("HasConflict") = HasConflict
            rowInfo("ProductCount") = productDict.count
            rowInfo("Section") = parts(2)  ' UPPER или LOWER
            
            analysis.Add rowSectionKey, rowInfo
            
            ModuleLogger.LogDebug "Ряд " & parts(1) & " (секция " & ModuleConfig.GetSectionName(parts(2)) & "): " & products & " (конфликт: " & HasConflict & ")"
        End If
    Next key
    
    Set AnalyzeRowProducts = analysis
End Function

Private Function ProcessSection(ws As Worksheet, rowAnalysis As Object, headerRow As Long, sectionName As String, warehouse As String) As String
    ' ОБНОВЛЕННАЯ ФУНКЦИЯ: обрабатывает только нужную секцию
    On Error GoTo ErrorHandler
    
    If headerRow = 0 Then
        ' Секция не используется
        ProcessSection = ""
        Exit Function
    End If
    
    ModuleLogger.LogMessage "Обработка " & sectionName & " секции ангара " & warehouse & " (строка " & headerRow & ")"
    
    Dim errorLog As String
    errorLog = ""
    
    ' Определяем максимальное количество рядов
    Dim maxRows As Integer
    maxRows = GetMaxRowsForWarehouse(warehouse)
    
    ' НОВАЯ ЛОГИКА: обрабатываем только ряды нужной секции
    Dim targetSection As String
    If sectionName = "верхняя" Then
        targetSection = "UPPER"
    Else
        targetSection = "LOWER"
    End If
    
    ' Проходим по всем рядам и ищем подходящие секции
    Dim currentRow As Integer
    For currentRow = 1 To maxRows
        Dim rowSectionKey As String
        rowSectionKey = CStr(currentRow) & "-" & targetSection
        
        If rowAnalysis.exists(rowSectionKey) Then
            Dim rowInfo As Object
            Set rowInfo = rowAnalysis(rowSectionKey)
            
            Dim applyResult As String
            applyResult = ApplyHeaderToRow(ws, rowInfo, headerRow, currentRow)
            
            If applyResult <> "" Then
                errorLog = errorLog & applyResult & vbNewLine
            End If
        End If
    Next currentRow
    
    ProcessSection = errorLog
    Exit Function
    
ErrorHandler:
    ModuleLogger.LogError "Ошибка в ProcessSection (" & sectionName & ") для ангара " & warehouse & ": " & Err.description
    ProcessSection = "Ошибка обработки " & sectionName & " секции ангара " & warehouse & ": " & Err.description
End Function

Private Function ApplyHeaderToRow(ws As Worksheet, rowInfo As Object, headerRow As Long, rowNumber As Integer) As String
    ' ОБНОВЛЕННАЯ ФУНКЦИЯ: применяет заголовок к конкретному ряду С ПЕРЕНОСОМ ТЕКСТА
    On Error GoTo ErrorHandler
    
    Dim startCol As Integer
    startCol = 3 + (rowNumber - 1) * 2  ' C, E, G, I...
    
    Dim endCol As Integer
    endCol = startCol + 1  ' D, F, H, J...
    
    Dim products As String
    products = rowInfo("Products")
    
    Dim HasConflict As Boolean
    HasConflict = rowInfo("HasConflict")
    
    Dim sectionName As String
    sectionName = ModuleConfig.GetSectionName(rowInfo("Section"))
    
    ' Очищаем существующие заголовки в этом диапазоне
    ws.Range(ws.Cells(headerRow, startCol), ws.Cells(headerRow, endCol)).UnMerge
    ws.Range(ws.Cells(headerRow, startCol), ws.Cells(headerRow, endCol)).ClearContents
    
    ' Объединяем ячейки для заголовка
    ws.Range(ws.Cells(headerRow, startCol), ws.Cells(headerRow, endCol)).Merge
    
    ' Устанавливаем текст заголовка
    ws.Cells(headerRow, startCol).value = products
    
    ' Форматируем заголовок С ПЕРЕНОСОМ ТЕКСТА
    With ws.Cells(headerRow, startCol)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True                    ' ? ДОБАВЛЕНО: Перенос текста
        .Font.Bold = True
        .Font.Size = 14
        
        If HasConflict Then
            ' Красный фон для конфликтных заголовков
            .Interior.Color = RGB(255, 200, 200)
            .Font.Color = RGB(150, 0, 0)
            ModuleLogger.LogWarning "Конфликт товаров в заголовке (" & sectionName & " секция): " & products
        Else
            ' Обычное форматирование
            .Interior.Color = RGB(200, 230, 255)
            .Font.Color = RGB(0, 50, 100)
        End If
        
        ' Добавляем границы
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(100, 100, 100)
        End With
    End With
    
    ModuleLogger.LogSuccess "Применен заголовок '" & products & "' к ряду " & rowNumber & _
                           " (" & sectionName & " секция, столбцы " & startCol & "-" & endCol & ", строка " & headerRow & ") + перенос текста"
    
    ApplyHeaderToRow = ""
    Exit Function
    
ErrorHandler:
    ModuleLogger.LogError "Ошибка применения заголовка: " & Err.description
    ApplyHeaderToRow = "Ошибка применения заголовка '" & products & "': " & Err.description
End Function

Private Function GetMaxRowsForWarehouse(warehouse As String) As Integer
    ' Определяем максимальное количество рядов для каждого ангара
    Select Case warehouse
        Case "5": GetMaxRowsForWarehouse = 60
        Case "6": GetMaxRowsForWarehouse = 60
        Case "7": GetMaxRowsForWarehouse = 60
        Case "8": GetMaxRowsForWarehouse = 60
        Case "9": GetMaxRowsForWarehouse = 60
        Case "10": GetMaxRowsForWarehouse = 60
        Case "11": GetMaxRowsForWarehouse = 60
        Case "12": GetMaxRowsForWarehouse = 60
        Case Else: GetMaxRowsForWarehouse = 60
    End Select
End Function

Public Sub ClearAllHeaders()
    ' Функция для очистки всех заголовков (используется в ModuleDataCleaner)
    On Error Resume Next
    
    ModuleLogger.LogMessage "=== ОЧИСТКА ВСЕХ ЗАГОЛОВКОВ ==="
    
    Dim i As Integer
    For i = 5 To 12
        Dim config As ModuleTypes.WarehouseConfig
        config = ModuleConfig.GetWarehouseConfig(CStr(i))
        
        Dim ws As Worksheet
        Set ws = ActiveWorkbook.Worksheets(config.sheetName)
        
        If Not ws Is Nothing Then
            ' Очищаем заголовки верхней секции
            If config.UpperHeaderRow > 0 Then
                ClearHeaderRow ws, config.UpperHeaderRow, "Ангар " & i & " (верхняя секция)"
            End If
            
            ' Очищаем заголовки нижней секции
            If config.LowerHeaderRow > 0 Then
                ClearHeaderRow ws, config.LowerHeaderRow, "Ангар " & i & " (нижняя секция)"
            End If
        End If
    Next i
    
    ModuleLogger.LogMessage "=== ЗАВЕРШЕНИЕ ОЧИСТКИ ЗАГОЛОВКОВ ==="
End Sub

Public Sub ClearHeaderRow(ws As Worksheet, headerRow As Long, sectionName As String)
    ' ОБНОВЛЕННАЯ ФУНКЦИЯ: очищает заголовки с учетом переноса текста
    On Error Resume Next
    
    ' Разъединяем все объединенные ячейки в строке заголовков (только в диапазоне данных C:DF)
    ws.Range("C" & headerRow & ":DF" & headerRow).UnMerge
    
    ' Очищаем содержимое (начиная с колонки C до DF, так как A и B содержат стеллажи)
    ws.Range("C" & headerRow & ":DF" & headerRow).ClearContents
    
    ' Очищаем форматирование заголовков товаров (включая WrapText)
    With ws.Range("C" & headerRow & ":DF" & headerRow)
        .Interior.ColorIndex = xlNone
        .WrapText = False  ' Сбрасываем перенос текста при очистке
    End With
    
    ModuleLogger.LogInfo "Очищены заголовки: " & sectionName & " (строка " & headerRow & ", столбцы C-DF) + сброшен перенос текста"
End Sub

' ===== ДОПОЛНИТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ ОТЛАДКИ И АНАЛИЗА =====

Public Function GetSectionStatistics(groupedData As Object) As String
    ' Функция для анализа распределения товаров по секциям
    Dim stats As String
    stats = "=== СТАТИСТИКА РАСПРЕДЕЛЕНИЯ ПО СЕКЦИЯМ ===" & vbNewLine & vbNewLine
    
    Dim upperCount As Integer
    Dim lowerCount As Integer
    upperCount = 0
    lowerCount = 0
    
    Dim key As Variant
    For Each key In groupedData.keys
        Dim parts() As String
        parts = Split(CStr(key), "-")
        
        If UBound(parts) >= 2 Then
            If parts(2) = "UPPER" Then
                upperCount = upperCount + 1
            ElseIf parts(2) = "LOWER" Then
                lowerCount = lowerCount + 1
            End If
        End If
    Next key
    
    stats = stats & "Групп в верхних секциях: " & upperCount & vbNewLine
    stats = stats & "Групп в нижних секциях: " & lowerCount & vbNewLine
    stats = stats & "Общее количество групп: " & (upperCount + lowerCount) & vbNewLine
    stats = stats & vbNewLine & "? Всем заголовкам будет установлен перенос текста"
    
    GetSectionStatistics = stats
End Function

Public Function ValidateSectionAssignment(warehouse As String, Shelf As String) As String
    ' Функция для проверки правильности определения секции
    Dim section As String
    section = ModuleConfig.DetermineSectionByShelf(Shelf, warehouse)
    
    Dim result As String
    result = "Ангар " & warehouse & ", стеллаж " & Shelf & " > " & _
             ModuleConfig.GetSectionName(section) & " секция (" & section & ")"
    
    ValidateSectionAssignment = result
End Function

Public Sub ShowSectionMappingForWarehouse(warehouse As String)
    ' Отладочная функция для показа распределения букв по секциям
    Dim info As String
    info = ModuleConfig.GetWarehouseSectionsList(warehouse)
    
    MsgBox info, vbInformation, "Распределение стеллажей по секциям"
End Sub

' ===== ФУНКЦИИ ДЛЯ РАБОТЫ С ПЕРЕНОСОМ ТЕКСТА =====

Public Sub ApplyWrapTextToAllHeaders()
    ' Применяет перенос текста ко всем существующим заголовкам
    On Error GoTo ErrorHandler
    
    If Not ModuleLogger.IsLoggerInitialized() Then
        Call ModuleLogger.InitializeLogger(True, ModuleLogger.LOG_LEVEL_INFO)
    End If
    
    ModuleLogger.LogOperationStart "Применение переноса текста ко всем заголовкам"
    
    Application.ScreenUpdating = False
    
    Dim totalHeaders As Integer
    totalHeaders = 0
    
    ' Обрабатываем каждый ангар
    Dim i As Integer
    For i = 5 To 12
        ModuleLogger.LogMessage "Применение переноса к заголовкам ангара " & i
        
        On Error Resume Next
        Dim config As ModuleTypes.WarehouseConfig
        config = ModuleConfig.GetWarehouseConfig(CStr(i))
        
        Dim ws As Worksheet
        Set ws = ActiveWorkbook.Worksheets(config.sheetName)
        
        If Not ws Is Nothing Then
            ' Обрабатываем верхнюю секцию
            If config.UpperHeaderRow > 0 Then
                Dim upperCount As Integer
                upperCount = ApplyWrapToRowHeaders(ws, config.UpperHeaderRow, "верхняя секция ангара " & i)
                totalHeaders = totalHeaders + upperCount
            End If
            
            ' Обрабатываем нижнюю секцию
            If config.LowerHeaderRow > 0 Then
                Dim lowerCount As Integer
                lowerCount = ApplyWrapToRowHeaders(ws, config.LowerHeaderRow, "нижняя секция ангара " & i)
                totalHeaders = totalHeaders + lowerCount
            End If
        End If
        On Error GoTo ErrorHandler
    Next i
    
    Application.ScreenUpdating = True
    
    MsgBox "? Перенос текста применен ко всем заголовкам!" & vbNewLine & _
           "Обработано заголовков: " & totalHeaders & vbNewLine & _
           "Теперь длинные названия товаров будут переноситься.", _
           vbInformation, "Перенос текста применен"
    
    ModuleLogger.LogOperationEnd "Применение переноса текста", True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    ModuleLogger.LogError "Ошибка применения переноса: " & Err.description
    MsgBox "Ошибка: " & Err.description, vbCritical, "Ошибка"
End Sub

Private Function ApplyWrapToRowHeaders(ws As Worksheet, headerRow As Long, sectionName As String) As Integer
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
    
    ApplyWrapToRowHeaders = processedCount
    On Error GoTo 0
End Function

Public Sub RemoveWrapTextFromAllHeaders()
    ' Убирает перенос текста со всех заголовков (если нужно отменить)
    On Error GoTo ErrorHandler
    
    If MsgBox("Убрать перенос текста со всех заголовков?", vbYesNo + vbQuestion, "Отмена переноса") = vbNo Then
        Exit Sub
    End If
    
    If Not ModuleLogger.IsLoggerInitialized() Then
        Call ModuleLogger.InitializeLogger(True, ModuleLogger.LOG_LEVEL_INFO)
    End If
    
    ModuleLogger.LogOperationStart "Удаление переноса текста из всех заголовков"
    
    Application.ScreenUpdating = False
    
    Dim totalHeaders As Integer
    totalHeaders = 0
    
    ' Обрабатываем каждый ангар
    Dim i As Integer
    For i = 5 To 12
        On Error Resume Next
        Dim config As ModuleTypes.WarehouseConfig
        config = ModuleConfig.GetWarehouseConfig(CStr(i))
        
        Dim ws As Worksheet
        Set ws = ActiveWorkbook.Worksheets(config.sheetName)
        
        If Not ws Is Nothing Then
            ' Обрабатываем верхнюю секцию
            If config.UpperHeaderRow > 0 Then
                Dim upperCount As Integer
                upperCount = RemoveWrapFromRowHeaders(ws, config.UpperHeaderRow, "верхняя секция ангара " & i)
                totalHeaders = totalHeaders + upperCount
            End If
            
            ' Обрабатываем нижнюю секцию
            If config.LowerHeaderRow > 0 Then
                Dim lowerCount As Integer
                lowerCount = RemoveWrapFromRowHeaders(ws, config.LowerHeaderRow, "нижняя секция ангара " & i)
                totalHeaders = totalHeaders + lowerCount
            End If
        End If
        On Error GoTo ErrorHandler
    Next i
    
    Application.ScreenUpdating = True
    
    MsgBox "? Перенос текста убран со всех заголовков!" & vbNewLine & _
           "Обработано заголовков: " & totalHeaders, _
           vbInformation, "Перенос текста убран"
    
    ModuleLogger.LogOperationEnd "Удаление переноса текста", True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    ModuleLogger.LogError "Ошибка удаления переноса: " & Err.description
    MsgBox "Ошибка: " & Err.description, vbCritical, "Ошибка"
End Sub

Private Function RemoveWrapFromRowHeaders(ws As Worksheet, headerRow As Long, sectionName As String) As Integer
    ' Убирает перенос текста со всех заголовков в строке
    On Error Resume Next
    
    Dim processedCount As Integer
    processedCount = 0
    
    ' Проходим по всем столбцам заголовков (C, E, G, I... до ряда 63)
    Dim col As Integer
    For col = 3 To 127 Step 2 ' Столбцы C, E, G, I... (через 2)
        Dim cellValue As String
        cellValue = Trim(CStr(ws.Cells(headerRow, col).value))
        
        ' Если в ячейке есть текст - убираем перенос
        If cellValue <> "" Then
            ws.Cells(headerRow, col).WrapText = False
            processedCount = processedCount + 1
        End If
    Next col
    
    If processedCount > 0 Then
        ModuleLogger.LogInfo "Убран перенос с " & processedCount & " заголовков в " & sectionName & " (строка " & headerRow & ")"
    End If
    
    RemoveWrapFromRowHeaders = processedCount
    On Error GoTo 0
End Function

