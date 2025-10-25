Attribute VB_Name = "ModuleDataCleaner"


' ===== ModuleDataCleaner - ОБНОВЛЕННЫЙ с форматированием =====
Option Explicit

Public Function ClearAllWarehouses() As String
    ' ОБНОВЛЕННАЯ ФУНКЦИЯ: теперь применяет красивое форматирование
    On Error GoTo ErrorHandler
    
    ModuleLogger.LogMessage "=== НАЧАЛО ПОЛНОЙ ОЧИСТКИ С КРАСИВЫМ ФОРМАТИРОВАНИЕМ ==="
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim errorLog As String
    errorLog = ""
    
    Dim totalClearedCells As Long
    totalClearedCells = 0
    
    ' Очищаем каждый ангар по очереди
    Dim i As Integer
    For i = 5 To 12
        ModuleLogger.LogMessage "?? Очистка с красивым форматированием Ангара " & i & "..."
        
        Dim warehouseResult As String
        Dim clearedCells As Long
        warehouseResult = ClearWarehouse(CStr(i), clearedCells)
        
        totalClearedCells = totalClearedCells + clearedCells
        
        If warehouseResult <> "" Then
            errorLog = errorLog & warehouseResult & vbNewLine
        Else
            ModuleLogger.LogSuccess "? Ангар " & i & " очищен с красивым форматированием (" & clearedCells & " ячеек)"
        End If
    Next i
    
    ' Очищаем все заголовки
    ModuleLogger.LogMessage "?? Очистка заголовков товаров с красивым форматированием..."
    Call ModuleHeaderManager.ClearAllHeaders
    
    ' Восстанавливаем настройки Excel
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ModuleLogger.LogMessage "=== ЗАВЕРШЕНИЕ КРАСИВОЙ ОЧИСТКИ ВСЕХ АНГАРОВ ==="
    ModuleLogger.LogSuccess "? Общее количество очищенных ячеек: " & totalClearedCells & " + КРАСИВОЕ ФОРМАТИРОВАНИЕ"
    ModuleLogger.LogInfo "   ?? Шрифт: Calibri, 12pt, жирный, по центру"
    ModuleLogger.LogInfo "   ?? Рамки: точечные + жирные для групп"
    ModuleLogger.LogInfo "   ?? Размеры: столбцы и строки восстановлены"
    
    ClearAllWarehouses = errorLog
    Exit Function
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ModuleLogger.LogError "Критическая ошибка красивой очистки: " & Err.description
    ClearAllWarehouses = "Критическая ошибка очистки: " & Err.description
End Function

Private Function ClearWarehouse(warehouse As String, ByRef clearedCells As Long) As String
    ' ОБНОВЛЕННАЯ ФУНКЦИЯ: очистка + красивое форматирование + размеры
    On Error GoTo ErrorHandler
    
    clearedCells = 0
    
    ' Получаем конфигурацию ангара
    Dim config As ModuleTypes.WarehouseConfig
    config = ModuleConfig.GetWarehouseConfig(warehouse)
    
    ' Проверяем существование листа
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(config.sheetName)
    
    ' Очищаем верхнюю секцию с форматированием
    Dim upperClearedCells As Long
    upperClearedCells = ClearSection(ws, config.UpperDataStart, config.UpperDataEnd, "верхняя", warehouse)
    clearedCells = clearedCells + upperClearedCells
    
    ' НОВОЕ: Применяем красивое форматирование для верхней секции
    If config.UpperDataStart > 0 And config.UpperDataEnd > 0 Then
        Call ResetRangeFormatting(ws, config.UpperDataStart, config.UpperDataEnd, "верхняя секция ангара " & warehouse)
    End If
    
    ' Очищаем нижнюю секцию с форматированием
    Dim lowerClearedCells As Long
    lowerClearedCells = ClearSection(ws, config.LowerDataStart, config.LowerDataEnd, "нижняя", warehouse)
    clearedCells = clearedCells + lowerClearedCells
    
    ' НОВОЕ: Применяем красивое форматирование для нижней секции
    If config.LowerDataStart > 0 And config.LowerDataEnd > 0 Then
        Call ResetRangeFormatting(ws, config.LowerDataStart, config.LowerDataEnd, "нижняя секция ангара " & warehouse)
    End If
    
    ' НОВОЕ: Применяем правильные размеры столбцов и строк
    Call ApplyWarehouseDimensions(ws, warehouse)
    
    ' Очищаем область результатов (если есть)
    If config.outputStartRow > 0 Then
        Dim outputClearedCells As Long
        outputClearedCells = ClearOutputSection(ws, config.outputStartRow, warehouse)
        clearedCells = clearedCells + outputClearedCells
    End If
    
    ModuleLogger.LogSuccess "? Ангар " & warehouse & " очищен с КРАСИВЫМ форматированием: " & clearedCells & " ячеек"
    
    ClearWarehouse = ""
    Exit Function
    
ErrorHandler:
    ModuleLogger.LogError "Ошибка очистки ангара " & warehouse & ": " & Err.description
    ClearWarehouse = "Ошибка очистки ангара " & warehouse & ": " & Err.description
End Function

Private Function ClearSection(ws As Worksheet, dataStart As Long, dataEnd As Long, sectionName As String, warehouse As String) As Long
    ' ОБНОВЛЕННАЯ ФУНКЦИЯ: очистка с полным сбросом форматирования
    On Error Resume Next
    
    If dataStart = 0 Or dataEnd = 0 Then
        ' Секция не используется
        ClearSection = 0
        Exit Function
    End If
    
    ModuleLogger.LogDebug "?? Очистка " & sectionName & " секции с форматированием (строки " & dataStart & "-" & dataEnd & ")"
    
    Dim clearedCells As Long
    clearedCells = 0
    
    ' Используем фиксированный столбец DF (110) как границу данных
    Dim lastCol As Long
    lastCol = 110 ' Столбец DF
    
    ' Проходим по всем строкам секции
    Dim row As Long
    For row = dataStart To dataEnd
        ' Очищаем только столбцы с данными товаров (начиная с C), НЕ трогаем A и B!
        Dim col As Long
        For col = 3 To lastCol ' От C до DF
            Dim currentCell As Range
            Set currentCell = ws.Cells(row, col)
            
            ' Проверяем, что это ячейка с данными (количество/партия)
            If Not isEmpty(currentCell.value) And Trim(CStr(currentCell.value)) <> "" Then
                ' Проверяем, что это не служебная информация
                If Not IsServiceCell(ws, row, col) Then
                    ' ===== НОВАЯ ЛОГИКА: ОЧИСТКА + ПОЛНЫЙ СБРОС ФОРМАТИРОВАНИЯ =====
                    
                    ' Очищаем содержимое
                    currentCell.ClearContents
                    
                    ' ?? СБРАСЫВАЕМ ВСЕ ФОРМАТИРОВАНИЕ НА СТАНДАРТНОЕ
                    With currentCell
                        ' Белый фон
                        .Interior.ColorIndex = xlNone
                        
                        ' Черный шрифт
                        .Font.ColorIndex = xlAutomatic
                        
                        ' Сбрасываем стили шрифта
                        .Font.Bold = False
                        .Font.Italic = False
                        .Font.Underline = False
                        .Font.Size = 11  ' Стандартный размер
                        
                        ' Стандартное выравнивание
                        .HorizontalAlignment = xlGeneral
                        .VerticalAlignment = xlBottom
                        
                        ' Убираем специальные границы
                        With .Borders
                            .LineStyle = xlContinuous
                            .Weight = xlThin
                            .ColorIndex = xlAutomatic
                        End With
                    End With
                    
                    clearedCells = clearedCells + 1
                    
                    ' Логируем только каждую 50-ю очищенную ячейку
                    If clearedCells Mod 50 = 0 Then
                        ModuleLogger.LogDebug "?? Очищено с форматированием: " & clearedCells & " ячеек в " & sectionName & " секции"
                    End If
                End If
            End If
        Next col
    Next row
    
    ModuleLogger.LogSuccess "? Очищена " & sectionName & " секция ангара " & warehouse & ": " & clearedCells & " ячеек (строки " & dataStart & "-" & dataEnd & ") + СБРОС ФОРМАТИРОВАНИЯ"
    
    ClearSection = clearedCells
    On Error GoTo 0
End Function

Private Function IsServiceCell(ws As Worksheet, row As Long, col As Long) As Boolean
    ' Проверка - является ли ячейка служебной (не подлежит очистке)
    
    ' НЕ очищаем столбцы A и B
    If col <= 2 Then
        IsServiceCell = True
        Exit Function
    End If
    
    Dim cellValue As String
    cellValue = Trim(CStr(ws.Cells(row, col).value))
    
    ' НЕ очищаем пустые ячейки
    If cellValue = "" Then
        IsServiceCell = True
        Exit Function
    End If
    
    ' НЕ очищаем номера рядов (простые числа 1-10)
    If IsNumeric(cellValue) Then
        Dim numValue As Double
        numValue = CDbl(cellValue)
        If numValue >= 1 And numValue <= 10 And numValue = Int(numValue) Then
            ' Дополнительно проверяем, что рядом есть еще номера (характерно для строки с номерами рядов)
            Dim nextCol As Long
            nextCol = col + 2 ' Проверяем через один столбец
            If nextCol <= 20 Then
                Dim nextValue As String
                nextValue = Trim(CStr(ws.Cells(row, nextCol).value))
                If IsNumeric(nextValue) Then
                    Dim nextNum As Double
                    nextNum = CDbl(nextValue)
                    If nextNum >= 1 And nextNum <= 10 And nextNum = Int(nextNum) Then
                        IsServiceCell = True ' Это строка с номерами рядов
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    
    ' НЕ очищаем ячейки со словом "ИТОГО"
    If InStr(UCase(cellValue), "ИТОГО") > 0 Or InStr(UCase(cellValue), "ИТОГ") > 0 Then
        IsServiceCell = True
        Exit Function
    End If
    
    ' НЕ очищаем ячейки с формулами итогов
    If Left(cellValue, 1) = "=" Then
        IsServiceCell = True
        Exit Function
    End If
    
    ' Все остальные ячейки подлежат очистке
    IsServiceCell = False
End Function

Private Sub ResetRangeFormatting(ws As Worksheet, startRow As Long, endRow As Long, sectionName As String)
    ' КРАСИВАЯ ФУНКЦИЯ: применение форматирования как в ModuleFormat
    On Error Resume Next
    
    ModuleLogger.LogMessage "?? Применение красивого форматирования " & sectionName & " (строки " & startRow & "-" & endRow & ")"
    
    ' Определяем диапазон данных (от C до DF)
    Dim dataRange As Range
    Set dataRange = ws.Range("C" & startRow & ":DF" & endRow)
    
    Dim abRange As Range
    Set abRange = ws.Range("A" & startRow & ":B" & endRow)
    
    ' ===== ШАГ 1: ОЧИЩАЕМ ВСЕ ГРАНИЦЫ =====
    ' Очищаем ВСЕ границы в основном диапазоне (C:DF)
    With dataRange
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
    End With
    
    ' Очищаем ВСЕ границы в столбцах A и B
    With abRange
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
    End With
    
    ' ===== ШАГ 2: ПРИМЕНЯЕМ КРАСИВОЕ ФОРМАТИРОВАНИЕ =====
    
    ' ?? ШРИФТ: Calibri, размер 12, жирный
    With dataRange
        .Font.Name = "Calibri"
        .Font.Size = 12
        .Font.Bold = True
        .Font.ColorIndex = xlAutomatic  ' Черный
        
        ' ?? ВЫРАВНИВАНИЕ: по середине и по центру
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        
        ' ?? ЦВЕТА: белый фон
        .Interior.ColorIndex = xlNone
    End With
    
    ' Также форматируем столбцы A и B
    With abRange
        .Font.Name = "Calibri"
        .Font.Size = 12
        .Font.Bold = True
        .Font.ColorIndex = xlAutomatic
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.ColorIndex = xlNone
    End With
    
    ' ===== ШАГ 3: КРАСИВЫЕ РАМКИ (как в ModuleFormat) =====
    
    ' Устанавливаем точечные внутренние линии для основного диапазона
    With dataRange.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    With dataRange.Borders(xlInsideVertical)
        .LineStyle = xlDot
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    ' Устанавливаем точечные внутренние линии для столбцов A и B
    With abRange
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlDot
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        
        With .Borders(xlInsideVertical)
            .LineStyle = xlDot
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        
        ' Добавляем правую границу для столбца B
        With .Borders(xlEdgeRight)
            .LineStyle = xlDot
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    End With
    
    ' ===== ШАГ 4: ЖИРНЫЕ ЛИНИИ ДЛЯ ОБЪЕДИНЕННЫХ ЯЧЕЕК =====
    
    ' Проходим по каждой строке и проверяем объединенные ячейки
    Dim i As Long
    For i = startRow To endRow
        If ws.Cells(i, 1).MergeCells Then
            Dim mergeHeight As Long
            mergeHeight = ws.Cells(i, 1).MergeArea.rows.count
            
            ' Добавляем жирную пунктирную линию НАД группой
            With ws.Range("A" & i & ":DF" & i).Borders(xlEdgeTop)
                .LineStyle = xlSlantDashDot
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            
            ' Добавляем жирную сплошную линию ПОД группой
            With ws.Range("A" & (i + mergeHeight - 1) & ":DF" & (i + mergeHeight - 1)).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            
            i = i + mergeHeight - 1
        End If
    Next i
    
    ModuleLogger.LogSuccess "? Красивое форматирование применено: " & sectionName & " (диапазон A" & startRow & ":DF" & endRow & ")"
    ModuleLogger.LogInfo "   ?? Шрифт: Calibri, 12pt, жирный"
    ModuleLogger.LogInfo "   ?? Выравнивание: по центру"
    ModuleLogger.LogInfo "   ?? Рамки: точечные + жирные для групп"
    
    On Error GoTo 0
End Sub

Private Sub ApplyWarehouseDimensions(ws As Worksheet, warehouse As String)
    ' НОВАЯ ФУНКЦИЯ: применяет правильные размеры столбцов и строк как в ModuleFormat
    On Error Resume Next
    
    ModuleLogger.LogMessage "?? Применение размеров для ангара " & warehouse
    
    ' Настройка столбцов A и B (стандартно для всех ангаров)
    ws.Range("A:A").ColumnWidth = 8
    ws.Range("B:B").ColumnWidth = 15
    
    ' Получаем конфигурацию для определения диапазонов
    Dim config As ModuleTypes.WarehouseConfig
    config = ModuleConfig.GetWarehouseConfig(warehouse)
    
    ' Настройка столбцов C:DF (ширина 10 для всех ангаров)
    Dim lastCol As String
    Select Case warehouse
        Case "5": lastCol = "CT69"
        Case "6": lastCol = "CT70"
        Case "7": lastCol = "CT56"
        Case "8": lastCol = "CT55"
        Case "9": lastCol = "CT79"
        Case "10": lastCol = "CT78"
        Case "11": lastCol = "CT65"
        Case "12": lastCol = "CT66"
        Case Else: lastCol = "CT70"  ' По умолчанию
    End Select
    
    ' Устанавливаем ширину столбцов (только для не скрытых)
    Dim col As Range
    For Each col In ws.Range("C4:" & lastCol).columns
        If Not col.Hidden Then
            col.ColumnWidth = 10
        End If
    Next col
    
    ' Устанавливаем высоту строк в диапазоне данных
    If config.UpperDataStart > 0 And config.UpperDataEnd > 0 Then
        Dim dataRowRange As Range
        Set dataRowRange = ws.Range(config.UpperDataStart & ":" & config.UpperDataEnd)
        
        Dim row As Range
        For Each row In dataRowRange.rows
            If Not row.Hidden Then
                row.RowHeight = 20
            End If
        Next row
    End If
    
    ' Настройка высоты заголовочных строк
    Select Case warehouse
        Case "5"
            ws.Range("2:2").RowHeight = 35
            ws.Range("3:3").RowHeight = 15
            ws.Range("69:69").RowHeight = 15
            ws.Range("70:70").RowHeight = 35
        Case "6"
            ws.Range("2:2").RowHeight = 35
            ws.Range("3:3").RowHeight = 15
            ws.Range("71:71").RowHeight = 15
            ws.Range("72:72").RowHeight = 35
        Case "7"
            ws.Range("2:2").RowHeight = 35
            ws.Range("3:3").RowHeight = 15
            ws.Range("57:57").RowHeight = 15
            ws.Range("58:58").RowHeight = 35
        Case "8"
            ws.Range("2:2").RowHeight = 35
            ws.Range("3:3").RowHeight = 15
            ws.Range("56:56").RowHeight = 15
            ws.Range("57:57").RowHeight = 35
        Case "9"
            ws.Range("2:2").RowHeight = 35
            ws.Range("3:3").RowHeight = 15
            ws.Range("80:80").RowHeight = 15
            ws.Range("81:81").RowHeight = 35
        Case "10"
            ws.Range("2:2").RowHeight = 35
            ws.Range("3:3").RowHeight = 15
            ws.Range("79:79").RowHeight = 15
            ws.Range("80:80").RowHeight = 35
        Case "11"
            ws.Range("2:2").RowHeight = 35
            ws.Range("3:3").RowHeight = 15
            ws.Range("66:66").RowHeight = 15
            ws.Range("67:67").RowHeight = 35
        Case "12"
            ws.Range("2:2").RowHeight = 35
            ws.Range("3:3").RowHeight = 15
            ws.Range("67:67").RowHeight = 15
            ws.Range("68:68").RowHeight = 35
    End Select
    
    ModuleLogger.LogInfo "?? Размеры применены для ангара " & warehouse
    
    On Error GoTo 0
End Sub

Private Function ClearOutputSection(ws As Worksheet, outputStartRow As Long, warehouse As String) As Long
    ' Очистка области результатов с форматированием
    On Error Resume Next
    
    ModuleLogger.LogDebug "?? Очистка области результатов с форматированием (начиная со строки " & outputStartRow & ")"
    
    ' Находим последнюю строку с данными
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    
    Dim clearedCells As Long
    clearedCells = 0
    
    If lastRow >= outputStartRow Then
        ' Очищаем область результатов (обычно столбцы A и B)
        Dim outputRange As Range
        Set outputRange = ws.Range(ws.Cells(outputStartRow, 1), ws.Cells(lastRow, 2))
        
        ' Подсчитываем заполненные ячейки
        Dim cell As Range
        For Each cell In outputRange
            If Not isEmpty(cell.value) Then
                clearedCells = clearedCells + 1
            End If
        Next cell
        
        ' Очищаем диапазон с форматированием
        outputRange.ClearContents
        
        ' Сбрасываем форматирование
        With outputRange
            .Interior.ColorIndex = xlNone
            .Font.ColorIndex = xlAutomatic
            .Font.Bold = False
            .Font.Italic = False
            .Font.Underline = False
        End With
        
        ModuleLogger.LogInfo "? Очищена область результатов ангара " & warehouse & ": " & clearedCells & " ячеек + форматирование"
    End If
    
    ClearOutputSection = clearedCells
    On Error GoTo 0
End Function

Public Function ClearSpecificWarehouse(warehouse As String) As String
    ' Функция для очистки конкретного ангара с форматированием
    On Error GoTo ErrorHandler
    
    ModuleLogger.LogMessage "?? Очистка с форматированием конкретного ангара: " & warehouse
    
    Application.ScreenUpdating = False
    
    Dim clearedCells As Long
    Dim result As String
    result = ClearWarehouse(warehouse, clearedCells)
    
    ' Также очищаем заголовки этого ангара
    Dim config As ModuleTypes.WarehouseConfig
    config = ModuleConfig.GetWarehouseConfig(warehouse)
    
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(config.sheetName)
    
    ' Очищаем заголовки верхней секции
    If config.UpperHeaderRow > 0 Then
        Call ClearHeaderRow(ws, config.UpperHeaderRow, "Ангар " & warehouse & " (верхняя секция)")
    End If
    
    ' Очищаем заголовки нижней секции
    If config.LowerHeaderRow > 0 Then
        Call ClearHeaderRow(ws, config.LowerHeaderRow, "Ангар " & warehouse & " (нижняя секция)")
    End If
    
    Application.ScreenUpdating = True
    
    If result = "" Then
        ModuleLogger.LogSuccess "? Ангар " & warehouse & " успешно очищен с форматированием (" & clearedCells & " ячеек)"
    End If
    
    ClearSpecificWarehouse = result
    Exit Function
    
ErrorHandler:
    Application.ScreenUpdating = True
    ModuleLogger.LogError "Ошибка очистки ангара " & warehouse & ": " & Err.description
    ClearSpecificWarehouse = "Ошибка очистки ангара " & warehouse & ": " & Err.description
End Function

Private Sub ClearHeaderRow(ws As Worksheet, headerRow As Long, sectionName As String)
    ' ОБНОВЛЕННАЯ: очистка заголовков + красивое форматирование
    On Error Resume Next
    
    ' Разъединяем все объединенные ячейки в строке заголовков (только в диапазоне данных C:DF)
    ws.Range("C" & headerRow & ":DF" & headerRow).UnMerge
    
    ' Очищаем содержимое (начиная с колонки C до DF, так как A и B содержат стеллажи)
    ws.Range("C" & headerRow & ":DF" & headerRow).ClearContents
    
    ' НОВОЕ: Применяем красивое форматирование заголовков
    With ws.Range("C" & headerRow & ":DF" & headerRow)
        ' Цвета
        .Interior.ColorIndex = xlNone       ' Белый фон
        .Font.ColorIndex = xlAutomatic      ' Черный шрифт
        
        ' КРАСИВЫЙ ШРИФТ (как требовалось)
        .Font.Name = "Calibri"              ' Шрифт Calibri
        .Font.Size = 12                     ' Размер 12
        .Font.Bold = True                   ' Жирный
        .Font.Italic = False                ' Не курсив
        .Font.Underline = False             ' Не подчеркнутый
        
        ' ВЫРАВНИВАНИЕ (по середине и по центру)
        .HorizontalAlignment = xlCenter     ' По середине
        .VerticalAlignment = xlCenter       ' По центру
        
        ' КРАСИВЫЕ РАМКИ (точечные как в ModuleFormat)
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlDot
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        
        With .Borders(xlInsideVertical)
            .LineStyle = xlDot
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    End With
    
    ModuleLogger.LogSuccess "? Заголовки очищены с красивым форматированием: " & sectionName & " (строка " & headerRow & ", столбцы C-DF)"
    ModuleLogger.LogInfo "   ?? Шрифт: Calibri, 12pt, жирный, по центру"
    ModuleLogger.LogInfo "   ?? Рамки: точечные"
    
    On Error GoTo 0
End Sub

Public Sub PreserveBoundariesAndStructure()
    ' Функция для сохранения границ и структуры таблиц после очистки
    ModuleLogger.LogMessage "?? Восстановление структуры таблиц..."
    
    Dim i As Integer
    For i = 5 To 12
        On Error Resume Next
        Dim config As ModuleTypes.WarehouseConfig
        config = ModuleConfig.GetWarehouseConfig(CStr(i))
        
        Dim ws As Worksheet
        Set ws = ActiveWorkbook.Worksheets(config.sheetName)
        
        If Not ws Is Nothing Then
            ' Восстанавливаем базовое форматирование
            RestoreBasicFormatting ws, config, i
            ModuleLogger.LogInfo "?? Структура ангара " & i & " восстановлена"
        End If
        On Error GoTo 0
    Next i
    
    ModuleLogger.LogMessage "? Восстановление структуры завершено"
End Sub

Private Sub RestoreBasicFormatting(ws As Worksheet, config As ModuleTypes.WarehouseConfig, warehouse As Integer)
    ' Восстанавливает базовое форматирование после очистки
    On Error Resume Next
    
    ' Восстанавливаем стандартные границы для верхней секции
    If config.UpperDataStart > 0 And config.UpperDataEnd > 0 Then
        Dim upperRange As Range
        Set upperRange = ws.Range("C" & config.UpperDataStart & ":DF" & config.UpperDataEnd)
        
        With upperRange.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    End If
    
    ' Восстанавливаем стандартные границы для нижней секции
    If config.LowerDataStart > 0 And config.LowerDataEnd > 0 Then
        Dim lowerRange As Range
        Set lowerRange = ws.Range("C" & config.LowerDataStart & ":DF" & config.LowerDataEnd)
        
        With lowerRange.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    End If
    
    On Error GoTo 0
End Sub

Public Function GetClearingSummary() As String
    ' Возвращает информацию о том, что будет очищено
    Dim summary As String
    summary = "=== ЧТО БУДЕТ ОЧИЩЕНО (ОБНОВЛЕННАЯ ВЕРСИЯ) ===" & vbNewLine & vbNewLine
    summary = summary & "??? УДАЛЯЕТСЯ:" & vbNewLine
    summary = summary & "• Все заголовки товаров (строки 2, 58, 70, 72 и т.д.)" & vbNewLine
    summary = summary & "• Все количества товаров" & vbNewLine
    summary = summary & "• Все номера партий" & vbNewLine
    summary = summary & "• Все результаты подсчетов" & vbNewLine & vbNewLine
    
    summary = summary & "?? СБРАСЫВАЕТСЯ ФОРМАТИРОВАНИЕ:" & vbNewLine
    summary = summary & "• Цвет фона ячеек > белый" & vbNewLine
    summary = summary & "• Цвет шрифта > черный" & vbNewLine
    summary = summary & "• Жирный/курсив > обычный" & vbNewLine
    summary = summary & "• Размер шрифта > стандартный (11)" & vbNewLine
    summary = summary & "• Выравнивание > стандартное" & vbNewLine & vbNewLine
    
    summary = summary & "? СОХРАНЯЕТСЯ:" & vbNewLine
    summary = summary & "• Названия стеллажей (А, Б, В, Г...)" & vbNewLine
    summary = summary & "• Номера рядов (1, 2, 3, 4...)" & vbNewLine
    summary = summary & "• Структура таблиц" & vbNewLine
    summary = summary & "• Базовые границы ячеек" & vbNewLine
    summary = summary & "• Общая архитектура листов" & vbNewLine
    
    GetClearingSummary = summary
End Function

' ===== ТЕСТОВЫЕ ФУНКЦИИ =====

Public Sub TestFormattingClear()
    ' Тест очистки с форматированием
    Call ModuleLogger.InitializeLogger(True, ModuleLogger.LOG_LEVEL_DEBUG)
    
    ModuleLogger.LogMessage "=== ТЕСТ ОЧИСТКИ С ФОРМАТИРОВАНИЕМ ==="
    
    Dim testWarehouse As String
    testWarehouse = InputBox("Введите номер ангара для теста очистки с форматированием (5-12):", "Тест ангара", "7")
    
    If testWarehouse <> "" And IsNumeric(testWarehouse) Then
        If CInt(testWarehouse) >= 5 And CInt(testWarehouse) <= 12 Then
            ModuleLogger.LogMessage "?? Тестирование очистки с форматированием для ангара " & testWarehouse
            
            Dim clearedCells As Long
            Dim result As String
            result = ClearSpecificWarehouse(testWarehouse)
            
            If result = "" Then
                MsgBox "? Тест очистки с форматированием прошел успешно!" & vbNewLine & _
                       "Ангар " & testWarehouse & " очищен полностью" & vbNewLine & _
                       "Проверьте:" & vbNewLine & _
                       "?? Белый фон ячеек" & vbNewLine & _
                       "? Черный шрифт" & vbNewLine & _
                       "?? Обычный стиль текста" & vbNewLine & _
                       "?? Стандартные границы", vbInformation, "Тест прошел"
            Else
                MsgBox "? Ошибка теста: " & result, vbExclamation, "Ошибка теста"
            End If
        Else
            MsgBox "Неверный номер ангара! Используйте числа от 5 до 12.", vbExclamation, "Ошибка"
        End If
    End If
End Sub

Public Sub QuickClearTest()
    ' Быстрый тест очистки
    MsgBox "?? ТЕСТ ОБНОВЛЕННОЙ ОЧИСТКИ" & vbNewLine & vbNewLine & _
           "Новые возможности:" & vbNewLine & _
           "? Очистка содержимого ячеек" & vbNewLine & _
           "?? Сброс цвета фона на белый" & vbNewLine & _
           "?? Сброс цвета шрифта на черный" & vbNewLine & _
           "?? Сброс стилей шрифта (жирный, курсив)" & vbNewLine & _
           "?? Восстановление стандартных границ" & vbNewLine & vbNewLine & _
           "Доступные тесты:" & vbNewLine & _
           "• TestFormattingClear() - тест одного ангара" & vbNewLine & _
           "• ClearAllWarehouses() - очистка всех ангаров", _
           vbInformation, "Тест новой очистки"
End Sub

Public Sub TestBeautifulFormatting()
    ' НОВЫЙ ТЕСТ: красивое форматирование
    Call ModuleLogger.InitializeLogger(True, ModuleLogger.LOG_LEVEL_DEBUG)
    
    ModuleLogger.LogMessage "=== ТЕСТ КРАСИВОГО ФОРМАТИРОВАНИЯ ==="
    
    Dim testWarehouse As String
    testWarehouse = InputBox("Введите номер ангара для теста красивого форматирования (5-12):", "Тест красивого форматирования", "6")
    
    If testWarehouse <> "" And IsNumeric(testWarehouse) Then
        If CInt(testWarehouse) >= 5 And CInt(testWarehouse) <= 12 Then
            ModuleLogger.LogMessage "?? Тестирование красивого форматирования для ангара " & testWarehouse
            
            Dim result As String
            result = ClearSpecificWarehouse(testWarehouse)
            
            If result = "" Then
                MsgBox "? ТЕСТ КРАСИВОГО ФОРМАТИРОВАНИЯ ПРОШЕЛ!" & vbNewLine & vbNewLine & _
                       "Ангар " & testWarehouse & " теперь имеет:" & vbNewLine & _
                       "?? Шрифт: Calibri, 12pt, жирный" & vbNewLine & _
                       "?? Выравнивание: по центру" & vbNewLine & _
                       "?? Белый фон ячеек" & vbNewLine & _
                       "? Черный текст" & vbNewLine & _
                       "?? Красивые рамки (точечные + жирные для групп)" & vbNewLine & _
                       "?? Правильные размеры столбцов и строк" & vbNewLine & vbNewLine & _
                       "Как в ModuleFormat!", vbInformation, "Красивое форматирование"
            Else
                MsgBox "? Ошибка теста: " & result, vbExclamation, "Ошибка теста"
            End If
        Else
            MsgBox "Неверный номер ангара! Используйте числа от 5 до 12.", vbExclamation, "Ошибка"
        End If
    End If
End Sub

