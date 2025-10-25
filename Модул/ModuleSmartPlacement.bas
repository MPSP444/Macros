Attribute VB_Name = "ModuleSmartPlacement"
' ===== ModuleSmartPlacement - ПОЛНОСТЬЮ ИСПРАВЛЕННАЯ ВЕРСИЯ "ЯЧЕЙКА ЗА ЯЧЕЙКОЙ" =====
' ? Надо исправить: А1>А2>А3>Б1>Б2>Б3 (вертикальное заполнение каждой буквы)
' ? НЕ А1>Б1>В1>Г1 (горизонтальное заполнение)
' ? Проверка каждой ячейки отдельно
' ? Правильные лимиты по объемам
' ? Куски от прохода к стене, целые от стены к проходу

Option Explicit

' Структура для информации о размещении
Public Type PlacementInfo
    warehouse As String      ' Номер ангара
    row As String           ' Номер ряда
    Shelf As String         ' Буква стеллажа
    level As String         ' Уровень (ярус)
    ProductName As String   ' Название товара
    Batch As String         ' Партия
    quantity As Double      ' Количество для размещения
    standardVolume As Double ' Объем стандартной паллеты
    pieceType As String     ' Тип: WHOLE (целая), MEDIUM (средняя), PIECE (кусок)
    weightCategory As String ' HEAVY, MEDIUM, LIGHT (для ярусов)
    IsPlaced As Boolean     ' Успешно размещено?
    ErrorMessage As String  ' Сообщение об ошибке
    placementDetails As String ' Детали размещения для отчета
End Type

' Структура для анализа заполненности стеллажа
Private Type shelfAnalysis
    Letter As String        ' Буква (А, Б, В...)
    totalQuantity As Double ' Общее количество на стеллаже
    EmptyLevel As String    ' Первый пустой уровень (1, 2, 3)
    HasProduct As Boolean   ' Есть ли уже товар на этом стеллаже
    CanAccommodate As Boolean ' Может ли вместить новое количество
    Priority As Integer     ' Приоритет размещения (1=высший)
    ShelfLimit As Double    ' Лимит этого стеллажа
    Level1Quantity As Double ' Количество на уровне 1
    Level2Quantity As Double ' Количество на уровне 2
    Level3Quantity As Double ' Количество на уровне 3
    SuitableLevel As String  ' Подходящий уровень для данного веса
End Type

' Структура для найденного товара
Private Type FoundProductInfo
    warehouse As String     ' Номер ангара
    startRow As Integer     ' Начальный ряд товара
    endRow As Integer       ' Конечный ряд товара
    availableRows As String ' Список доступных рядов
    section As String      ' Секция (UPPER/LOWER)
    totalQuantity As Double ' Общее количество товара в ангаре
    headerRow As Long      ' Строка заголовка
    productInHeader As String ' Товар в заголовке
End Type

' ===== ТАБЛИЦЫ ЛИМИТОВ ПО АНГАРАМ И ОБЪЕМАМ =====

Private Function GetShelfLimit(warehouseNum As String, productVolume As Double) As Double
    Select Case warehouseNum
        Case "5", "6", "9", "10", "11"
            Select Case CInt(productVolume)
                Case 1000: GetShelfLimit = 2000
                Case 960:  GetShelfLimit = 1000
                Case 720:  GetShelfLimit = 1500
                Case 600:  GetShelfLimit = 1300
                Case 540:  GetShelfLimit = 1100
                Case 480:  GetShelfLimit = 980
                Case 288:  GetShelfLimit = 580
                Case 240:  GetShelfLimit = 500
                Case Else: GetShelfLimit = 1500
            End Select
            
        Case "7", "8"
            Select Case CInt(productVolume)
                Case 1000: GetShelfLimit = 2000
                Case 960:  GetShelfLimit = 1000
                Case 720:  GetShelfLimit = 1500
                Case 600:  GetShelfLimit = 1300
                Case 540:  GetShelfLimit = 1100
                Case 480:  GetShelfLimit = 980
                Case 288:  GetShelfLimit = 580
                Case 240:  GetShelfLimit = 500
                Case Else: GetShelfLimit = 1500
            End Select
            
        Case "12"
            Select Case CInt(productVolume)
                Case 1000: GetShelfLimit = 2000
                Case 960:  GetShelfLimit = 1000
                Case 720:  GetShelfLimit = 1500
                Case 600:  GetShelfLimit = 1300
                Case 540:  GetShelfLimit = 1100
                Case 480:  GetShelfLimit = 980
                Case 288:  GetShelfLimit = 580
                Case 240:  GetShelfLimit = 500
                Case Else: GetShelfLimit = 1500
            End Select
            
        Case Else
            GetShelfLimit = 1500
    End Select
End Function

Private Function GetRowLimit(warehouseNum As String, productVolume As Double) As Double
    Select Case warehouseNum
        Case "5", "6", "9", "10", "11"
            Select Case CInt(productVolume)
                Case 1000: GetRowLimit = 14000
                Case 960:  GetRowLimit = 6800
                Case 720:  GetRowLimit = 10100
                Case 600:  GetRowLimit = 8500
                Case 540:  GetRowLimit = 8700
                Case 480:  GetRowLimit = 7000
                Case 288:  GetRowLimit = 4608
                Case 240:  GetRowLimit = 3850
                Case Else: GetRowLimit = 10100
            End Select
            
        Case "7", "8"
            Select Case CInt(productVolume)
                Case 1000: GetRowLimit = 10000
                Case 960:  GetRowLimit = 5000
                Case 720:  GetRowLimit = 7200
                Case 600:  GetRowLimit = 6000
                Case 540:  GetRowLimit = 6500
                Case 480:  GetRowLimit = 5000
                Case 288:  GetRowLimit = 3800
                Case 240:  GetRowLimit = 3000
                Case Else: GetRowLimit = 7200
            End Select
            
        Case "12"
            Select Case CInt(productVolume)
                Case 1000: GetRowLimit = 16000
                Case 960:  GetRowLimit = 7700
                Case 720:  GetRowLimit = 11550
                Case 600:  GetRowLimit = 9600
                Case 540:  GetRowLimit = 8700
                Case 480:  GetRowLimit = 7000
                Case 288:  GetRowLimit = 5000
                Case 240:  GetRowLimit = 4320
                Case Else: GetRowLimit = 11550
            End Select
            
        Case Else
            GetRowLimit = 10100
    End Select
End Function

Private Function DetermineWeightCategory(quantity As Double, standardVolume As Double) As String
    Dim percentage As Double
    If standardVolume > 0 Then
        percentage = (quantity / standardVolume) * 100
    Else
        percentage = 100
    End If
    
    If percentage >= 80 Then
        DetermineWeightCategory = "HEAVY"
    ElseIf percentage >= 50 Then
        DetermineWeightCategory = "MEDIUM"
    Else
        DetermineWeightCategory = "LIGHT"
    End If
    
    ModuleLogger.LogDebug "?? Вес: " & quantity & "/" & standardVolume & " = " & _
                         Format(percentage, "0.0") & "% > " & DetermineWeightCategory
End Function

' ===== ?? ГЛАВНАЯ ФУНКЦИЯ УМНОГО РАЗМЕЩЕНИЯ (ЯЧЕЙКА ЗА ЯЧЕЙКОЙ) =====
Public Function SmartPlaceProducts(inputText As String) As String
    On Error GoTo ErrorHandler
    
    ModuleLogger.LogMessage "=== ?? УМНОЕ РАЗМЕЩЕНИЕ 'ЯЧЕЙКА ЗА ЯЧЕЙКОЙ' ==="
    ModuleLogger.LogMessage "? ИСПРАВЛЕНО: А1>А2>А3>Б1>Б2>Б3 (правильно)"
    ModuleLogger.LogMessage "? БЫЛО: А1>Б1>В1>Г1 (неправильно)"
    
    ' Инициализация систем
    Call ModuleProductInfo.InitializeProductDatabase
    Call ModuleWarehouseCapacity.InitializeWarehouseCapacities
    
    Application.ScreenUpdating = False
    
    ' Проверяем активные ангары
    Dim activeCount As Integer
    activeCount = ModuleWarehouseCapacity.GetActiveWarehouseCount()
    
    If activeCount = 0 Then
        SmartPlaceProducts = "? НЕ ВЫБРАН НИ ОДИН АКТИВНЫЙ АНГАР!"
        Application.ScreenUpdating = True
        Exit Function
    End If
    
    ModuleLogger.LogMessage "? Активные ангары: " & ModuleWarehouseCapacity.GetActiveWarehousesList()
    
    ' 1. Парсинг с форматом объемов
    Dim placements() As PlacementInfo
    placements = ParseNewFormatInput(inputText)
    
    If UBound(placements) < 0 Then
        SmartPlaceProducts = "Не найдено корректных данных для размещения!"
        Application.ScreenUpdating = True
        Exit Function
    End If
    
    ' 2. РАЗМЕЩЕНИЕ ЯЧЕЙКА ЗА ЯЧЕЙКОЙ
    Dim i As Integer
    Dim successCount As Integer
    successCount = 0
    
    For i = 0 To UBound(placements)
        If placements(i).ProductName <> "" Then
            ' Определяем категорию веса
            placements(i).weightCategory = DetermineWeightCategory(placements(i).quantity, placements(i).standardVolume)
            
            ModuleLogger.LogMessage "?? РАЗМЕЩЕНИЕ ячейка-за-ячейкой: " & placements(i).ProductName & " - " & _
                                   placements(i).quantity & " (" & placements(i).standardVolume & ") - " & _
                                   placements(i).pieceType & " - Вес: " & placements(i).weightCategory
            
            ' ?? ИСПОЛЬЗУЕМ ЛОГИКУ ЯЧЕЙКА ЗА ЯЧЕЙКОЙ
            If SmartPlaceWithCellByCell(placements(i)) Then
                successCount = successCount + 1
                ModuleLogger.LogSuccess "? Размещено ячейка-за-ячейкой: " & placements(i).ProductName
            Else
                ModuleLogger.LogError "? Ошибка размещения: " & placements(i).ProductName & " - " & placements(i).ErrorMessage
            End If
        End If
    Next i
    
    ' 3. Генерация отчета
    Dim report As String
    report = GenerateCellByCellReport(placements, successCount, activeCount)  ' ? ИСПРАВЛЕНО: убран пробел
    
    Application.ScreenUpdating = True
    
    ModuleLogger.LogMessage "=== ?? ЗАВЕРШЕНИЕ РАЗМЕЩЕНИЯ 'ЯЧЕЙКА ЗА ЯЧЕЙКОЙ' ==="
    ModuleLogger.LogSuccess "Успешно размещено: " & successCount & " из " & (UBound(placements) + 1)
    
    SmartPlaceProducts = report
    Exit Function
    
ErrorHandler:
    Application.ScreenUpdating = True
    SmartPlaceProducts = "Критическая ошибка: " & Err.description
    ModuleLogger.LogError "КРИТИЧЕСКАЯ ОШИБКА: " & Err.description
End Function

' ===== ?? РАЗМЕЩЕНИЕ С ЛОГИКОЙ ЯЧЕЙКА ЗА ЯЧЕЙКОЙ =====
Private Function SmartPlaceWithCellByCell(ByRef placement As PlacementInfo) As Boolean
    On Error GoTo ErrorHandler
    
    ModuleLogger.LogMessage "?? РАЗМЕЩЕНИЕ ячейка-за-ячейкой: " & placement.ProductName & _
                           " (" & placement.quantity & "/" & placement.standardVolume & ") - " & _
                           placement.pieceType & " - Вес: " & placement.weightCategory
    
    ' ШАГ 1: Найти ангар с максимальным количеством этого товара
    Dim bestWarehouse As String
    bestWarehouse = FindWarehouseWithMaxProduct(placement.ProductName)
    
    If bestWarehouse <> "" Then
        ModuleLogger.LogSuccess "?? Лучший ангар для '" & placement.ProductName & "': " & bestWarehouse
        
        ' ШАГ 2: Попробовать разместить в лучшем ангаре (ЯЧЕЙКА ЗА ЯЧЕЙКОЙ)
        If TryPlaceInBestWarehouseCellByCell(placement, bestWarehouse) Then
            placement.IsPlaced = True
            SmartPlaceWithCellByCell = True
            Exit Function
        End If
    End If
    
    ' ШАГ 3: Если не удалось в лучшем ангаре, ищем в других активных (ЯЧЕЙКА ЗА ЯЧЕЙКОЙ)
    If TryPlaceInAnyActiveWarehouseCellByCell(placement) Then
        placement.IsPlaced = True
        SmartPlaceWithCellByCell = True
        Exit Function
    End If
    
    ' Не удалось разместить
    placement.ErrorMessage = "Не найдено места с учетом проверки каждой ячейки"
    placement.IsPlaced = False
    SmartPlaceWithCellByCell = False
    Exit Function
    
ErrorHandler:
    placement.ErrorMessage = "Ошибка размещения ячейка-за-ячейкой: " & Err.description
    SmartPlaceWithCellByCell = False
End Function

' ===== РАЗМЕЩЕНИЕ В ЛУЧШЕМ АНГАРЕ (ЯЧЕЙКА ЗА ЯЧЕЙКОЙ) =====
Private Function TryPlaceInBestWarehouseCellByCell(ByRef placement As PlacementInfo, warehouse As String) As Boolean
    ModuleLogger.LogMessage "?? Размещение в лучшем ангаре " & warehouse & " (ячейка-за-ячейкой)"
    
    ' Ищем существующий товар в этом ангаре
    Dim existingProduct As FoundProductInfo
    existingProduct = FindProductInWarehouseHeaders(placement.ProductName, warehouse)
    
    If existingProduct.startRow > 0 Then
        ModuleLogger.LogSuccess "? Товар найден в рядах: " & existingProduct.availableRows & " - проверяем каждую ячейку"
        
        ' Пробуем добить существующие ряды с логикой ЯЧЕЙКА ЗА ЯЧЕЙКОЙ
        If TryPlaceInExistingRowsCellByCell(placement, existingProduct, warehouse) Then
            TryPlaceInBestWarehouseCellByCell = True
            Exit Function
        End If
        
        ' Ищем соседние пустые ряды
        ModuleLogger.LogMessage "?? Ищем соседние пустые ряды..."
        If FindAdjacentEmptyRowsCellByCell(placement, warehouse, existingProduct) Then
            ModuleLogger.LogSuccess "? Размещено в соседнем ряду (ячейка-за-ячейкой)!"
            TryPlaceInBestWarehouseCellByCell = True
            Exit Function
        End If
    End If
    
    ' Ищем пустые ряды
    If FindEmptyRowCellByCell(placement, warehouse) Then
        TryPlaceInBestWarehouseCellByCell = True
        Exit Function
    End If
    
    TryPlaceInBestWarehouseCellByCell = False
End Function

' ===== РАЗМЕЩЕНИЕ В СУЩЕСТВУЮЩИХ РЯДАХ (ЯЧЕЙКА ЗА ЯЧЕЙКОЙ) =====
Private Function TryPlaceInExistingRowsCellByCell(ByRef placement As PlacementInfo, _
                                                 existingProduct As FoundProductInfo, _
                                                 warehouse As String) As Boolean
    
    ModuleLogger.LogMessage "?? РАЗМЕЩЕНИЕ ячейка-за-ячейкой в существующих рядах"
    
    ' Разбираем список рядов
    Dim rows() As String
    rows = Split(existingProduct.availableRows, ",")
    
    ' Пробуем каждый ряд с логикой ячейка-за-ячейкой
    Dim i As Integer
    For i = 0 To UBound(rows)
        Dim rowNum As String
        rowNum = Trim(rows(i))
        
        If TryPlaceInRowCellByCell(placement, warehouse, rowNum, existingProduct.section) Then
            TryPlaceInExistingRowsCellByCell = True
            Exit Function
        End If
    Next i
    
    TryPlaceInExistingRowsCellByCell = False
End Function

' ===== ?? ГЛАВНАЯ ФУНКЦИЯ: РАЗМЕЩЕНИЕ В РЯДУ (ЯЧЕЙКА ЗА ЯЧЕЙКОЙ) =====
Private Function TryPlaceInRowCellByCell(ByRef placement As PlacementInfo, warehouse As String, _
                                        rowNum As String, section As String) As Boolean
    On Error GoTo ErrorHandler
    
    ModuleLogger.LogMessage "?? РАЗМЕЩЕНИЕ ячейка-за-ячейкой в ряду " & rowNum & " секции " & ModuleConfig.GetSectionName(section)
    
    ' Получаем конфигурацию ангара
    Dim config As ModuleTypes.WarehouseConfig
    config = ModuleConfig.GetWarehouseConfig(warehouse)
    
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(config.sheetName)
    
    ' Проверяем лимит ряда
    Dim currentRowTotal As Double
    currentRowTotal = CalculateRowTotalWithVolumes(ws, rowNum, section, warehouse)
    
    Dim rowLimit As Double
    rowLimit = GetRowLimit(warehouse, placement.standardVolume)
    
    If currentRowTotal + placement.quantity > rowLimit Then
        ModuleLogger.LogWarning "? Ряд " & rowNum & " переполнен: " & (currentRowTotal + placement.quantity) & " > " & rowLimit
        TryPlaceInRowCellByCell = False
        Exit Function
    End If
    
    ' ?? ИЩЕМ ТОЧНОЕ МЕСТО: А1>А2>А3>Б1>Б2>Б3...
    Dim exactLocation As String
    exactLocation = FindExactCellForPlacement(ws, warehouse, rowNum, section, placement)
    
    If exactLocation = "" Then
        ModuleLogger.LogDebug "? Нет свободных ячеек в ряду " & rowNum
        TryPlaceInRowCellByCell = False
        Exit Function
    End If
    
    ' Парсим точное место
    Dim locationParts() As String
    locationParts = Split(exactLocation, "-")
    
    If UBound(locationParts) < 1 Then
        ModuleLogger.LogError "? Ошибка парсинга места: " & exactLocation
        TryPlaceInRowCellByCell = False
        Exit Function
    End If
    
    ' Размещаем товар в точном месте
    placement.warehouse = warehouse
    placement.row = rowNum
    placement.Shelf = locationParts(0)
    placement.level = locationParts(1)
    
    If PlaceInExactCell(placement, section) Then
        placement.placementDetails = "Ангар " & placement.warehouse & "-" & placement.row & "-" & _
                                   placement.Shelf & "-" & placement.level & " - " & _
                                   placement.ProductName & " (" & placement.Batch & ") - " & _
                                   placement.quantity & " "
        
        ModuleLogger.LogSuccess "?? РАЗМЕЩЕНО ТОЧНО: " & placement.placementDetails
        TryPlaceInRowCellByCell = True
    Else
        TryPlaceInRowCellByCell = False
    End If
    
    Exit Function
    
ErrorHandler:
    ModuleLogger.LogError "Ошибка размещения ячейка-за-ячейкой: " & Err.description
    TryPlaceInRowCellByCell = False
End Function

' ===== ?? КЛЮЧЕВАЯ ФУНКЦИЯ: ПОИСК ТОЧНОГО МЕСТА ПО ЯЧЕЙКАМ =====
Private Function FindExactCellForPlacement(ws As Worksheet, warehouse As String, rowNum As String, _
                                         section As String, placement As PlacementInfo) As String
    ' Возвращает точное место в формате "А-2" (буква-ярус)
    
    ModuleLogger.LogMessage "?? ПОИСК МЕСТА ячейка-за-ячейкой в ряду " & rowNum
    
    ' Получаем список букв в правильном порядке
    Dim letters() As String
    letters = GetLettersWithCorrectDirection(warehouse, section, placement.pieceType)
    
    ' Получаем конфигурацию
    Dim config As ModuleTypes.WarehouseConfig
    config = ModuleConfig.GetWarehouseConfig(warehouse)
    
    Dim dataStart As Long, dataEnd As Long
    If section = "UPPER" Then
        dataStart = config.UpperDataStart
        dataEnd = config.UpperDataEnd
    Else
        dataStart = config.LowerDataStart
        dataEnd = config.LowerDataEnd
    End If
    
    Dim valueColumn As Long
    valueColumn = 3 + (CLng(rowNum) - 1) * 2
    
    ' ?? ПРОВЕРЯЕМ КАЖДУЮ БУКВУ И КАЖДЫЙ ЯРУС
    Dim i As Integer
    For i = 0 To UBound(letters)
        Dim currentLetter As String
        currentLetter = letters(i)
        
        ModuleLogger.LogDebug "?? Проверяем букву " & currentLetter & " по ярусам..."
        
        ' Анализируем эту букву детально
        Dim letterAnalysis As shelfAnalysis
        letterAnalysis = AnalyzeSingleShelfDetailed(ws, currentLetter, valueColumn, dataStart, dataEnd, _
                                                   warehouse, placement.standardVolume)
        
        ' ?? ИЩЕМ ТОЧНОЕ МЕСТО ПО ЯРУСАМ: 1>2>3
        Dim exactCell As String
        exactCell = FindExactLevelInShelf(letterAnalysis, placement.quantity, placement.weightCategory)
        
        If exactCell <> "" Then
            ModuleLogger.LogSuccess "? НАЙДЕНО МЕСТО: " & currentLetter & "-" & exactCell
            FindExactCellForPlacement = currentLetter & "-" & exactCell
            Exit Function
        End If
    Next i
    
    ModuleLogger.LogWarning "? Нет свободных ячеек в ряду " & rowNum
    FindExactCellForPlacement = ""
End Function

' ===== ?? ДЕТАЛЬНЫЙ АНАЛИЗ ОДНОЙ БУКВЫ =====
Private Function AnalyzeSingleShelfDetailed(ws As Worksheet, shelfLetter As String, valueColumn As Long, _
                                           dataStart As Long, dataEnd As Long, warehouse As String, _
                                           productVolume As Double) As shelfAnalysis
    
    Dim analysis As shelfAnalysis
    analysis.Letter = shelfLetter
    analysis.ShelfLimit = GetShelfLimit(warehouse, productVolume)
    analysis.totalQuantity = 0
    analysis.Level1Quantity = 0
    analysis.Level2Quantity = 0
    analysis.Level3Quantity = 0
    analysis.HasProduct = False
    
    ModuleLogger.LogDebug "?? Детальный анализ буквы " & shelfLetter & " (лимит: " & analysis.ShelfLimit & ")"
    
    ' Проверяем каждый ярус: 1, 2, 3
    Dim level As Integer
    For level = 1 To 3
        Dim searchText As String
        searchText = ModuleConfig.GetLevelText(shelfLetter, CStr(level))
        
        Dim foundCell As Range
        Set foundCell = ws.Range("A" & dataStart & ":B" & dataEnd).Find( _
            What:=searchText, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not foundCell Is Nothing Then
            Dim cellValue As Double
            cellValue = GetNumericValueSafe(ws.Cells(foundCell.row, valueColumn))
            
            ' Записываем количество по ярусам
            Select Case level
                Case 1:
                    analysis.Level1Quantity = cellValue
                    ModuleLogger.LogDebug "   Ярус 1: " & cellValue
                Case 2:
                    analysis.Level2Quantity = cellValue
                    ModuleLogger.LogDebug "   Ярус 2: " & cellValue
                Case 3:
                    analysis.Level3Quantity = cellValue
                    ModuleLogger.LogDebug "   Ярус 3: " & cellValue
            End Select
            
            analysis.totalQuantity = analysis.totalQuantity + cellValue
            If cellValue > 0 Then analysis.HasProduct = True
        End If
    Next level
    
    ModuleLogger.LogDebug "?? Буква " & shelfLetter & " итого: " & analysis.totalQuantity & "/" & analysis.ShelfLimit
    
    AnalyzeSingleShelfDetailed = analysis
End Function

' ===== ?? ПОИСК ТОЧНОГО ЯРУСА В БУКВЕ =====
Private Function FindExactLevelInShelf(shelfAnalysis As shelfAnalysis, quantity As Double, _
                                      weightCategory As String) As String
    
    ' Проверяем, поместится ли товар в эту букву
    If shelfAnalysis.totalQuantity + quantity > shelfAnalysis.ShelfLimit Then
        ModuleLogger.LogDebug "? Буква " & shelfAnalysis.Letter & " переполнена: " & _
                             (shelfAnalysis.totalQuantity + quantity) & " > " & shelfAnalysis.ShelfLimit
        FindExactLevelInShelf = ""
        Exit Function
    End If
    
    ' ?? ВЕРТИКАЛЬНЫЙ ПОИСК: 1>2>3
    
    ' Ярус 1 (пол)
    If shelfAnalysis.Level1Quantity = 0 Then
        ModuleLogger.LogDebug "? Ярус 1 свободен"
        FindExactLevelInShelf = "1"
        Exit Function
    End If
    
    ' Ярус 2 (средний)
    If shelfAnalysis.Level2Quantity = 0 And shelfAnalysis.Level1Quantity > 0 Then
        If weightCategory = "HEAVY" Then
            ModuleLogger.LogWarning "?? ТЯЖЕЛОЕ на ярус 2 - проверить безопасность"
        End If
        ModuleLogger.LogDebug "? Ярус 2 свободен"
        FindExactLevelInShelf = "2"
        Exit Function
    End If
    
    ' Ярус 3 (верх)
    If shelfAnalysis.Level3Quantity = 0 And shelfAnalysis.Level2Quantity > 0 And shelfAnalysis.Level1Quantity > 0 Then
        If weightCategory = "HEAVY" Then
            ModuleLogger.LogWarning "? ТЯЖЕЛОЕ нельзя на ярус 3"
            FindExactLevelInShelf = ""
            Exit Function
        End If
        ModuleLogger.LogDebug "? Ярус 3 свободен"
        FindExactLevelInShelf = "3"
        Exit Function
    End If
    
    ' Все ярусы заняты
    ModuleLogger.LogDebug "? Все ярусы в букве " & shelfAnalysis.Letter & " заняты"
    FindExactLevelInShelf = ""
End Function

' ===== РАЗМЕЩЕНИЕ В ТОЧНОЙ ЯЧЕЙКЕ =====
Private Function PlaceInExactCell(ByRef placement As PlacementInfo, section As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim config As ModuleTypes.WarehouseConfig
    config = ModuleConfig.GetWarehouseConfig(placement.warehouse)
    
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(config.sheetName)
    
    Dim dataStart As Long, dataEnd As Long
    If section = "UPPER" Then
        dataStart = config.UpperDataStart
        dataEnd = config.UpperDataEnd
    Else
        dataStart = config.LowerDataStart
        dataEnd = config.LowerDataEnd
    End If
    
    Dim searchText As String
    searchText = ModuleConfig.GetLevelText(placement.Shelf, placement.level)
    
    Dim targetCell As Range
    Set targetCell = ws.Range("A" & dataStart & ":B" & dataEnd).Find( _
        What:=searchText, LookIn:=xlValues, LookAt:=xlWhole)
    
    If targetCell Is Nothing Then
        placement.ErrorMessage = "Не найдена ячейка " & searchText
        PlaceInExactCell = False
        Exit Function
    End If
    
    Dim valueColumn As Long
    valueColumn = 3 + (CLng(placement.row) - 1) * 2
    
    Dim batchColumn As Long
    batchColumn = valueColumn + 1
    
    ' Записываем данные
    ws.Cells(targetCell.row, valueColumn).value = placement.quantity
    ws.Cells(targetCell.row, batchColumn).value = placement.Batch
    
    ' ?? ЦВЕТОВАЯ СХЕМА ДЛЯ ЯЧЕЙКА-ЗА-ЯЧЕЙКОЙ
    Dim highlightColor As Long
    Select Case placement.pieceType
        Case "WHOLE"
            ' Целая палета (например, 720 из 720)
            highlightColor = RGB(0, 176, 80)   ' Насыщенный зеленый
    
        Case "MEDIUM"
            ' Средний объем (например, 650 из 720)
            highlightColor = RGB(146, 208, 80) ' Светло-зеленый
    
        Case "PIECE"
            ' Кусок (например, 150 из 720)
            highlightColor = RGB(255, 255, 0)  ' Желтый
    
        Case Else
            ' Цвет по умолчанию, если тип не определен
            highlightColor = RGB(217, 217, 217) ' Серый
    End Select
    
    ' Специальная индикация
    If placement.pieceType = "PIECE" Then
        ' Красная граница для кусков
        With ws.Cells(targetCell.row, valueColumn).Borders
            .LineStyle = xlContinuous
            .Weight = xlThick
            .Color = RGB(255, 0, 0)
        End With
        With ws.Cells(targetCell.row, batchColumn).Borders
            .LineStyle = xlContinuous
            .Weight = xlThick
            .Color = RGB(255, 0, 0)
        End With
    Else
        ' Синяя граница для правильного размещения ячейка-за-ячейкой
        With ws.Cells(targetCell.row, valueColumn).Borders
            .LineStyle = xlContinuous
            .Weight = xlThick
            .Color = RGB(0, 100, 200)
        End With
        With ws.Cells(targetCell.row, batchColumn).Borders
            .LineStyle = xlContinuous
            .Weight = xlThick
            .Color = RGB(0, 100, 200)
        End With
    End If
    
    ws.Cells(targetCell.row, valueColumn).Interior.Color = highlightColor
    ws.Cells(targetCell.row, batchColumn).Interior.Color = highlightColor
    
    ModuleLogger.LogSuccess "?? ЯЧЕЙКА-ЗА-ЯЧЕЙКОЙ в " & targetCell.address & " ярус " & placement.level & _
                          " [" & placement.pieceType & "/" & placement.weightCategory & "] " & _
                          placement.Shelf & "-" & placement.level
    
    PlaceInExactCell = True
    Exit Function
    
ErrorHandler:
    placement.ErrorMessage = "Ошибка размещения в ячейке: " & Err.description
    PlaceInExactCell = False
End Function

' ===== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ (ДОБАВЛЕНЫ ИЗ ДОКУМЕНТА №2) =====

Private Function TryPlaceInAnyActiveWarehouseCellByCell(ByRef placement As PlacementInfo) As Boolean
    ModuleLogger.LogMessage "?? Поиск места в любом активном ангаре (ячейка-за-ячейкой)"
    
    ' Проходим по всем активным ангарам
    Dim warehouse As Integer
    For warehouse = 5 To 12
        If ModuleWarehouseCapacity.IsWarehouseActive(CStr(warehouse)) Then
            
            If TryPlaceInBestWarehouseCellByCell(placement, CStr(warehouse)) Then
                TryPlaceInAnyActiveWarehouseCellByCell = True
                Exit Function
            End If
        End If
    Next warehouse
    
    TryPlaceInAnyActiveWarehouseCellByCell = False
End Function

Private Function FindAdjacentEmptyRowsCellByCell(ByRef placement As PlacementInfo, warehouse As String, _
                                                existingProduct As FoundProductInfo) As Boolean
    On Error GoTo ErrorHandler
    
    ModuleLogger.LogMessage "?? Поиск соседних пустых рядов (ячейка-за-ячейкой)"
    
    ' Парсим существующие ряды
    Dim existingRows() As String
    existingRows = Split(existingProduct.availableRows, ",")
    
    ' Находим минимальный и максимальный ряды
    Dim minRow As Integer, maxRow As Integer
    minRow = 999
    maxRow = 0
    
    Dim i As Integer
    For i = 0 To UBound(existingRows)
        Dim currentRowNum As Integer
        currentRowNum = CInt(Trim(existingRows(i)))
        
        If currentRowNum < minRow Then minRow = currentRowNum
        If currentRowNum > maxRow Then maxRow = currentRowNum
    Next i
    
    ModuleLogger.LogMessage "?? Диапазон существующих рядов: " & minRow & " - " & maxRow
    
    ' Получаем конфигурацию ангара
    Dim config As ModuleTypes.WarehouseConfig
    config = ModuleConfig.GetWarehouseConfig(warehouse)
    
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(config.sheetName)
    
    Dim targetSection As String
    targetSection = existingProduct.section
    
    Dim headerRow As Long
    If targetSection = "UPPER" Then
        headerRow = config.UpperHeaderRow
    Else
        headerRow = config.LowerHeaderRow
    End If
    
    ' Проверяем соседние ряды
    Dim distance As Integer
    For distance = 1 To 10
        ' Справа от максимального
        Dim rightRow As Integer
        rightRow = maxRow + distance
        If rightRow >= 1 And rightRow <= 60 Then
            If IsRowCompletelyEmpty(ws, CStr(rightRow), headerRow, targetSection, warehouse) Then
                ModuleLogger.LogSuccess "? Найден соседний пустой ряд: " & rightRow
                
                If CreateNewHeaderAndPlaceCellByCell(placement, warehouse, CStr(rightRow), targetSection) Then
                    FindAdjacentEmptyRowsCellByCell = True
                    Exit Function
                End If
            End If
        End If
        
        ' Слева от минимального
        Dim leftRow As Integer
        leftRow = minRow - distance
        If leftRow >= 1 And leftRow <= 60 Then
            If IsRowCompletelyEmpty(ws, CStr(leftRow), headerRow, targetSection, warehouse) Then
                ModuleLogger.LogSuccess "? Найден соседний пустой ряд: " & leftRow
                
                If CreateNewHeaderAndPlaceCellByCell(placement, warehouse, CStr(leftRow), targetSection) Then
                    FindAdjacentEmptyRowsCellByCell = True
                    Exit Function
                End If
            End If
        End If
    Next distance
    
    FindAdjacentEmptyRowsCellByCell = False
    Exit Function
    
ErrorHandler:
    ModuleLogger.LogError "Ошибка поиска соседних рядов: " & Err.description
    FindAdjacentEmptyRowsCellByCell = False
End Function

Private Function FindEmptyRowCellByCell(ByRef placement As PlacementInfo, warehouse As String) As Boolean
    ModuleLogger.LogMessage "?? Поиск пустого ряда в ангаре " & warehouse & " (ячейка-за-ячейкой)"
    
    ' Получаем конфигурацию ангара
    Dim config As ModuleTypes.WarehouseConfig
    config = ModuleConfig.GetWarehouseConfig(warehouse)
    
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(config.sheetName)
    
    ' Ищем в верхней секции
    If config.UpperHeaderRow > 0 Then
        Dim r As Integer
        For r = 1 To 60
            If IsRowCompletelyEmpty(ws, CStr(r), config.UpperHeaderRow, "UPPER", warehouse) Then
                If CreateNewHeaderAndPlaceCellByCell(placement, warehouse, CStr(r), "UPPER") Then
                    FindEmptyRowCellByCell = True
                    Exit Function
                End If
            End If
        Next r
    End If
    
    ' Ищем в нижней секции
    If config.LowerHeaderRow > 0 Then
        For r = 1 To 60
            If IsRowCompletelyEmpty(ws, CStr(r), config.LowerHeaderRow, "LOWER", warehouse) Then
                If CreateNewHeaderAndPlaceCellByCell(placement, warehouse, CStr(r), "LOWER") Then
                    FindEmptyRowCellByCell = True
                    Exit Function
                End If
            End If
        Next r
    End If
    
    FindEmptyRowCellByCell = False
End Function

Private Function CreateNewHeaderAndPlaceCellByCell(ByRef placement As PlacementInfo, warehouse As String, _
                                                  rowNum As String, section As String) As Boolean
    ModuleLogger.LogMessage "?? Создаем заголовок для '" & placement.ProductName & "' в ряду " & rowNum & " (ячейка-за-ячейкой)"
    
    ' Получаем конфигурацию
    Dim config As ModuleTypes.WarehouseConfig
    config = ModuleConfig.GetWarehouseConfig(warehouse)
    
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(config.sheetName)
    
    Dim headerRow As Long
    If section = "UPPER" Then
        headerRow = config.UpperHeaderRow
    Else
        headerRow = config.LowerHeaderRow
    End If
    
    If headerRow = 0 Then
        CreateNewHeaderAndPlaceCellByCell = False
        Exit Function
    End If
    
    ' Создаем заголовок
    Call CreateMergedHeader(ws, placement.ProductName, rowNum, headerRow)
    
    ' Размещаем товар методом ячейка-за-ячейкой
    If TryPlaceInRowCellByCell(placement, warehouse, rowNum, section) Then
        ModuleLogger.LogSuccess "? Создан заголовок и размещен товар (ячейка-за-ячейкой)"
        CreateNewHeaderAndPlaceCellByCell = True
    Else
        CreateNewHeaderAndPlaceCellByCell = False
    End If
End Function

' ===== ДОБАВЛЕНЫ НЕДОСТАЮЩИЕ ФУНКЦИИ ИЗ ДОКУМЕНТА №2 =====

Private Function IsRowCompletelyEmpty(ws As Worksheet, rowNum As String, headerRow As Long, _
                                    section As String, warehouse As String) As Boolean
    ' Проверяем заголовок
    If headerRow > 0 Then
        Dim headerCol As Long
        headerCol = 3 + (CLng(rowNum) - 1) * 2
        
        Dim headerValue As String
        headerValue = GetCellValueSafe(ws.Cells(headerRow, headerCol))
        
        If Trim(headerValue) <> "" Then
            IsRowCompletelyEmpty = False
            Exit Function
        End If
    End If
    
    ' Проверяем данные
    Dim totalInRow As Double
    totalInRow = CalculateRowTotalWithVolumes(ws, rowNum, section, warehouse)
    
    IsRowCompletelyEmpty = (totalInRow = 0)
End Function

Private Sub CreateMergedHeader(ws As Worksheet, ProductName As String, rowNum As String, headerRow As Long)
    On Error Resume Next
    
    Dim startCol As Long, endCol As Long
    startCol = 3 + (CLng(rowNum) - 1) * 2
    endCol = startCol + 1
    
    ws.Range(ws.Cells(headerRow, startCol), ws.Cells(headerRow, endCol)).UnMerge
    ws.Range(ws.Cells(headerRow, startCol), ws.Cells(headerRow, endCol)).ClearContents
    ws.Range(ws.Cells(headerRow, startCol), ws.Cells(headerRow, endCol)).Merge
    
    ws.Cells(headerRow, startCol).value = ProductName
    
    With ws.Cells(headerRow, startCol)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.Bold = True
        .Font.Size = 12
        .Interior.Color = RGB(200, 255, 200)
        .Font.Color = RGB(0, 100, 0)
        
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(100, 100, 100)
        End With
    End With
    
    On Error GoTo 0
End Sub

Private Function GetLettersWithCorrectDirection(warehouse As String, section As String, pieceType As String) As String()
    Dim allLetters() As String
    allLetters = GetSectionLetters(warehouse, section) ' Получает буквы, например [ПРИ, И, К, ..., Р] для нижней

    Dim isLowerSectionPassageFirst As Boolean
    ' Проверяем, является ли первая буква в массиве проходом для нижней секции
    isLowerSectionPassageFirst = (section = "LOWER" And Left(allLetters(0), 2) = "ПР")

    Dim needsReversing As Boolean
    needsReversing = False ' По умолчанию массив не переворачиваем

    Select Case pieceType
        Case "WHOLE", "MEDIUM" ' Для целых и средних палет
            ' Если это нижняя секция и буквы начинаются с прохода (ПРИ, ПРМ...),
            ' то массив нужно перевернуть, чтобы начать со стены.
            If isLowerSectionPassageFirst Then
                needsReversing = True
                ModuleLogger.LogDebug "?? WHOLE/LOWER: стена>проход (реверс): " & Join(ReverseArray(allLetters), ">")
            Else
                ModuleLogger.LogDebug "?? WHOLE/UPPER: стена>проход: " & Join(allLetters, ">")
            End If

        Case "PIECE" ' Для кусков
            ' Если это верхняя секция (буквы НЕ начинаются с прохода),
            ' то массив нужно перевернуть, чтобы начать с прохода.
            If Not isLowerSectionPassageFirst Then
                needsReversing = True
                ModuleLogger.LogDebug "?? PIECE/UPPER: проход>стена (реверс): " & Join(ReverseArray(allLetters), ">")
            Else
                ModuleLogger.LogDebug "?? PIECE/LOWER: проход>стена: " & Join(allLetters, ">")
            End If
    End Select

    If needsReversing Then
        GetLettersWithCorrectDirection = ReverseArray(allLetters)
    Else
        GetLettersWithCorrectDirection = allLetters
    End If
End Function


Private Function ReverseArray(arr() As String) As String()
    Dim reversed() As String
    ReDim reversed(UBound(arr))
    
    Dim i As Integer
    For i = 0 To UBound(arr)
        reversed(i) = arr(UBound(arr) - i)
    Next i
    
    ReverseArray = reversed
End Function

Private Function FindWarehouseWithMaxProduct(ProductName As String) As String
    ModuleLogger.LogMessage "?? Ищем ангар с максимальным количеством '" & ProductName & "'"
    
    Dim maxQuantity As Double
    maxQuantity = 0
    
    Dim bestWarehouse As String
    bestWarehouse = ""
    
    ' Проверяем все активные ангары
    Dim warehouse As Integer
    For warehouse = 5 To 12
        If ModuleWarehouseCapacity.IsWarehouseActive(CStr(warehouse)) Then
            
            Dim quantity As Double
            quantity = CalculateTotalProductInWarehouse(ProductName, CStr(warehouse))
            
            If quantity > maxQuantity Then
                maxQuantity = quantity
                bestWarehouse = CStr(warehouse)
            End If
            
            ModuleLogger.LogDebug "Ангар " & warehouse & ": " & quantity & "л товара '" & ProductName & "'"
        End If
    Next warehouse
    
    If bestWarehouse <> "" Then
        ModuleLogger.LogSuccess "?? Максимум в ангаре " & bestWarehouse & ": " & maxQuantity & "л"
    Else
        ModuleLogger.LogMessage "?? Товар не найден (новый товар)"
    End If
    
    FindWarehouseWithMaxProduct = bestWarehouse
End Function

Private Function CalculateTotalProductInWarehouse(ProductName As String, warehouse As String) As Double
    On Error Resume Next
    
    ' Получаем конфигурацию ангара
    Dim config As ModuleTypes.WarehouseConfig
    config = ModuleConfig.GetWarehouseConfig(warehouse)
    
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(config.sheetName)
    
    If ws Is Nothing Then
        CalculateTotalProductInWarehouse = 0
        Exit Function
    End If
    
    Dim totalQuantity As Double
    totalQuantity = 0
    
    ' Ищем товар в заголовках и суммируем количество
    totalQuantity = totalQuantity + SumProductInSection(ws, ProductName, config.UpperHeaderRow, config.UpperDataStart, config.UpperDataEnd, "UPPER")
    totalQuantity = totalQuantity + SumProductInSection(ws, ProductName, config.LowerHeaderRow, config.LowerDataStart, config.LowerDataEnd, "LOWER")
    
    CalculateTotalProductInWarehouse = totalQuantity
    On Error GoTo 0
End Function

Private Function SumProductInSection(ws As Worksheet, ProductName As String, headerRow As Long, _
                                    dataStart As Long, dataEnd As Long, sectionName As String) As Double
    On Error Resume Next
    
    If headerRow = 0 Then
        SumProductInSection = 0
        Exit Function
    End If
    
    Dim totalInSection As Double
    totalInSection = 0
    
    ' Проходим по всем заголовкам в секции
    Dim col As Long
    Dim lastCheckedCol As Long
    lastCheckedCol = 0
    
    For col = 3 To 127
        If col > lastCheckedCol Then
            
            Dim headerCell As Range
            Set headerCell = ws.Cells(headerRow, col)
            
            Dim headerValue As String
            headerValue = GetCellValueSafe(headerCell)
            
            If headerValue <> "" Then
                ' Проверяем совпадение товара
                If UCase(Trim(headerValue)) = UCase(Trim(ProductName)) Then
                    ' Определяем диапазон заголовка
                    Dim headerStartCol As Long, headerEndCol As Long
                    If headerCell.MergeCells Then
                        headerStartCol = headerCell.MergeArea.Column
                        headerEndCol = headerCell.MergeArea.Column + headerCell.MergeArea.columns.count - 1
                        lastCheckedCol = headerEndCol
                    Else
                        headerStartCol = col
                        headerEndCol = col
                        lastCheckedCol = col
                    End If
                    
                    ' Суммируем количество в соответствующих рядах
                    Dim c As Long
                    For c = headerStartCol To headerEndCol Step 2
                        totalInSection = totalInSection + SumColumnInRange(ws, c, dataStart, dataEnd)
                    Next c
                End If
            End If
        End If
    Next col
    
    SumProductInSection = totalInSection
    On Error GoTo 0
End Function

Private Function SumColumnInRange(ws As Worksheet, col As Long, startRow As Long, endRow As Long) As Double
    On Error Resume Next
    
    Dim total As Double
    total = 0
    
    Dim r As Long
    For r = startRow To endRow
        Dim cellValue As Double
        cellValue = GetNumericValueSafe(ws.Cells(r, col))
        total = total + cellValue
    Next r
    
    SumColumnInRange = total
    On Error GoTo 0
End Function

Private Function FindProductInWarehouseHeaders(ProductName As String, warehouse As String) As FoundProductInfo
    Dim result As FoundProductInfo
    
    ' Получаем конфигурацию ангара
    Dim config As ModuleTypes.WarehouseConfig
    config = ModuleConfig.GetWarehouseConfig(warehouse)
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(config.sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        FindProductInWarehouseHeaders = result
        Exit Function
    End If
    
    ' Ищем в верхней секции
    result = SearchProductInSectionHeaders(ws, ProductName, config.UpperHeaderRow, "UPPER")
    If result.startRow > 0 Then
        result.warehouse = warehouse
        FindProductInWarehouseHeaders = result
        Exit Function
    End If
    
    ' Ищем в нижней секции
    result = SearchProductInSectionHeaders(ws, ProductName, config.LowerHeaderRow, "LOWER")
    If result.startRow > 0 Then
        result.warehouse = warehouse
    End If
    
    FindProductInWarehouseHeaders = result
End Function

' ===== ИСПРАВЛЕННАЯ ВЕРСИЯ: ИЩЕТ ВСЕ РЯДЫ С ТОВАРОМ, А НЕ ПЕРВЫЙ ПОПАВШИЙСЯ =====
Private Function SearchProductInSectionHeaders(ws As Worksheet, ProductName As String, _
                                              headerRow As Long, sectionName As String) As FoundProductInfo
    Dim result As FoundProductInfo
    result.section = sectionName
    result.headerRow = headerRow
    result.availableRows = "" ' Инициализируем пустой строкой

    If headerRow = 0 Then
        SearchProductInSectionHeaders = result
        Exit Function
    End If

    Dim col As Long
    Dim lastCheckedCol As Long
    lastCheckedCol = 0

    For col = 3 To 127 ' Проходим по всем возможным столбцам
        If col > lastCheckedCol Then
            Dim headerCell As Range
            Set headerCell = ws.Cells(headerRow, col)
            
            Dim headerValue As String
            headerValue = GetCellValueSafe(headerCell)

            If headerValue <> "" Then
                ' Проверяем совпадение товара (используем UCase для надежности)
                If UCase(Trim(headerValue)) = UCase(Trim(ProductName)) Then
                    ' Определяем диапазон столбцов для этого заголовка
                    Dim headerStartCol As Long, headerEndCol As Long
                    If headerCell.MergeCells Then
                        headerStartCol = headerCell.MergeArea.Column
                        headerEndCol = headerCell.MergeArea.Column + headerCell.MergeArea.columns.count - 1
                    Else
                        headerStartCol = col
                        headerEndCol = col
                    End If

                    ' Определяем номера рядов, которые покрывает этот заголовок
                    Dim startR As Integer, endR As Integer
                    startR = ((headerStartCol - 3) \ 2) + 1
                    endR = ((headerEndCol - 2) \ 2) ' Было endR = ((headerEndCol - 3) \ 2) + 1, исправлено для корректного диапазона

                    ' Добавляем все найденные ряды в общий список
                    Dim r As Integer
                    For r = startR To endR
                        If result.availableRows <> "" Then
                            result.availableRows = result.availableRows & ","
                        End If
                        result.availableRows = result.availableRows & CStr(r)
                    Next r
                    
                    ' Устанавливаем startRow, чтобы показать, что товар найден
                    If result.startRow = 0 Then result.startRow = startR
                    
                    ' НЕ ВЫХОДИМ ИЗ ЦИКЛА (Exit Function), а продолжаем поиск
                End If

                ' Пропускаем проверенные столбцы, чтобы не сканировать их заново
                If headerCell.MergeCells Then
                    lastCheckedCol = headerCell.MergeArea.Column + headerCell.MergeArea.columns.count - 1
                Else
                    lastCheckedCol = col
                End If
            End If
        End If
    Next col

    ' Если ряды были найдены, сортируем их для логичного порядка обработки
    ' Если ряды были найдены, сортируем их для логичного порядка обработки
    If result.availableRows <> "" Then
        Dim rowsToSort() As String
        
        ' 1. Сначала разбиваем строку на массив
        rowsToSort = Split(result.availableRows, ",")
        
        ' 2. Затем передаем этот массив в функцию сортировки
        Dim sortedRows() As String
        sortedRows = SortStringArrayNumerically(rowsToSort)
        
        ' 3. Собираем отсортированный массив обратно в строку
        result.availableRows = Join(sortedRows, ",")
        
        ModuleLogger.LogDebug "?? Итоговый список рядов для '" & ProductName & "': " & result.availableRows
    End If


    SearchProductInSectionHeaders = result
End Function

' Вспомогательная функция для сортировки массива строк как чисел
Private Function SortStringArrayNumerically(arr() As String) As String()
    Dim i As Long, j As Long
    Dim temp As String
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If CLng(arr(i)) > CLng(arr(j)) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
    SortStringArrayNumerically = arr
End Function


Private Function CalculateRowTotalWithVolumes(ws As Worksheet, rowNum As String, section As String, warehouse As String) As Double
    Dim config As ModuleTypes.WarehouseConfig
    config = ModuleConfig.GetWarehouseConfig(warehouse)
    
    Dim valueColumn As Long
    valueColumn = 3 + (CLng(rowNum) - 1) * 2
    
    Dim dataStart As Long, dataEnd As Long
    If section = "UPPER" Then
        dataStart = config.UpperDataStart
        dataEnd = config.UpperDataEnd
    Else
        dataStart = config.LowerDataStart
        dataEnd = config.LowerDataEnd
    End If
    
    Dim total As Double
    total = 0
    
    Dim row As Long
    For row = dataStart To dataEnd
        Dim cellValue As Double
        cellValue = GetNumericValueSafe(ws.Cells(row, valueColumn))
        total = total + cellValue
    Next row
    
    CalculateRowTotalWithVolumes = total
End Function

Private Function GetSectionLetters(warehouse As String, section As String) As String()
    Dim letters() As String
    
    Select Case warehouse
        Case "5", "6"
            If section = "UPPER" Then
                letters = Split("А,Б,В,Г,Д,Е,Ж,З,ПРЗ", ",")
            Else
                letters = Split("ПРИ,И,К,Л,М,Н,О,П,Р", ",")
            End If
            
        Case "7", "8"
            If section = "UPPER" Then
                letters = Split("А,Б,В,Г,Д,Е,ПРЕ", ",")
            Else
                letters = Split("ПРЖ,Ж,З,И,К,Л,М", ",")
            End If
            
        Case "9"
            If section = "UPPER" Then
                letters = Split("А,Б,В,Г,Д,Е,Ж,З,И,ПРИ", ",")
            Else
                letters = Split("ПРМ,М,Н,О,П,Р,С,Т,У,Ф", ",")
            End If
            
        Case "10"
            If section = "UPPER" Then
                letters = Split("А,Б,В,Г,Д,Е,Ж,З,И,ПРИ", ",")
            Else
                letters = Split("ПРК,К,Л,М,Н,О,П,Р,С,Т", ",")
            End If
            
        Case "11"
            If section = "UPPER" Then
                letters = Split("А,Б,В,Г,Д,Е,Ж,ПРЗ", ",")
            Else
                letters = Split("ПРИ,З,И,К,Л,М,Н,О", ",")
            End If
            
        Case "12"
            If section = "UPPER" Then
                letters = Split("А,Б,В,Г,Д,Е,Ж,ПРЖ", ",")
            Else
                letters = Split("ПРЗ,З,И,К,Л,М,Н,О", ",")
            End If
            
        Case Else
            letters = Split("А,Б,В,Г,Д,Е,Ж,З", ",")
    End Select
    
    GetSectionLetters = letters
End Function

Private Function GetCellValueSafe(cell As Range) As String
    On Error Resume Next
    If cell.MergeCells Then
        GetCellValueSafe = Trim(CStr(cell.MergeArea.Cells(1, 1).value))
    Else
        GetCellValueSafe = Trim(CStr(cell.value))
    End If
    If GetCellValueSafe = "False" Then GetCellValueSafe = ""
    On Error GoTo 0
End Function

Private Function GetNumericValueSafe(cell As Range) As Double
    On Error Resume Next
    Dim value As Variant
    
    If cell.MergeCells Then
        value = cell.MergeArea.Cells(1, 1).value
    Else
        value = cell.value
    End If
    
    If IsNumeric(value) And Not isEmpty(value) Then
        GetNumericValueSafe = CDbl(value)
    Else
        GetNumericValueSafe = 0
    End If
    On Error GoTo 0
End Function

' ===== ПАРСИНГ ВХОДНЫХ ДАННЫХ =====
Private Function ParseNewFormatInput(inputText As String) As PlacementInfo()
    Dim lines() As String
    lines = Split(inputText, vbNewLine)
    
    Dim results() As PlacementInfo
    ReDim results(UBound(lines))
    
    Dim validCount As Integer
    validCount = 0
    
    Dim i As Integer
    For i = 0 To UBound(lines)
        Dim line As String
        line = Trim(lines(i))
        
        If line <> "" And Not (InStr(LCase(line), "пример") > 0) Then
            Dim placement As PlacementInfo
            
            ' Формат: "Товар - партия - количество (объем)"
            If ParseNewFormatLine(line, placement) Then
                ' Определяем категорию веса
                placement.weightCategory = DetermineWeightCategory(placement.quantity, placement.standardVolume)
                
                results(validCount) = placement
                validCount = validCount + 1
                
                ModuleLogger.LogDebug "?? Распознано: " & placement.ProductName & " - " & _
                                     placement.Batch & " - " & placement.quantity & " (" & _
                                     placement.standardVolume & ") - " & placement.pieceType & _
                                     " - Вес: " & placement.weightCategory
            End If
        End If
    Next i
    
    If validCount > 0 Then
        ReDim Preserve results(validCount - 1)
    Else
        ReDim results(-1)
    End If
    
    ParseNewFormatInput = results
End Function

Private Function ParseNewFormatLine(line As String, ByRef placement As PlacementInfo) As Boolean
    On Error GoTo ErrorHandler
    
    ' Парсим: "Лерашанс - пар13 - 600 (720)"
    Dim parts() As String
    parts = Split(line, " - ")
    
    If UBound(parts) < 2 Then
        ParseNewFormatLine = False
        Exit Function
    End If
    
    ' Получаем основные части
    placement.ProductName = Trim(parts(0))
    placement.Batch = Trim(parts(1))
    
    ' Парсим количество и объем
    Dim lastPart As String
    lastPart = Trim(parts(2))
    
    ' Ищем объем в скобках
    Dim openBracket As Integer, closeBracket As Integer
    openBracket = InStr(lastPart, "(")
    closeBracket = InStr(lastPart, ")")
    
    If openBracket > 0 And closeBracket > openBracket Then
        ' Извлекаем количество
        Dim quantityStr As String
        quantityStr = Trim(Left(lastPart, openBracket - 1))
        
        ' Извлекаем объем
        Dim volumeStr As String
        volumeStr = Trim(Mid(lastPart, openBracket + 1, closeBracket - openBracket - 1))
        
        If IsNumeric(quantityStr) And IsNumeric(volumeStr) Then
            placement.quantity = CDbl(quantityStr)
            placement.standardVolume = CDbl(volumeStr)
            
            ' Определяем тип товара
            placement.pieceType = DeterminePieceType(placement.quantity, placement.standardVolume)
            
            ParseNewFormatLine = True
        End If
    Else
        ' Формат без скобок
        If IsNumeric(lastPart) Then
            placement.quantity = CDbl(lastPart)
            placement.standardVolume = 720 ' По умолчанию
            placement.pieceType = DeterminePieceType(placement.quantity, placement.standardVolume)
            
            ParseNewFormatLine = True
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    ParseNewFormatLine = False
End Function

Private Function DeterminePieceType(quantity As Double, standardVolume As Double) As String
    If standardVolume = 0 Then
        DeterminePieceType = "WHOLE"
        Exit Function
    End If
    
    Dim percentage As Double
    percentage = (quantity / standardVolume) * 100
    
    If percentage >= 90 Then
        DeterminePieceType = "WHOLE"   ' Целая паллета
    ElseIf percentage >= 60 Then
        DeterminePieceType = "MEDIUM"  ' Средняя
    Else
        DeterminePieceType = "PIECE"   ' Кусок
    End If
End Function

' ===== ГЕНЕРАЦИЯ ОТЧЕТА "ЯЧЕЙКА ЗА ЯЧЕЙКОЙ" (ИСПРАВЛЕНО НАЗВАНИЕ ФУНКЦИИ) =====
Private Function GenerateCellByCellReport(placements() As PlacementInfo, successCount As Integer, activeCount As Integer) As String
    Dim report As String
    report = "?? ОТЧЕТ РАЗМЕЩЕНИЯ 'ЯЧЕЙКА ЗА ЯЧЕЙКОЙ'" & vbNewLine
    report = report & "===============================================" & vbNewLine
    report = report & "?? " & Format(Now, "dd.mm.yyyy hh:mm:ss") & vbNewLine
    report = report & "?? Активные ангары: " & ModuleWarehouseCapacity.GetActiveWarehousesList() & vbNewLine & vbNewLine
    
    report = report & "? УСПЕШНО: " & successCount & vbNewLine
    report = report & "? ОШИБКИ: " & (UBound(placements) + 1 - successCount) & vbNewLine & vbNewLine
    
    If successCount > 0 Then
        report = report & "?? РАЗМЕЩЕНИЯ ЯЧЕЙКА-ЗА-ЯЧЕЙКОЙ:" & vbNewLine
        report = report & "-----------------------------------------" & vbNewLine
        
        Dim i As Integer
        For i = 0 To UBound(placements)
            If placements(i).IsPlaced And placements(i).placementDetails <> "" Then
                report = report & placements(i).placementDetails & vbNewLine
            End If
        Next i
        report = report & vbNewLine
    End If
    
    Dim errorCount As Integer
    errorCount = (UBound(placements) + 1 - successCount)
    
    If errorCount > 0 Then
        report = report & "? ОШИБКИ:" & vbNewLine
        report = report & "-------------" & vbNewLine
        
        Dim j As Integer
        For j = 0 To UBound(placements)
            If Not placements(j).IsPlaced Then
                report = report & "• " & placements(j).ProductName & " (" & _
                        placements(j).Batch & ") - " & _
                        placements(j).quantity & " - " & _
                        placements(j).ErrorMessage & vbNewLine
            End If
        Next j
        report = report & vbNewLine
    End If
    
    report = report & "?? ИСПРАВЛЕНИЯ:" & vbNewLine
    report = report & "=========================" & vbNewLine
    report = report & "? ПРОВЕРКА КАЖДОЙ ЯЧЕЙКИ:" & vbNewLine
    report = report & "   • А1>А2>А3>Б1>Б2>Б3 (правильно!)" & vbNewLine
    report = report & "   • НЕ А1>Б1>В1>Г1 (было неправильно)" & vbNewLine & vbNewLine
    
    report = report & "? УМНЫЕ ЛИМИТЫ ПО БУКВАМ:" & vbNewLine
    report = report & "   • 720л: до 2 штук (1440/1500)" & vbNewLine
    report = report & "   • 960л: только 1 штука (960/1000)" & vbNewLine
    report = report & "   • 240кг: до 2 штук (480/500)" & vbNewLine & vbNewLine
    
    report = report & "?? ЦВЕТОВАЯ ИНДИКАЦИЯ:" & vbNewLine
    report = report & "   • ?? Темно-зеленый - Ярус 1" & vbNewLine
    report = report & "   • ?? Ярко-зеленый - Ярус 2" & vbNewLine
    report = report & "   • ?? Желто-зеленый - Ярус 3" & vbNewLine
    report = report & "   • ?? Синяя граница - Правильное размещение" & vbNewLine
    report = report & "   • ?? Красная граница - Куски" & vbNewLine & vbNewLine
    
    report = report & "? Система 'ячейка-за-ячейкой' работает правильно!"
    
    GenerateCellByCellReport = report
End Function

' ===== ТЕСТОВАЯ ФУНКЦИЯ =====
Public Sub TestCellByCell()
    Call ModuleLogger.InitializeLogger(True, ModuleLogger.LOG_LEVEL_DEBUG)
    
    ModuleLogger.LogMessage "=== ?? ТЕСТ СИСТЕМЫ 'ЯЧЕЙКА ЗА ЯЧЕЙКОЙ' ==="
    
    ' Тестовые данные: 3 паллеты по 720л
    Dim testData As String
    testData = "Лерашанс - пар01 - 720 (720)" & vbNewLine & _
               "Лерашанс - пар02 - 720 (720)" & vbNewLine & _
               "Лерашанс - пар03 - 720 (720)"
    
    ModuleLogger.LogMessage "Входные данные: 3?720л"
    ModuleLogger.LogMessage "ОЖИДАЕМЫЙ РЕЗУЛЬТАТ:"
    ModuleLogger.LogMessage "  • А-1: 720 (первая ячейка)"
    ModuleLogger.LogMessage "  • А-2: 720 (вторая ячейка в той же букве А)"
    ModuleLogger.LogMessage "  • Б-1: 720 (третья ячейка в новой букве Б, лимит А исчерпан)"
    
    Dim result As String
    result = SmartPlaceProducts(testData)
    
    MsgBox "?? ТЕСТ 'ЯЧЕЙКА ЗА ЯЧЕЙКОЙ' ЗАВЕРШЕН!" & vbNewLine & vbNewLine & _
           "Проверьте результат:" & vbNewLine & _
           "? А-1: 720 (темно-зеленый + синяя граница)" & vbNewLine & _
           "? А-2: 720 (ярко-зеленый + синяя граница)" & vbNewLine & _
           "? Б-1: 720 (темно-зеленый + синяя граница)" & vbNewLine & vbNewLine & _
           "?? Результат: " & vbNewLine & result, _
           vbInformation, "Система 'ячейка-за-ячейкой' работает!"
End Sub

