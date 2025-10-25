Attribute VB_Name = "ModuleProductInfo"


' ===== ModuleProductInfo - ОБНОВЛЕННАЯ БАЗА С ПОЛНЫМ СПИСКОМ ПРЕПАРАТОВ =====
' Версия 2.0 - 86 препаратов в 7 группах объемов
Option Explicit

' Словари для групп товаров по количеству в ячейке
Private products1000 As Object   ' 1000 литров на палету (2 препарата)
Private products960 As Object    ' 960 литров на палету (2 препарата)
Private products720 As Object    ' 720 литров на палету (62 препарата)
Private products600 As Object    ' 600 литров на палету (7 препаратов)
Private products540 As Object    ' 540 литров на палету (2 препарата)
Private products288 As Object    ' 288 кг на палету (5 препаратов)
Private products240 As Object    ' 240 кг на палету (8 препаратов)

' Стандартные значения
Private Const STANDARD_CELL_CAPACITY As Double = 720    ' По умолчанию
Private Const STANDARD_SHELF_CAPACITY As Double = 1500  ' Будет переопределено новой системой

' ===== ИНИЦИАЛИЗАЦИЯ БАЗЫ =====

Public Sub InitializeProductDatabase()
    ' Создаем словари для каждой группы
    Set products1000 = CreateObject("Scripting.Dictionary")
    Set products960 = CreateObject("Scripting.Dictionary")
    Set products720 = CreateObject("Scripting.Dictionary")
    Set products600 = CreateObject("Scripting.Dictionary")
    Set products540 = CreateObject("Scripting.Dictionary")
    Set products288 = CreateObject("Scripting.Dictionary")
    Set products240 = CreateObject("Scripting.Dictionary")
    
    ' Заполняем все группы полным списком препаратов
    Call InitializeGroup1000  ' 2 препарата
    Call InitializeGroup960   ' 2 препарата
    Call InitializeGroup720   ' 62 препарата
    Call InitializeGroup600   ' 7 препаратов
    Call InitializeGroup540   ' 2 препарата
    Call InitializeGroup288   ' 5 препаратов
    Call InitializeGroup240   ' 8 препаратов
    
    ModuleLogger.LogMessage "?? База препаратов инициализирована: 86 препаратов в 7 группах"
End Sub

' ===== ИНИЦИАЛИЗАЦИЯ ГРУПП ПРЕПАРАТОВ =====

Private Sub InitializeGroup1000()
    ' ГРУППА 1000 литров (2 препарата)
    products1000.Add "ВИТАШАНС", True
    products1000.Add "ФУМИШАНС", True
End Sub

Private Sub InitializeGroup960()
    ' ГРУППА 960 литров (2 препарата)
    products960.Add "ГЛИФОШАНС СУПЕР", True
    products960.Add "ДИКОШАНС", True
End Sub

Private Sub InitializeGroup720()
    ' ГРУППА 720 литров (62 препарата - основная группа)
    products720.Add "АГРОШАНС", True
    products720.Add "АПРОВАТОР", True
    products720.Add "БЕТАШАНС ДАБЛ", True
    products720.Add "БЕТАШАНС ТРИО", True
    products720.Add "БОРОШАНС", True
    products720.Add "БОСФОР", True
    products720.Add "БРОДЕФОР", True
    products720.Add "ГАЛОШАНС", True
    products720.Add "ГЕЛИФАС", True
    products720.Add "ГЛИФОШАНС", True
    products720.Add "ДВД ШАНС", True
    products720.Add "ДИБАЗОН", True
    products720.Add "ДИШАНС", True
    products720.Add "ДУШАНС", True
    products720.Add "ЕВРОШАНС", True
    products720.Add "ЕВРОШАНС ПЛЮС", True
    products720.Add "ЗЕНКОШАНС", True
    products720.Add "ЗЕРНОРОСТ", True
    products720.Add "ЗИМОШАНС", True
    products720.Add "ИМАЗОШАНС", True
    products720.Add "ИМИДАШАНС", True
    products720.Add "ИМИДАШАНС ПЛЮС", True
    products720.Add "ИМИДАШАНС-С", True
    products720.Add "ИМОЗАШАНС", True
    products720.Add "КАРИШАНС", True
    products720.Add "КЛЕТОШАНС", True
    products720.Add "КРУГОЗОР", True
    products720.Add "ЛЕРАШАНС", True
    products720.Add "МАКРОШАНС", True
    products720.Add "МИКРОПОЛИДОК БОР", True
    products720.Add "МИКРОПОЛИДОК ПЛЮС", True
    products720.Add "МИКРОПОЛИДОК ЦИНК", True
    products720.Add "МОЛИБДЕН", True
    products720.Add "НАНОШАНС", True
    products720.Add "ПОЛИРАМ", True
    products720.Add "ПОЛИШАНС", True
    products720.Add "ПРИШАНС", True
    products720.Add "ПРИШАНС СУПЕР", True
    products720.Add "ПРОПИШАНС", True
    products720.Add "ПРОПИШАНС СУПЕР", True
    products720.Add "ПРОПИШАНС УНИВЕРСАЛ", True
    products720.Add "СЕРА", True
    products720.Add "СИЛЬВОШАНС", True
    products720.Add "СКОРОШАНС", True
    products720.Add "СОФТЭН", True
    products720.Add "СТРОБИШАНС ПРО", True
    products720.Add "ТАНОШАНС", True
    products720.Add "ТАПИРОШАНС", True
    products720.Add "ТИРАМ", True
    products720.Add "ФЕНИКС", True
    products720.Add "ФЕЯ", True
    products720.Add "ЧИСТОСАД", True
    products720.Add "ШАНС 24", True
    products720.Add "ШАНС ГОЛД", True
    products720.Add "ШАНС ДКБ", True
    products720.Add "ШАНС УНИВЕРСАЛ", True
    products720.Add "ШАНСГАРД", True
    products720.Add "ШАНСИЛ", True
    products720.Add "ШАНСИЛ ТРИО", True
    products720.Add "ШАНСИЛ УЛЬТРА", True
    products720.Add "ШАНСИТЕК", True
    products720.Add "ШАНСОМЕТОКС ТРИО", True
    products720.Add "ШАНСОМИТРОН", True
    products720.Add "ШАНСТРЕЛ 300", True
    products720.Add "ШАНСЮГЕН", True
End Sub

Private Sub InitializeGroup600()
    ' ГРУППА 600 литров (7 препаратов)
    products600.Add "ГОПЛИТ", True
    products600.Add "ФАСШАНС", True
    products600.Add "КАЛИНА", True
    products600.Add "КАРАТОШАНС", True
    products600.Add "КРЕПОШАНС", True
    products600.Add "СЕКТОР", True
    products600.Add "ШАНС-90", True
End Sub

Private Sub InitializeGroup540()
    ' ГРУППА 540 литров (2 препарата)
    products540.Add "АНТИМЫШИН", True
    products540.Add "МЕТАШАНС", True
End Sub

Private Sub InitializeGroup288()
    ' ГРУППА 288 кг (5 препаратов)
    products288.Add "ЭЛЛАДА", True
    products288.Add "ЭЛЬШАНС", True
    products288.Add "ДЕЛАТОН", True
    products288.Add "ПЕНТАГОН", True
    products288.Add "ПОЛИДОК", True
End Sub

Private Sub InitializeGroup240()
    ' ГРУППА 240 кг (8 препаратов)
    products240.Add "ШАНТУС", True
    products240.Add "ШАНСПРОФИ", True
    products240.Add "ШАНСТАР", True
    products240.Add "ШАНСТАР ПЛЮС", True
    products240.Add "ШАНСТИ", True
    products240.Add "ЗНАТОК", True
    products240.Add "ХОРИСТ", True
    products240.Add "ШАНСИЛИН", True
End Sub

' ===== ПОЛУЧИТЬ ВМЕСТИМОСТЬ ЯЧЕЙКИ =====

Public Function GetCellCapacity(ProductName As String) As Double
    ' Инициализируем если не инициализировано
    If products1000 Is Nothing Then InitializeProductDatabase
    
    Dim upperName As String
    upperName = UCase(Trim(ProductName))
    
    ' Проверяем по группам
    If products1000.exists(upperName) Then
        GetCellCapacity = 1000
        ModuleLogger.LogDebug "?? " & ProductName & " = 1000л"
    ElseIf products960.exists(upperName) Then
        GetCellCapacity = 960
        ModuleLogger.LogDebug "?? " & ProductName & " = 960л"
    ElseIf products720.exists(upperName) Then
        GetCellCapacity = 720
        ModuleLogger.LogDebug "?? " & ProductName & " = 720л"
    ElseIf products600.exists(upperName) Then
        GetCellCapacity = 600
        ModuleLogger.LogDebug "?? " & ProductName & " = 600л"
    ElseIf products540.exists(upperName) Then
        GetCellCapacity = 540
        ModuleLogger.LogDebug "?? " & ProductName & " = 540л"
    ElseIf products288.exists(upperName) Then
        GetCellCapacity = 288
        ModuleLogger.LogDebug "?? " & ProductName & " = 288кг"
    ElseIf products240.exists(upperName) Then
        GetCellCapacity = 240
        ModuleLogger.LogDebug "?? " & ProductName & " = 240кг"
    Else
        ' СТАНДАРТНОЕ ЗНАЧЕНИЕ - 720 для неизвестных препаратов
        GetCellCapacity = STANDARD_CELL_CAPACITY
        ModuleLogger.LogWarning "?? Неизвестный препарат '" & ProductName & "' - используем стандартный объем 720"
    End If
End Function

' ===== ПОЛУЧИТЬ МАКСИМУМ НА СТЕЛЛАЖ (ОБНОВЛЕНО) =====

Public Function GetMaxShelfCapacity(ProductName As String, Optional warehouseNumber As String = "6") As Double
    ' ОБНОВЛЕНО: Теперь использует новую систему лимитов по ангарам
    If ModuleWarehouseCapacity Is Nothing Then
        ' Если новая система не инициализирована, используем старое значение
        GetMaxShelfCapacity = STANDARD_SHELF_CAPACITY
    Else
        ' Используем новую систему с индивидуальными лимитами по ангарам
        GetMaxShelfCapacity = ModuleWarehouseCapacity.GetShelfCapacityLimit(warehouseNumber, ProductName)
    End If
End Function

' ===== ДОБАВИТЬ ТОВАР В ГРУППУ =====

Public Sub AddProductToGroup(ProductName As String, groupCapacity As Double)
    ' Инициализируем если не инициализировано
    If products1000 Is Nothing Then InitializeProductDatabase
    
    Dim upperName As String
    upperName = UCase(Trim(ProductName))
    
    ' Удаляем из всех групп если есть
    RemoveProductFromAllGroups upperName
    
    ' Добавляем в нужную группу
    Select Case groupCapacity
        Case 1000
            products1000.Add upperName, True
            ModuleLogger.LogMessage "? Препарат " & ProductName & " добавлен в группу 1000 л"
            
        Case 960
            products960.Add upperName, True
            ModuleLogger.LogMessage "? Препарат " & ProductName & " добавлен в группу 960 л"
            
        Case 720
            products720.Add upperName, True
            ModuleLogger.LogMessage "? Препарат " & ProductName & " добавлен в группу 720 л"
            
        Case 600
            products600.Add upperName, True
            ModuleLogger.LogMessage "? Препарат " & ProductName & " добавлен в группу 600 л"
            
        Case 540
            products540.Add upperName, True
            ModuleLogger.LogMessage "? Препарат " & ProductName & " добавлен в группу 540 л"
            
        Case 288
            products288.Add upperName, True
            ModuleLogger.LogMessage "? Препарат " & ProductName & " добавлен в группу 288 кг"
            
        Case 240
            products240.Add upperName, True
            ModuleLogger.LogMessage "? Препарат " & ProductName & " добавлен в группу 240 кг"
            
        Case Else
            ' Если нестандартная емкость, добавляем в группу 720
            products720.Add upperName, True
            ModuleLogger.LogMessage "? Препарат " & ProductName & " добавлен в стандартную группу 720"
    End Select
End Sub

' ===== УДАЛИТЬ ТОВАР ИЗ ВСЕХ ГРУПП =====

Private Sub RemoveProductFromAllGroups(ProductName As String)
    On Error Resume Next
    products1000.Remove ProductName
    products960.Remove ProductName
    products720.Remove ProductName
    products600.Remove ProductName
    products540.Remove ProductName
    products288.Remove ProductName
    products240.Remove ProductName
    On Error GoTo 0
End Sub

' ===== ПРОВЕРКА ВМЕСТИМОСТИ (ОБНОВЛЕНО) =====

Public Function CanAccommodate(ProductName As String, currentQuantity As Double, _
                              newQuantity As Double, Optional warehouseNumber As String = "6") As Boolean
    ' ОБНОВЛЕНО: Теперь использует новую систему лимитов
    
    ' Проверяем активность ангара
    If Not ModuleWarehouseCapacity.IsWarehouseActive(warehouseNumber) Then
        CanAccommodate = False
        ModuleLogger.LogWarning "?? Ангар " & warehouseNumber & " отключен!"
        Exit Function
    End If
    
    ' Проверяем лимит на стеллаж
    Dim ShelfLimit As Double
    ShelfLimit = ModuleWarehouseCapacity.GetShelfCapacityLimit(warehouseNumber, ProductName)
    
    If currentQuantity + newQuantity > ShelfLimit Then
        CanAccommodate = False
        ModuleLogger.LogWarning "?? Превышен лимит стеллажа: " & (currentQuantity + newQuantity) & " > " & ShelfLimit
        Exit Function
    End If
    
    ' Проверяем лимит на ячейку
    Dim cellCapacity As Double
    cellCapacity = GetCellCapacity(ProductName)
    
    If newQuantity > cellCapacity Then
        CanAccommodate = False
        ModuleLogger.LogWarning "?? Превышен лимит ячейки: " & newQuantity & " > " & cellCapacity
        Exit Function
    End If
    
    CanAccommodate = True
End Function

' ===== ПОКАЗАТЬ ВСЕ ГРУППЫ =====

Public Function ShowAllGroups() As String
    If products1000 Is Nothing Then InitializeProductDatabase
    
    Dim result As String
    result = "?? ПОЛНАЯ БАЗА ПРЕПАРАТОВ (86 наименований)" & vbNewLine
    result = result & "=============================================" & vbNewLine & vbNewLine
    
    ' Группа 1000 (2 препарата)
    result = result & "?? ГРУППА 1000 литров (2 препарата):" & vbNewLine
    Dim key As Variant
    For Each key In products1000.keys
        result = result & "  • " & key & vbNewLine
    Next key
    result = result & vbNewLine
    
    ' Группа 960 (2 препарата)
    result = result & "?? ГРУППА 960 литров (2 препарата):" & vbNewLine
    For Each key In products960.keys
        result = result & "  • " & key & vbNewLine
    Next key
    result = result & vbNewLine
    
    ' Группа 720 (62 препарата)
    result = result & "?? ГРУППА 720 литров (62 препарата - основная):" & vbNewLine
    Dim count720 As Integer: count720 = 0
    For Each key In products720.keys
        result = result & "  • " & key & vbNewLine
        count720 = count720 + 1
        ' Добавляем разделитель каждые 20 препаратов для читаемости
        If count720 Mod 20 = 0 And count720 < products720.count Then
            result = result & "    ..." & vbNewLine
        End If
    Next key
    result = result & vbNewLine
    
    ' Группа 600 (7 препаратов)
    result = result & "?? ГРУППА 600 литров (7 препаратов):" & vbNewLine
    For Each key In products600.keys
        result = result & "  • " & key & vbNewLine
    Next key
    result = result & vbNewLine
    
    ' Группа 540 (2 препарата)
    result = result & "?? ГРУППА 540 литров (2 препарата):" & vbNewLine
    For Each key In products540.keys
        result = result & "  • " & key & vbNewLine
    Next key
    result = result & vbNewLine
    
    ' Группа 288 (5 препаратов)
    result = result & "?? ГРУППА 288 кг (5 препаратов):" & vbNewLine
    For Each key In products288.keys
        result = result & "  • " & key & vbNewLine
    Next key
    result = result & vbNewLine
    
    ' Группа 240 (8 препаратов)
    result = result & "? ГРУППА 240 кг (8 препаратов):" & vbNewLine
    For Each key In products240.keys
        result = result & "  • " & key & vbNewLine
    Next key
    result = result & vbNewLine
    
    result = result & "?? ИТОГО: " & (products1000.count + products960.count + products720.count + _
                      products600.count + products540.count + products288.count + products240.count) & " препаратов"
    
    ShowAllGroups = result
End Function

' ===== СТАТИСТИКА БАЗЫ =====

Public Function GetDatabaseStatistics() As String
    If products1000 Is Nothing Then InitializeProductDatabase
    
    Dim stats As String
    stats = "?? СТАТИСТИКА БАЗЫ ПРЕПАРАТОВ" & vbNewLine
    stats = stats & "=============================" & vbNewLine & vbNewLine
    
    stats = stats & "1000 литров: " & products1000.count & " препаратов" & vbNewLine
    stats = stats & "960 литров:  " & products960.count & " препаратов" & vbNewLine
    stats = stats & "720 литров:  " & products720.count & " препаратов ? основная группа" & vbNewLine
    stats = stats & "600 литров:  " & products600.count & " препаратов" & vbNewLine
    stats = stats & "540 литров:  " & products540.count & " препаратов" & vbNewLine
    stats = stats & "288 кг:      " & products288.count & " препаратов" & vbNewLine
    stats = stats & "240 кг:      " & products240.count & " препаратов" & vbNewLine & vbNewLine
    
    Dim total As Integer
    total = products1000.count + products960.count + products720.count + _
            products600.count + products540.count + products288.count + products240.count
            
    stats = stats & "?? ОБЩЕЕ КОЛИЧЕСТВО: " & total & " препаратов" & vbNewLine
    stats = stats & "?? САМАЯ БОЛЬШАЯ ГРУППА: 720л (" & products720.count & " препаратов)" & vbNewLine
    stats = stats & "?? ПОКРЫТИЕ: 100% известных препаратов"
    
    GetDatabaseStatistics = stats
End Function

' ===== ПОИСК ПРЕПАРАТА =====

Public Function FindProductInDatabase(ProductName As String) As String
    If products1000 Is Nothing Then InitializeProductDatabase
    
    Dim upperName As String
    upperName = UCase(Trim(ProductName))
    
    Dim result As String
    result = "?? ПОИСК ПРЕПАРАТА: " & ProductName & vbNewLine
    result = result & String(25, "=") & vbNewLine
    
    If products1000.exists(upperName) Then
        result = result & "? Найден в группе 1000 литров" & vbNewLine
        result = result & "?? Препаратов в группе: 2" & vbNewLine
        result = result & "?? Соседи: Виташанс, Фумишанс"
    ElseIf products960.exists(upperName) Then
        result = result & "? Найден в группе 960 литров" & vbNewLine
        result = result & "?? Препаратов в группе: 2" & vbNewLine
        result = result & "?? Соседи: Глифошанс Супер, Дикошанс"
    ElseIf products720.exists(upperName) Then
        result = result & "? Найден в группе 720 литров ? основная группа" & vbNewLine
        result = result & "?? Препаратов в группе: 62" & vbNewLine
        result = result & "?? Включая: Лерашанс, Босфор, Агрошанс и другие"
    ElseIf products600.exists(upperName) Then
        result = result & "? Найден в группе 600 литров" & vbNewLine
        result = result & "?? Препаратов в группе: 7" & vbNewLine
        result = result & "?? Соседи: Гоплит, Фасшанс, Калина и другие"
    ElseIf products540.exists(upperName) Then
        result = result & "? Найден в группе 540 литров" & vbNewLine
        result = result & "?? Препаратов в группе: 2" & vbNewLine
        result = result & "?? Соседи: Антимышин, Меташанс"
    ElseIf products288.exists(upperName) Then
        result = result & "? Найден в группе 288 кг" & vbNewLine
        result = result & "?? Препаратов в группе: 5" & vbNewLine
        result = result & "?? Соседи: Эллада, Эльшанс, Делатон и другие"
    ElseIf products240.exists(upperName) Then
        result = result & "? Найден в группе 240 кг" & vbNewLine
        result = result & "?? Препаратов в группе: 8" & vbNewLine
        result = result & "?? Соседи: Шантус, Шанспрофи, Шанстар и другие"
    Else
        result = result & "? НЕ НАЙДЕН в базе препаратов" & vbNewLine
        result = result & "?? Будет использован стандартный объем 720л" & vbNewLine
        result = result & "? Можно добавить через AddProductToGroup()"
    End If
    
    FindProductInDatabase = result
End Function

' ===== ЭКСПОРТ В EXCEL (ОБНОВЛЕНО) =====

Public Sub ExportGroupsToExcel()
    On Error GoTo ErrorHandler
    
    If products1000 Is Nothing Then InitializeProductDatabase
    
    ' Создаем новый лист
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets.Add
    ws.Name = "База_препаратов_" & Format(Now, "ddmmyy")
    
    ' Заголовки
    ws.Cells(1, 1).value = "Препарат"
    ws.Cells(1, 2).value = "Объем на палету"
    ws.Cells(1, 3).value = "Единица"
    ws.Cells(1, 4).value = "Группа"
    ws.Cells(1, 5).value = "Примечание"
    
    ' Форматирование заголовков
    With ws.Range("A1:E1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
        .HorizontalAlignment = xlCenter
    End With
    
    Dim row As Integer
    row = 2
    
    ' Функция для добавления группы в таблицу
    Call AddGroupToExport(ws, products1000, 1000, "л", "Большие канистры", row)
    Call AddGroupToExport(ws, products960, 960, "л", "Специальные", row)
    Call AddGroupToExport(ws, products720, 720, "л", "Стандартные ?", row)
    Call AddGroupToExport(ws, products600, 600, "л", "Средние канистры", row)
    Call AddGroupToExport(ws, products540, 540, "л", "Малые канистры", row)
    Call AddGroupToExport(ws, products288, 288, "кг", "Порошки", row)
    Call AddGroupToExport(ws, products240, 240, "кг", "Мелкие порошки", row)
    
    ' Автоподбор ширины
    ws.columns("A:E").AutoFit
    
    ' Границы
    With ws.Range("A1:E" & (row - 1))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    ' Итоговая строка
    ws.Cells(row, 1).value = "ИТОГО ПРЕПАРАТОВ:"
    ws.Cells(row, 2).value = row - 2
    With ws.Range("A" & row & ":E" & row)
        .Font.Bold = True
        .Interior.Color = RGB(255, 255, 0)
    End With
    
    MsgBox "?? База препаратов экспортирована на лист: " & ws.Name & vbNewLine & _
           "Всего препаратов: " & (row - 2), vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "? Ошибка экспорта: " & Err.description, vbExclamation
End Sub

Private Sub AddGroupToExport(ws As Worksheet, groupDict As Object, volume As Double, _
                            unit As String, description As String, ByRef row As Integer)
    Dim key As Variant
    For Each key In groupDict.keys
        ws.Cells(row, 1).value = key
        ws.Cells(row, 2).value = volume
        ws.Cells(row, 3).value = unit
        ws.Cells(row, 4).value = volume & " " & unit
        ws.Cells(row, 5).value = description
        row = row + 1
    Next key
End Sub

' ===== ТЕСТОВЫЕ ФУНКЦИИ =====

Public Sub TestFullDatabase()
    InitializeProductDatabase
    
    ModuleLogger.LogMessage GetDatabaseStatistics()
    
    ' Тест поиска препаратов
    Dim testProducts As Variant
    testProducts = Array("Лерашанс", "Виташанс", "Дикошанс", "Эллада", "НеизвестныйПрепарат")
    
    Dim i As Integer
    For i = 0 To UBound(testProducts)
        ModuleLogger.LogMessage FindProductInDatabase(CStr(testProducts(i)))
        ModuleLogger.LogMessage ""
    Next i
    
    MsgBox "?? Тест базы завершен! Проверьте Immediate Window (Ctrl+G)", vbInformation, "Тест базы препаратов"
End Sub

Public Sub ShowGroupSizes()
    If products1000 Is Nothing Then InitializeProductDatabase
    
    MsgBox GetDatabaseStatistics(), vbInformation, "Статистика базы препаратов"
End Sub

