Attribute VB_Name = "ModuleWarehouseCapacity"
' ===== ModuleWarehouseCapacity - ИСПРАВЛЕННАЯ ВЕРСИЯ БЕЗ TYPE =====
Option Explicit

' ===== ГЛОБАЛЬНЫЕ ПЕРЕМЕННЫЕ ВМЕСТО ПОЛЬЗОВАТЕЛЬСКИХ ТИПОВ =====
Private warehouseConfigs As Object        ' Конфигурации всех ангаров
Private warehouseActiveStates As Object   ' Состояние активности ангаров

' ===== ИНИЦИАЛИЗАЦИЯ СИСТЕМЫ ЛИМИТОВ =====

Public Sub InitializeWarehouseCapacities()
    ' Инициализируем главные словари
    Set warehouseConfigs = CreateObject("Scripting.Dictionary")
    Set warehouseActiveStates = CreateObject("Scripting.Dictionary")
    
    ' Настраиваем каждый ангар с индивидуальными лимитами
    Call SetupWarehouse5
    Call SetupWarehouse6
    Call SetupWarehouse7
    Call SetupWarehouse8
    Call SetupWarehouse9
    Call SetupWarehouse10
    Call SetupWarehouse11
    Call SetupWarehouse12
    
    ' По умолчанию все ангары активны
    Dim i As Integer
    For i = 5 To 12
        warehouseActiveStates.Add CStr(i), True
    Next i
    
    ModuleLogger.LogMessage "? Система лимитов ангаров инициализирована (7 групп объемов, 86 препаратов)"
End Sub

' ===== НАСТРОЙКА ОТДЕЛЬНЫХ АНГАРОВ (БЕЗ ПОЛЬЗОВАТЕЛЬСКИХ ТИПОВ) =====

Private Sub SetupWarehouse5()
    Dim warehouseData As Object
    Set warehouseData = CreateObject("Scripting.Dictionary")
    
    ' Основные параметры
    warehouseData.Add "number", "5"
    warehouseData.Add "isActive", True
    
    ' Создаем словари для лимитов
    Dim shelfLimits As Object
    Dim rowLimits As Object
    Set shelfLimits = CreateObject("Scripting.Dictionary")
    Set rowLimits = CreateObject("Scripting.Dictionary")
    
    ' АНГАР 5 - Лимиты на стеллаж ("В буквах")
    shelfLimits.Add 1000, 2000   ' Виташанс, Фумишанс
    shelfLimits.Add 960, 1000    ' Глифошанс Супер, Дикошанс
    shelfLimits.Add 720, 1500    ' Лерашанс, Босфор и 60 других
    shelfLimits.Add 600, 1300    ' Гоплит, Фасшанс и 5 других
    shelfLimits.Add 540, 1100    ' Антимышин, Меташанс
    shelfLimits.Add 288, 580     ' Эллада, Эльшанс и 3 других
    shelfLimits.Add 240, 500     ' Шантус, Шанспрофи и 6 других
    
    ' АНГАР 5 - Лимиты на ряд ("В ряд")
    rowLimits.Add 1000, 14000
    rowLimits.Add 960, 6800
    rowLimits.Add 720, 10100
    rowLimits.Add 600, 8500
    rowLimits.Add 540, 8700
    rowLimits.Add 288, 4608
    rowLimits.Add 240, 3850
    
    ' Сохраняем лимиты в структуре ангара
    warehouseData.Add "shelfLimits", shelfLimits
    warehouseData.Add "rowLimits", rowLimits
    
    ' Добавляем в общий реестр
    warehouseConfigs.Add "5", warehouseData
End Sub

Private Sub SetupWarehouse6()
    Dim warehouseData As Object
    Set warehouseData = CreateObject("Scripting.Dictionary")
    
    warehouseData.Add "number", "6"
    warehouseData.Add "isActive", True
    
    Dim shelfLimits As Object
    Dim rowLimits As Object
    Set shelfLimits = CreateObject("Scripting.Dictionary")
    Set rowLimits = CreateObject("Scripting.Dictionary")
    
    ' АНГАР 6 - Идентичные лимиты с ангаром 5
    shelfLimits.Add 1000, 2000
    shelfLimits.Add 960, 1000
    shelfLimits.Add 720, 1500
    shelfLimits.Add 600, 1300
    shelfLimits.Add 540, 1100
    shelfLimits.Add 288, 580
    shelfLimits.Add 240, 500
    
    rowLimits.Add 1000, 14000
    rowLimits.Add 960, 6800
    rowLimits.Add 720, 10100
    rowLimits.Add 600, 8500
    rowLimits.Add 540, 8700
    rowLimits.Add 288, 4608
    rowLimits.Add 240, 3850
    
    warehouseData.Add "shelfLimits", shelfLimits
    warehouseData.Add "rowLimits", rowLimits
    warehouseConfigs.Add "6", warehouseData
End Sub

Private Sub SetupWarehouse7()
    Dim warehouseData As Object
    Set warehouseData = CreateObject("Scripting.Dictionary")
    
    warehouseData.Add "number", "7"
    warehouseData.Add "isActive", True
    
    Dim shelfLimits As Object
    Dim rowLimits As Object
    Set shelfLimits = CreateObject("Scripting.Dictionary")
    Set rowLimits = CreateObject("Scripting.Dictionary")
    
    ' АНГАР 7 - Лимиты на стеллаж (те же что у 5,6)
    shelfLimits.Add 1000, 2000
    shelfLimits.Add 960, 1000
    shelfLimits.Add 720, 1500
    shelfLimits.Add 600, 1300
    shelfLimits.Add 540, 1100
    shelfLimits.Add 288, 580
    shelfLimits.Add 240, 500
    
    ' АНГАР 7 - Меньшие лимиты на ряд
    rowLimits.Add 1000, 12000  ' Меньше чем в 5-6
    rowLimits.Add 960, 5800
    rowLimits.Add 720, 8640
    rowLimits.Add 600, 7200
    rowLimits.Add 540, 6500
    rowLimits.Add 288, 4000
    rowLimits.Add 240, 3000
    
    warehouseData.Add "shelfLimits", shelfLimits
    warehouseData.Add "rowLimits", rowLimits
    warehouseConfigs.Add "7", warehouseData
End Sub

Private Sub SetupWarehouse8()
    Dim warehouseData As Object
    Set warehouseData = CreateObject("Scripting.Dictionary")
    
    warehouseData.Add "number", "8"
    warehouseData.Add "isActive", True
    
    Dim shelfLimits As Object
    Dim rowLimits As Object
    Set shelfLimits = CreateObject("Scripting.Dictionary")
    Set rowLimits = CreateObject("Scripting.Dictionary")
    
    ' АНГАР 8 - Лимиты на стеллаж (те же)
    shelfLimits.Add 1000, 2000
    shelfLimits.Add 960, 1000
    shelfLimits.Add 720, 1500
    shelfLimits.Add 600, 1300
    shelfLimits.Add 540, 1100
    shelfLimits.Add 288, 580
    shelfLimits.Add 240, 500
    
    ' АНГАР 8 - Самые маленькие лимиты на ряд
    rowLimits.Add 1000, 10000  ' Самые маленькие
    rowLimits.Add 960, 5000
    rowLimits.Add 720, 7200
    rowLimits.Add 600, 6000
    rowLimits.Add 540, 6500
    rowLimits.Add 288, 3800
    rowLimits.Add 240, 3720
    
    warehouseData.Add "shelfLimits", shelfLimits
    warehouseData.Add "rowLimits", rowLimits
    warehouseConfigs.Add "8", warehouseData
End Sub

Private Sub SetupWarehouse9()
    Dim warehouseData As Object
    Set warehouseData = CreateObject("Scripting.Dictionary")
    
    warehouseData.Add "number", "9"
    warehouseData.Add "isActive", True
    
    Dim shelfLimits As Object
    Dim rowLimits As Object
    Set shelfLimits = CreateObject("Scripting.Dictionary")
    Set rowLimits = CreateObject("Scripting.Dictionary")
    
    ' АНГАР 9 - Лимиты как у ангаров 5,6
    shelfLimits.Add 1000, 2000
    shelfLimits.Add 960, 1000
    shelfLimits.Add 720, 1500
    shelfLimits.Add 600, 1300
    shelfLimits.Add 540, 1100
    shelfLimits.Add 288, 580
    shelfLimits.Add 240, 500
    
    rowLimits.Add 1000, 14000
    rowLimits.Add 960, 6800
    rowLimits.Add 720, 10100
    rowLimits.Add 600, 8500
    rowLimits.Add 540, 8700
    rowLimits.Add 288, 4608
    rowLimits.Add 240, 3850
    
    warehouseData.Add "shelfLimits", shelfLimits
    warehouseData.Add "rowLimits", rowLimits
    warehouseConfigs.Add "9", warehouseData
End Sub

Private Sub SetupWarehouse10()
    Dim warehouseData As Object
    Set warehouseData = CreateObject("Scripting.Dictionary")
    
    warehouseData.Add "number", "10"
    warehouseData.Add "isActive", True
    
    Dim shelfLimits As Object
    Dim rowLimits As Object
    Set shelfLimits = CreateObject("Scripting.Dictionary")
    Set rowLimits = CreateObject("Scripting.Dictionary")
    
    ' АНГАР 10 - Лимиты как у ангаров 5,6,9
    shelfLimits.Add 1000, 2000
    shelfLimits.Add 960, 1000
    shelfLimits.Add 720, 1500
    shelfLimits.Add 600, 1300
    shelfLimits.Add 540, 1100
    shelfLimits.Add 288, 580
    shelfLimits.Add 240, 500
    
    rowLimits.Add 1000, 14000
    rowLimits.Add 960, 6800
    rowLimits.Add 720, 10100
    rowLimits.Add 600, 8500
    rowLimits.Add 540, 8700
    rowLimits.Add 288, 4608
    rowLimits.Add 240, 3850
    
    warehouseData.Add "shelfLimits", shelfLimits
    warehouseData.Add "rowLimits", rowLimits
    warehouseConfigs.Add "10", warehouseData
End Sub

Private Sub SetupWarehouse11()
    Dim warehouseData As Object
    Set warehouseData = CreateObject("Scripting.Dictionary")
    
    warehouseData.Add "number", "11"
    warehouseData.Add "isActive", True
    
    Dim shelfLimits As Object
    Dim rowLimits As Object
    Set shelfLimits = CreateObject("Scripting.Dictionary")
    Set rowLimits = CreateObject("Scripting.Dictionary")
    
    ' АНГАР 11 - Лимиты как у ангаров 5,6,9,10
    shelfLimits.Add 1000, 2000
    shelfLimits.Add 960, 1000
    shelfLimits.Add 720, 1500
    shelfLimits.Add 600, 1300
    shelfLimits.Add 540, 1100
    shelfLimits.Add 288, 580
    shelfLimits.Add 240, 500
    
    rowLimits.Add 1000, 14000
    rowLimits.Add 960, 6800
    rowLimits.Add 720, 10100
    rowLimits.Add 600, 8500
    rowLimits.Add 540, 8700
    rowLimits.Add 288, 4608
    rowLimits.Add 240, 3850
    
    warehouseData.Add "shelfLimits", shelfLimits
    warehouseData.Add "rowLimits", rowLimits
    warehouseConfigs.Add "11", warehouseData
End Sub

Private Sub SetupWarehouse12()
    Dim warehouseData As Object
    Set warehouseData = CreateObject("Scripting.Dictionary")
    
    warehouseData.Add "number", "12"
    warehouseData.Add "isActive", True
    
    Dim shelfLimits As Object
    Dim rowLimits As Object
    Set shelfLimits = CreateObject("Scripting.Dictionary")
    Set rowLimits = CreateObject("Scripting.Dictionary")
    
    ' АНГАР 12 - Лимиты как у ангаров 5,6,9,10,11
    shelfLimits.Add 1000, 2000
    shelfLimits.Add 960, 1000
    shelfLimits.Add 720, 1500
    shelfLimits.Add 600, 1300
    shelfLimits.Add 540, 1100
    shelfLimits.Add 288, 580
    shelfLimits.Add 240, 500
    
    rowLimits.Add 1000, 14000
    rowLimits.Add 960, 6800
    rowLimits.Add 720, 10100
    rowLimits.Add 600, 8500
    rowLimits.Add 540, 8700
    rowLimits.Add 288, 4608
    rowLimits.Add 240, 3850
    
    warehouseData.Add "shelfLimits", shelfLimits
    warehouseData.Add "rowLimits", rowLimits
    warehouseConfigs.Add "12", warehouseData
End Sub

' ===== ПУБЛИЧНЫЕ ФУНКЦИИ ДЛЯ РАБОТЫ С ЛИМИТАМИ =====

Public Function GetShelfLimit(warehouseNumber As String, productVolume As Integer) As Double
    ' Получает лимит стеллажа для конкретного ангара и объема препарата
    On Error GoTo ErrorHandler
    
    If Not warehouseConfigs.exists(warehouseNumber) Then
        GetShelfLimit = 1000  ' Лимит по умолчанию
        Exit Function
    End If
    
    Dim warehouseData As Object
    Set warehouseData = warehouseConfigs(warehouseNumber)
    
    Dim shelfLimits As Object
    Set shelfLimits = warehouseData("shelfLimits")
    
    If shelfLimits.exists(productVolume) Then
        GetShelfLimit = CDbl(shelfLimits(productVolume))
    Else
        GetShelfLimit = 1000  ' Лимит по умолчанию
    End If
    
    Exit Function
    
ErrorHandler:
    GetShelfLimit = 1000
End Function

Public Function GetRowLimit(warehouseNumber As String, productVolume As Integer) As Double
    ' Получает лимит ряда для конкретного ангара и объема препарата
    On Error GoTo ErrorHandler
    
    If Not warehouseConfigs.exists(warehouseNumber) Then
        GetRowLimit = 7000  ' Лимит по умолчанию
        Exit Function
    End If
    
    Dim warehouseData As Object
    Set warehouseData = warehouseConfigs(warehouseNumber)
    
    Dim rowLimits As Object
    Set rowLimits = warehouseData("rowLimits")
    
    If rowLimits.exists(productVolume) Then
        GetRowLimit = CDbl(rowLimits(productVolume))
    Else
        GetRowLimit = 7000  ' Лимит по умолчанию
    End If
    
    Exit Function
    
ErrorHandler:
    GetRowLimit = 7000
End Function

' ===== ФУНКЦИИ УПРАВЛЕНИЯ АКТИВНОСТЬЮ АНГАРОВ =====

Public Sub SetWarehouseActive(warehouseNumber As String, isActive As Boolean)
    ' Устанавливает активность ангара
    On Error Resume Next
    
    If warehouseActiveStates.exists(warehouseNumber) Then
        warehouseActiveStates.Remove warehouseNumber
    End If
    
    warehouseActiveStates.Add warehouseNumber, isActive
    
    ' Обновляем конфигурацию ангара
    If warehouseConfigs.exists(warehouseNumber) Then
        Dim warehouseData As Object
        Set warehouseData = warehouseConfigs(warehouseNumber)
        
        If warehouseData.exists("isActive") Then
            warehouseData.Remove "isActive"
        End If
        warehouseData.Add "isActive", isActive
    End If
End Sub

Public Function IsWarehouseActive(warehouseNumber As String) As Boolean
    ' Проверяет, активен ли ангар
    On Error Resume Next
    
    If warehouseActiveStates.exists(warehouseNumber) Then
        IsWarehouseActive = CBool(warehouseActiveStates(warehouseNumber))
    Else
        IsWarehouseActive = True  ' По умолчанию активен
    End If
End Function

Public Function GetActiveWarehouseCount() As Integer
    ' Возвращает количество активных ангаров
    Dim count As Integer
    count = 0
    
    Dim i As Integer
    For i = 5 To 12
        If IsWarehouseActive(CStr(i)) Then
            count = count + 1
        End If
    Next i
    
    GetActiveWarehouseCount = count
End Function

Public Function GetActiveWarehousesList() As String
    ' Возвращает список активных ангаров
    Dim list As String
    list = ""
    
    Dim i As Integer
    For i = 5 To 12
        If IsWarehouseActive(CStr(i)) Then
            list = list & CStr(i) & ", "
        End If
    Next i
    
    ' Убираем последнюю запятую
    If Len(list) > 2 Then
        list = Left(list, Len(list) - 2)
    End If
    
    GetActiveWarehousesList = list
End Function

Public Function GetActiveWarehousesArray() As Variant
    ' Возвращает массив активных ангаров
    Dim activeWarehouses() As String
    Dim count As Integer
    count = 0
    
    ' Считаем активные ангары
    Dim i As Integer
    For i = 5 To 12
        If IsWarehouseActive(CStr(i)) Then
            count = count + 1
        End If
    Next i
    
    ' Заполняем массив
    If count > 0 Then
        ReDim activeWarehouses(count - 1)
        Dim index As Integer
        index = 0
        
        For i = 5 To 12
            If IsWarehouseActive(CStr(i)) Then
                activeWarehouses(index) = CStr(i)
                index = index + 1
            End If
        Next i
        
        GetActiveWarehousesArray = activeWarehouses
    Else
        GetActiveWarehousesArray = Array()  ' Пустой массив
    End If
End Function

' ===== ИНФОРМАЦИОННЫЕ ФУНКЦИИ =====

Public Function GetWarehouseCapacityInfo(warehouseNumber As String) As String
    ' Возвращает информацию о лимитах конкретного ангара
    Dim info As String
    info = "=== ЛИМИТЫ АНГАРА " & warehouseNumber & " ===" & vbNewLine
    
    If Not warehouseConfigs.exists(warehouseNumber) Then
        info = info & "Ангар не найден в конфигурации!"
        GetWarehouseCapacityInfo = info
        Exit Function
    End If
    
    Dim warehouseData As Object
    Set warehouseData = warehouseConfigs(warehouseNumber)
    
    info = info & "Статус: " & IIf(CBool(warehouseData("isActive")), "АКТИВЕН", "НЕАКТИВЕН") & vbNewLine
    info = info & vbNewLine & "ЛИМИТЫ НА СТЕЛЛАЖ:" & vbNewLine
    
    Dim shelfLimits As Object
    Set shelfLimits = warehouseData("shelfLimits")
    
    Dim key As Variant
    For Each key In shelfLimits.keys
        info = info & "• Объем " & key & "мл: " & shelfLimits(key) & vbNewLine
    Next key
    
    info = info & vbNewLine & "ЛИМИТЫ НА РЯД:" & vbNewLine
    
    Dim rowLimits As Object
    Set rowLimits = warehouseData("rowLimits")
    
    For Each key In rowLimits.keys
        info = info & "• Объем " & key & "мл: " & rowLimits(key) & vbNewLine
    Next key
    
    GetWarehouseCapacityInfo = info
End Function

Public Function GetSystemCapacityStats() As String
    ' Возвращает общую статистику системы лимитов
    Dim stats As String
    stats = "=== СТАТИСТИКА СИСТЕМЫ ЛИМИТОВ ===" & vbNewLine & vbNewLine
    
    stats = stats & "Всего ангаров в системе: 8 (с 5 по 12)" & vbNewLine
    stats = stats & "Активных ангаров: " & GetActiveWarehouseCount() & vbNewLine
    stats = stats & "Список активных: " & GetActiveWarehousesList() & vbNewLine
    stats = stats & "Групп объемов препаратов: 7 (1000, 960, 720, 600, 540, 288, 240мл)" & vbNewLine
    
    GetSystemCapacityStats = stats
End Function
