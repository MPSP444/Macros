Attribute VB_Name = "ModuleTypes"
' ===== ModuleTypes - ИСПРАВЛЕННАЯ ВЕРСИЯ БЕЗ КОНФЛИКТОВ =====
Option Explicit

' ===== ОСНОВНЫЕ ТИПЫ ДАННЫХ СИСТЕМЫ =====

' Конфигурация ангара (централизованное определение)
Public Type WarehouseConfig
    sheetName As String
    UpperDataStart As Long
    UpperDataEnd As Long
    UpperHeaderRow As Long
    LowerDataStart As Long
    LowerDataEnd As Long
    LowerHeaderRow As Long
    outputStartRow As Long
    RowNumbersToSkip As Variant
End Type

' Операция в ангаре
Public Type WarehouseOperation
    warehouse As String  ' Номер ангара
    row As String       ' Номер ряда
    Letter As String    ' Буква ячейки
    level As String     ' Уровень (ярус)
    value As Double     ' Значение
    Batch As String     ' Номер партии
    ProductName As String ' Название товара (добавлено для совместимости)
End Type

' Местоположение в ангаре
Public Type WarehouseLocation
    warehouse As String  ' Номер ангара
    row As String       ' Номер ряда
    Letter As String    ' Буква ячейки
    level As String     ' Уровень (ярус)
End Type

' Структуры для оптимизации заголовков
Public Type HeaderGroup
    ProductName As String      ' Название товара
    startRow As Integer       ' Начальный ряд группы
    endRow As Integer         ' Конечный ряд группы
    StartColumn As Integer    ' Начальный столбец (C, E, G...)
    EndColumn As Integer      ' Конечный столбец (D, F, H...)
    section As String         ' Секция (UPPER/LOWER)
    IsConflict As Boolean     ' Флаг конфликта товаров
    RowsInGroup As String     ' Список рядов для отладки (например: "3,5,6")
End Type

' Структура для анализа существующих заголовков
Public Type ExistingHeader
    ProductName As String     ' Название товара
    row As Integer           ' Номер ряда
    StartColumn As Integer   ' Начальный столбец
    EndColumn As Integer     ' Конечный столбец
    section As String        ' Секция
    IsConflict As Boolean    ' Есть ли конфликт (содержит "+")
End Type

' ===== ДОПОЛНИТЕЛЬНЫЕ ТИПЫ ДЛЯ РАСШИРЕННОГО ФУНКЦИОНАЛА =====

' Информация о товаре (для модуля ProductInfo)
Public Type ProductInfo
    ProductName As String
    volumeGroup As String    ' Группа объема (1-7)
    description As String
    standardQuantity As Double
End Type

' Информация о размещении (для умного размещения)
Public Type PlacementResult
    warehouse As String      ' Номер ангара
    row As String           ' Номер ряда
    Shelf As String         ' Буква стеллажа
    level As String         ' Уровень
    ProductName As String   ' Название товара
    Batch As String         ' Партия
    quantity As Double      ' Количество
    IsPlaced As Boolean     ' Успешно размещено?
    ErrorMessage As String  ' Сообщение об ошибке
    cellAddress As String   ' Адрес ячейки (например: "D15")
End Type

' Конфигурация лимитов ангара
Public Type CapacityConfig
    warehouseNumber As String
    maxItemsPerShelf As Integer    ' Максимум товаров на стеллаж
    maxQuantityPerRow As Double    ' Максимум количества в ряду
    isActive As Boolean           ' Активен ли ангар
    Notes As String              ' Дополнительные заметки
End Type

' Результат валидации
Public Type ValidationResult
    IsValid As Boolean           ' Прошла ли валидация
    ErrorMessage As String       ' Сообщение об ошибке
    WarningMessage As String     ' Предупреждение
    ValidatedData As String      ' Проверенные данные
End Type

' Статистика ангара
Public Type WarehouseStats
    warehouseNumber As String
    totalItems As Integer        ' Общее количество позиций
    totalQuantity As Double      ' Общее количество товара
    occupiedShelves As Integer   ' Занятые стеллажи
    emptyPositions As Integer    ' Свободные позиции
    lastUpdateTime As Date       ' Время последнего обновления
End Type

' ===== ФУНКЦИИ ДЛЯ РАБОТЫ С ТИПАМИ =====

' Создает пустую операцию ангара
Public Function CreateEmptyOperation() As WarehouseOperation
    Dim emptyOp As WarehouseOperation
    emptyOp.warehouse = ""
    emptyOp.row = ""
    emptyOp.Letter = ""
    emptyOp.level = ""
    emptyOp.value = 0
    emptyOp.Batch = ""
    emptyOp.ProductName = ""
    CreateEmptyOperation = emptyOp
End Function

' Создает пустой результат размещения
Public Function CreateEmptyPlacement() As PlacementResult
    Dim emptyPlace As PlacementResult
    emptyPlace.warehouse = ""
    emptyPlace.row = ""
    emptyPlace.Shelf = ""
    emptyPlace.level = ""
    emptyPlace.ProductName = ""
    emptyPlace.Batch = ""
    emptyPlace.quantity = 0
    emptyPlace.IsPlaced = False
    emptyPlace.ErrorMessage = ""
    emptyPlace.cellAddress = ""
    CreateEmptyPlacement = emptyPlace
End Function

' Создает пустую конфигурацию ангара
Public Function CreateEmptyConfig() As WarehouseConfig
    Dim emptyConfig As WarehouseConfig
    emptyConfig.sheetName = ""
    emptyConfig.UpperDataStart = 0
    emptyConfig.UpperDataEnd = 0
    emptyConfig.UpperHeaderRow = 0
    emptyConfig.LowerDataStart = 0
    emptyConfig.LowerDataEnd = 0
    emptyConfig.LowerHeaderRow = 0
    emptyConfig.outputStartRow = 0
    emptyConfig.RowNumbersToSkip = Array()
    CreateEmptyConfig = emptyConfig
End Function

' Проверяет, является ли операция валидной
Public Function IsValidOperation(op As WarehouseOperation) As Boolean
    IsValidOperation = (op.warehouse <> "" And op.row <> "" And _
                       op.Letter <> "" And op.level <> "" And _
                       op.value > 0 And op.Batch <> "")
End Function

' Форматирует операцию в читаемый вид
Public Function FormatOperation(op As WarehouseOperation) As String
    FormatOperation = "Ангар " & op.warehouse & ", ряд " & op.row & _
                     ", стеллаж " & op.Letter & ", уровень " & op.level & _
                     " - " & op.ProductName & " (партия: " & op.Batch & _
                     ", количество: " & op.value & ")"
End Function

' ===== КОНСТАНТЫ СИСТЕМЫ =====

' Поддерживаемые ангары
Public Const MIN_WAREHOUSE As Integer = 5
Public Const MAX_WAREHOUSE As Integer = 12

' Максимальные уровни стеллажей
Public Const MAX_SHELF_LEVEL As Integer = 3

' Стандартные секции ангара
Public Const UPPER_SECTION As String = "UPPER"
Public Const LOWER_SECTION As String = "LOWER"

' Стандартные разделители
Public Const MAIN_DELIMITER As String = " - "
Public Const LOCATION_DELIMITER As String = "-"

' ===== ФУНКЦИИ ВАЛИДАЦИИ ТИПОВ =====

Public Function IsValidWarehouseNumber(warehouse As String) As Boolean
    Dim num As Integer
    On Error Resume Next
    num = CInt(warehouse)
    On Error GoTo 0
    IsValidWarehouseNumber = (num >= MIN_WAREHOUSE And num <= MAX_WAREHOUSE)
End Function

Public Function IsValidShelfLevel(level As String) As Boolean
    Dim num As Integer
    On Error Resume Next
    num = CInt(level)
    On Error GoTo 0
    IsValidShelfLevel = (num >= 1 And num <= MAX_SHELF_LEVEL)
End Function

Public Function IsValidBatchFormat(Batch As String) As Boolean
    ' Партия должна начинаться с "пар" и содержать числа
    IsValidBatchFormat = (Len(Batch) >= 4 And Left(LCase(Batch), 3) = "пар")
End Function

