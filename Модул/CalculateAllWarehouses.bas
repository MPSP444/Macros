Attribute VB_Name = "CalculateAllWarehouses"
' А затем используем эту конфигурацию для подсчета итогов:
Option Explicit


' Глобальные переменные
Private dict As Object
Private ws As Worksheet


Private Function IsRowEmpty(ByVal row As Long, ByVal col As Long) As Boolean
    On Error Resume Next
    Dim cellValue As Variant
    cellValue = ws.Cells(row, col).value
    IsRowEmpty = (isEmpty(cellValue) Or cellValue = "" Or cellValue = 0)
    On Error GoTo 0
End Function
Private Function ShouldProcessRow(ByVal row As Long, ByVal headerRow As Long, ByVal RowNumbersToSkip As Variant) As Boolean
    On Error Resume Next
    Dim skipRow As Variant
    
    ' Проверяем, не является ли строка строкой с номерами рядов
    For Each skipRow In RowNumbersToSkip
        If row = skipRow Then
            ShouldProcessRow = False
            Exit Function
        End If
    Next skipRow
    
    ' Проверяем, не является ли строка пустой
    Dim lastCol As Long
    lastCol = ws.Cells(headerRow, ws.columns.count).End(xlToLeft).Column
    
    Dim isEmpty As Boolean
    isEmpty = True
    Dim col As Long
    For col = 1 To lastCol
        If Not IsRowEmpty(row, col) Then
            isEmpty = False
            Exit For
        End If
    Next col
    
    ShouldProcessRow = Not isEmpty
    On Error GoTo 0
End Function

Private Sub AddProductToDict(ByVal ProductName As String)
    On Error Resume Next
    If Len(Trim(ProductName)) > 0 Then
        If Not dict.exists(ProductName) Then
            dict.Add ProductName, 0
        End If
    End If
    On Error GoTo 0
End Sub

Private Function GetMergedCellValue(ByVal cell As Range) As String
    On Error Resume Next
    If cell.MergeCells Then
        GetMergedCellValue = CStr(cell.MergeArea.Cells(1, 1).value)
    Else
        GetMergedCellValue = CStr(cell.value)
    End If
    If IsNull(GetMergedCellValue) Then GetMergedCellValue = ""
    On Error GoTo 0
End Function

Private Function GetValueFromCell(ByVal cell As Range) As Double
    On Error Resume Next
    Dim cellValue As Variant
    
    ' Если ячейка объединена, берем значение из первой ячейки объединенной области
    If cell.MergeCells Then
        cellValue = cell.MergeArea.Cells(1, 1).value
    Else
        cellValue = cell.value
    End If
    
    ' Проверяем, является ли значение числовым
    If IsNumeric(cellValue) Then
        GetValueFromCell = CDbl(cellValue)
    Else
        GetValueFromCell = 0
    End If
    
    On Error GoTo 0
End Function

Private Sub ProcessSection(ByVal headerRow As Long, ByVal dataStartRow As Long, ByVal dataEndRow As Long, ByVal RowNumbersToSkip As Variant)
    On Error Resume Next
    Dim col As Long
    Dim row As Long
    Dim ProductName As Variant
    Dim cellValue As Double
    Dim lastCol As Long
    Dim currentCell As Range
    Dim processedRanges As Object
    
    Set processedRanges = CreateObject("Scripting.Dictionary")
    lastCol = ws.Cells(headerRow, ws.columns.count).End(xlToLeft).Column
    
    ' Собираем названия препаратов из заголовка
    For col = 1 To lastCol
        ProductName = Trim(GetMergedCellValue(ws.Cells(headerRow, col)))
        Call AddProductToDict(ProductName)
    Next col
    
    ' Суммируем значения для каждого препарата
    For Each ProductName In dict.keys
        For col = 1 To lastCol
            If Trim(GetMergedCellValue(ws.Cells(headerRow, col))) = ProductName Then
                row = dataStartRow
                
                While row <= dataEndRow
                    If ShouldProcessRow(row, headerRow, RowNumbersToSkip) Then
                        Set currentCell = ws.Cells(row, col)
                        
                        ' Создаем уникальный ключ для объединенной области
                        Dim rangeKey As String
                        If currentCell.MergeCells Then
                            rangeKey = currentCell.MergeArea.address
                        Else
                            rangeKey = currentCell.address
                        End If
                        
                        ' Проверяем, не обработали ли мы уже эту область
                        If Not processedRanges.exists(rangeKey) Then
                            cellValue = GetValueFromCell(currentCell)
                            If cellValue > 0 Then
                                dict(ProductName) = dict(ProductName) + cellValue
                                processedRanges.Add rangeKey, True
                            End If
                        End If
                        
                        ' Переходим к следующей строке
                        If currentCell.MergeCells Then
                            row = row + currentCell.MergeArea.rows.count
                        Else
                            row = row + 1
                        End If
                    Else
                        row = row + 1
                    End If
                Wend
            End If
        Next col
    Next ProductName
    
    Set processedRanges = Nothing
    On Error GoTo 0
End Sub

Private Function FindProductRow(ByVal ProductName As String, ByVal startRow As Long) As Long
    On Error Resume Next
    Dim row As Long
    row = startRow
    
    While Not isEmpty(ws.Cells(row, 1))
        If CStr(ws.Cells(row, 1).value) = ProductName Then
            FindProductRow = row
            Exit Function
        End If
        row = row + 1
    Wend
    
    FindProductRow = row ' Возвращаем первую пустую строку, если товар не найден
    On Error GoTo 0
End Function

Public Sub CalculateProductTotalsForWarehouse(config As ModuleTypes.WarehouseConfig)
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    Dim outputStartRow As Long
    Dim currentRow As Long
    Dim key As Variant
    Dim totalSum As Double
    totalSum = 0
    
    ' Инициализация
    Set ws = ActiveWorkbook.Worksheets(config.sheetName)
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Обрабатываем секции
    Call ProcessSection(config.UpperHeaderRow, config.UpperDataStart, config.UpperDataEnd, config.RowNumbersToSkip)
    Call ProcessSection(config.LowerHeaderRow, config.LowerDataStart, config.LowerDataEnd, config.RowNumbersToSkip)
    
    ' Очищаем предыдущие результаты
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If lastRow >= config.outputStartRow Then
        ws.Range(ws.Cells(config.outputStartRow, 1), ws.Cells(lastRow, 2)).ClearContents
    End If
    
    ' Добавляем заголовки
    ws.Cells(config.outputStartRow, 1).value = "Наименование товара"
    ws.Cells(config.outputStartRow, 2).value = "Количество"
    
    ' Выводим результаты
    For Each key In dict.keys
        If dict(key) > 0 Then
            currentRow = FindProductRow(CStr(key), config.outputStartRow + 1)
            
            ' Записываем наименование товара
            If isEmpty(ws.Cells(currentRow, 1)) Then
                ws.Cells(currentRow, 1).value = key
            End If
            
            ' Записываем количество
            ws.Cells(currentRow, 2).value = dict(key)
            
            ' Добавляем к общей сумме
            totalSum = totalSum + dict(key)
        End If
    Next key
    
    ' Находим последнюю заполненную строку
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    
    ' Добавляем итоговую строку
    lastRow = lastRow + 1
    ws.Cells(lastRow, 1).value = "ИТОГО:"
    ws.Cells(lastRow, 2).value = totalSum
    
    ' Форматирование всей таблицы
    With ws.Range(ws.Cells(config.outputStartRow, 1), ws.Cells(lastRow, 2))
        .Borders.Weight = xlThin
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .columns.AutoFit
    End With
    
    ' Форматирование заголовков
    With ws.Range(ws.Cells(config.outputStartRow, 1), ws.Cells(config.outputStartRow, 2))
        .Interior.Color = RGB(200, 200, 200)
        .Font.Bold = True
    End With
    
    ' Форматирование итоговой строки
    With ws.Range(ws.Cells(lastRow, 1), ws.Cells(lastRow, 2))
        .Interior.Color = RGB(255, 255, 0)
        .Font.Bold = True
        .Font.Size = 12
    End With

ExitSub:
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "Произошла ошибка в ангаре " & config.sheetName & ": " & Err.description, vbCritical
    Resume ExitSub
End Sub
Public Sub CalculateAllWarehouses()
    If ActiveWorkbook Is Nothing Then
        MsgBox "Нет открытой книги Excel!", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' Обрабатываем все ангары
    Dim i As Integer
    Dim config As ModuleTypes.WarehouseConfig
    
    For i = 5 To 12
        config = ModuleConfig.GetWarehouseConfig(CStr(i))  ' Добавили ModuleConfig.
        Call CalculateProductTotalsForWarehouse(config)
    Next i
    
    Application.ScreenUpdating = True
    MsgBox "Расчеты для всех ангаров завершены!", vbInformation
End Sub
