Attribute VB_Name = "ModuleSpecialCases"

' === ModuleSpecialCases ===
Option Explicit

Private Type SpecialRowConfig
    warehouseNumber As String
    sheetName As String
    UpperDataStart As Long
    UpperDataEnd As Long
    UpperHeaderRow As Long
    LowerDataStart As Long   ' Только для Ангара 5
    LowerDataEnd As Long     ' Только для Ангара 5
    LowerHeaderRow As Long   ' Только для Ангара 5
    outputStartRow As Long
    Format As Integer        ' 1 = вертикальный, 2 = горизонтальный
    DataColumn As Long       ' Колонка с названиями
    valueColumn As Long      ' Колонка со значениями
End Type

Private Function GetSpecialConfig(warehouseNumber As String) As SpecialRowConfig
    Dim config As SpecialRowConfig
    
    Select Case warehouseNumber
        Case "5"
            With config
                .warehouseNumber = "5"
                .sheetName = "Ангар 5"
                ' Первая секция
                .UpperDataStart = 4
                .UpperDataEnd = 30
                .UpperHeaderRow = 2
                ' Вторая секция
                .LowerDataStart = 42
                .LowerDataEnd = 68
                .LowerHeaderRow = 70
                ' Настройки вывода
                .outputStartRow = 73
                .Format = 2
                .DataColumn = 3
                .valueColumn = 4
            End With
            
        Case "8"
            With config
                .warehouseNumber = "8"
                .sheetName = "Ангар 8"
                ' Только первая секция
                .UpperDataStart = 4
                .UpperDataEnd = 24
                .UpperHeaderRow = 2
                ' Настройки вывода
                .outputStartRow = 59
                .Format = 2
                .DataColumn = 3
                .valueColumn = 4
            End With
            
        Case "12"
            With config
                .warehouseNumber = "12"
                .sheetName = "Ангар 12"
                ' Только первая секция
                .UpperDataStart = 4
                .UpperDataEnd = 27
                .UpperHeaderRow = 2
                ' Настройки вывода
                .outputStartRow = 70
                .Format = 2
                .DataColumn = 3
                .valueColumn = 4
            End With
    End Select
    
    GetSpecialConfig = config
End Function

Public Sub ProcessSpecialRows()
    On Error GoTo ErrorHandler
    
    If ActiveWorkbook Is Nothing Then
        MsgBox "Нет открытой книги Excel!", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' ОТЛАДКА: проверяем наличие листов
    Debug.Print "=== НАЧАЛО ОБРАБОТКИ СПЕЦИАЛЬНЫХ РЯДОВ ==="
    If Not CheckWorksheetExists("Ангар 5") Then GoTo ErrorHandler
    If Not CheckWorksheetExists("Ангар 8") Then GoTo ErrorHandler
    If Not CheckWorksheetExists("Ангар 12") Then GoTo ErrorHandler
    
    ' Обработка Ангара 5 (обе секции)
    Debug.Print "Начинаем обработку Ангара 5"
    ProcessWarehouse "5"
    
    ' Обработка Ангара 8 (только первая секция)
    Debug.Print "Начинаем обработку Ангара 8"
    ProcessWarehouse "8"
    
    ' Обработка Ангара 12
    Debug.Print "Начинаем обработку Ангара 12"
    ProcessWarehouse "12"
    
    Application.ScreenUpdating = True
    MsgBox "Обработка специальных рядов завершена!", vbInformation
    Debug.Print "=== ОБРАБОТКА ЗАВЕРШЕНА УСПЕШНО ==="
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Произошла ошибка: " & Err.description & vbNewLine & _
           "Номер ошибки: " & Err.Number, vbCritical
    Debug.Print "ОШИБКА: " & Err.description & " (№" & Err.Number & ")"
End Sub

Private Function CheckWorksheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(sheetName)
    If Err.Number <> 0 Then
        MsgBox "Лист '" & sheetName & "' не найден в активной книге!", vbExclamation
        CheckWorksheetExists = False
        Err.Clear
    Else
        Debug.Print "Лист '" & sheetName & "' найден успешно"
        CheckWorksheetExists = True
    End If
    On Error GoTo 0
End Function

Private Function IsValidProductName(ProductName As String) As Boolean
    ' Проверяем, не является ли строка пустой
    If Trim(ProductName) = "" Then
        IsValidProductName = False
        Exit Function
    End If
    
    ' Проверяем, не является ли строка числом
    If IsNumeric(ProductName) Then
        IsValidProductName = False
        Exit Function
    End If
    
    ' Проверяем, не начинается ли строка с "пар"
    If Left(LCase(Trim(ProductName)), 3) = "пар" Then
        IsValidProductName = False
        Exit Function
    End If
    
    ' Проверяем на служебные слова/символы
    If InStr(1, LCase(ProductName), "итого") > 0 Or _
       InStr(1, ProductName, "/") > 0 Or _
       InStr(1, ProductName, "_") > 0 Or _
       Len(Trim(ProductName)) <= 1 Then  ' Убрали проверку на дефис
        IsValidProductName = False
        Exit Function
    End If
    
    IsValidProductName = True
End Function

Private Function IsSameProduct(product1 As String, product2 As String) As Boolean
    ' Удаляем лишние пробелы
    product1 = Trim(product1)
    product2 = Trim(product2)
    
    ' Если строки полностью совпадают
    If product1 = product2 Then
        IsSameProduct = True
        Exit Function
    End If
    
    ' Если хотя бы одна строка пустая
    If product1 = "" Or product2 = "" Then
        IsSameProduct = False
        Exit Function
    End If
    
    ' Приводим к нижнему регистру для сравнения
    product1 = LCase(product1)
    product2 = LCase(product2)
    
    IsSameProduct = (product1 = product2)
End Function

Private Function SafeGetCellValue(ws As Worksheet, rowNum As Long, colNum As Long) As Variant
    ' Безопасное получение значения ячейки с проверкой границ
    On Error Resume Next
    
    ' Проверяем границы листа
    If rowNum < 1 Or rowNum > ws.rows.count Or colNum < 1 Or colNum > ws.columns.count Then
        SafeGetCellValue = ""
        Exit Function
    End If
    
    SafeGetCellValue = ws.Cells(rowNum, colNum).value
    If Err.Number <> 0 Then
        SafeGetCellValue = ""
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Function SafeCheckMerged(ws As Worksheet, rowNum As Long, colNum As Long) As Boolean
    ' Безопасная проверка объединения ячеек
    On Error Resume Next
    
    ' Проверяем границы листа
    If rowNum < 1 Or rowNum > ws.rows.count Or colNum < 1 Or colNum > ws.columns.count Then
        SafeCheckMerged = False
        Exit Function
    End If
    
    SafeCheckMerged = ws.Cells(rowNum, colNum).MergeCells
    If Err.Number <> 0 Then
        SafeCheckMerged = False
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Sub ProcessSection(ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, _
                         config As SpecialRowConfig, ByRef resultsDict As Object)
    Debug.Print "Обработка секции в " & ws.Name & " строки " & startRow & " - " & endRow
    
    Dim rowNum As Long, colNum As Long
    
    ' БЕЗОПАСНОЕ ОПРЕДЕЛЕНИЕ МАКСИМАЛЬНОГО СТОЛБЦА
    Dim maxCol As Long
    On Error Resume Next
    maxCol = ws.UsedRange.columns.count + ws.UsedRange.Column - 1
    If Err.Number <> 0 Or maxCol > 100 Then
        maxCol = 50  ' Безопасный максимум
        Err.Clear
    End If
    If maxCol < 10 Then maxCol = 30    ' Минимальный разумный диапазон
    On Error GoTo 0
    
    Debug.Print "Максимальный столбец для обработки: " & maxCol
    
    ' Проходим по всем строкам
    For rowNum = startRow To endRow
        ' Проходим по столбцам с БЕЗОПАСНЫМИ ГРАНИЦАМИ
        For colNum = 3 To maxCol - 1  ' -1 чтобы оставить место для проверки справа
            Dim ProductName As String
            Dim totalValue As Double
            totalValue = 0
            
            ' БЕЗОПАСНОЕ ПОЛУЧЕНИЕ ЗНАЧЕНИЯ ЯЧЕЙКИ
            ProductName = Trim(CStr(SafeGetCellValue(ws, rowNum, colNum)))
            
            If ProductName <> "" And IsValidProductName(ProductName) Then
                Debug.Print "Найден товар: " & ProductName & " в ячейке " & _
                           ws.Cells(rowNum, colNum).address(False, False)
                
                ' БЕЗОПАСНАЯ ПРОВЕРКА ЯЧЕЙКИ СПРАВА
                Dim rightCellValue As Variant
                Dim isRightMerged As Boolean
                
                rightCellValue = SafeGetCellValue(ws, rowNum, colNum + 1)
                isRightMerged = SafeCheckMerged(ws, rowNum, colNum + 1)
                
                ' 1. Сначала проверяем значение справа
                If Not isRightMerged And Not isEmpty(rightCellValue) Then
                    ' Если ячейка справа не объединена и не пуста
                    If IsNumeric(rightCellValue) Then
                        totalValue = CDbl(rightCellValue)
                        Debug.Print "Добавлено значение справа: " & totalValue
                    End If
                Else
                    ' 2. Если справа пусто или ячейка объединена, ищем снизу
                    Debug.Print "Ячейка справа пуста или объединена, ищем снизу"
                    Dim lookDownRow As Long
                    For lookDownRow = 1 To 4
                        Dim currentRow As Long
                        currentRow = rowNum + lookDownRow
                        
                        ' Проверяем, не вышли ли за пределы секции
                        If currentRow > endRow Then Exit For
                        
                        ' БЕЗОПАСНАЯ ПРОВЕРКА ЯЧЕЙКИ СНИЗУ
                        Dim cellBelowValue As Variant
                        cellBelowValue = SafeGetCellValue(ws, currentRow, colNum)
                        
                        If Not isEmpty(cellBelowValue) Then
                            Dim cellBelowStr As String
                            cellBelowStr = CStr(cellBelowValue)
                            
                            ' Если нашли новое наименование товара, прекращаем поиск
                            If IsValidProductName(cellBelowStr) And _
                               Not IsSameProduct(cellBelowStr, ProductName) Then
                                Debug.Print "Найдено другое наименование: " & cellBelowStr
                                Exit For
                            End If
                            
                            ' Проверяем на числовое значение
                            If IsNumeric(cellBelowValue) Then
                                totalValue = totalValue + CDbl(cellBelowValue)
                                Debug.Print "Добавлено значение снизу: " & cellBelowValue
                            End If
                        End If
                    Next lookDownRow
                End If
                
                ' Если нашли какие-то значения, добавляем их в словарь
                If totalValue > 0 Then
                    On Error Resume Next
                    If resultsDict.exists(ProductName) Then
                        ' Суммируем новое значение с уже существующим
                        resultsDict(ProductName) = resultsDict(ProductName) + totalValue
                        Debug.Print "Обновлено значение для " & ProductName & ": " & resultsDict(ProductName) & _
                                   " (добавлено " & totalValue & ")"
                    Else
                        ' Добавляем новый товар
                        resultsDict.Add ProductName, totalValue
                        Debug.Print "Добавлен новый товар " & ProductName & ": " & totalValue
                    End If
                    If Err.Number <> 0 Then
                        Debug.Print "Ошибка при работе со словарем: " & Err.description
                        Err.Clear
                    End If
                    On Error GoTo 0
                End If
            End If
        Next colNum
    Next rowNum
    
    Debug.Print "Завершена обработка секции. Найдено товаров: " & resultsDict.count
End Sub

Public Sub ProcessWarehouse(ByVal warehouseNumber As String)
    On Error GoTo ErrorHandler
    
    Debug.Print "Начало обработки ангара " & warehouseNumber
    
    Dim config As SpecialRowConfig
    config = GetSpecialConfig(warehouseNumber)
    
    ' Проверяем наличие листа
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(config.sheetName)
    If Err.Number <> 0 Then
        MsgBox "Лист '" & config.sheetName & "' не найден!", vbExclamation
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    ' Словарь для накопления результатов
    Dim resultsDict As Object
    Set resultsDict = CreateObject("Scripting.Dictionary")
    
    ' Обработка первой секции
    Debug.Print "Обработка первой секции: " & config.UpperDataStart & "-" & config.UpperDataEnd
    ProcessSection ws, config.UpperDataStart, config.UpperDataEnd, config, resultsDict
    
    ' Обработка второй секции (только для Ангара 5)
    If warehouseNumber = "5" Then
        Debug.Print "Обработка второй секции: " & config.LowerDataStart & "-" & config.LowerDataEnd
        ProcessSection ws, config.LowerDataStart, config.LowerDataEnd, config, resultsDict
    End If
    
    ' Вывод результатов
    OutputResults ws, resultsDict, config.outputStartRow
    Debug.Print "Завершена обработка ангара " & warehouseNumber
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при обработке ангара " & warehouseNumber & ": " & Err.description, vbCritical
    Debug.Print "ОШИБКА в ангаре " & warehouseNumber & ": " & Err.description
End Sub

Private Sub OutputResults(ws As Worksheet, ByRef resultsDict As Object, ByVal outputRow As Long)
    On Error GoTo ErrorHandler
    
    ' Очищаем область вывода
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If lastRow >= outputRow Then
        ws.Range(ws.Cells(outputRow, 1), ws.Cells(lastRow, 2)).ClearContents
    End If
    
    ' Создаем новый словарь для объединенных результатов
    Dim finalDict As Object
    Set finalDict = CreateObject("Scripting.Dictionary")
    
    ' Суммируем значения для одинаковых препаратов
    Dim ProductName As Variant
    For Each ProductName In resultsDict.keys
        Dim cleanName As String
        cleanName = UCase(Trim(ProductName))  ' Преобразуем в верхний регистр
        
        If finalDict.exists(cleanName) Then
            finalDict(cleanName) = finalDict(cleanName) + resultsDict(ProductName)
        Else
            finalDict.Add cleanName, resultsDict(ProductName)
        End If
    Next ProductName
    
    ' Создаем массив для сортировки
    Dim sortArray() As Variant
    If finalDict.count > 0 Then
        ReDim sortArray(0 To finalDict.count - 1, 0 To 1)
        
        Dim i As Long
        i = 0
        For Each ProductName In finalDict.keys
            sortArray(i, 0) = StrConv(ProductName, vbProperCase)  ' Преобразуем первую букву в заглавную
            sortArray(i, 1) = finalDict(ProductName)
            i = i + 1
        Next ProductName
        
        ' Сортируем массив по названиям (первый столбец)
        If UBound(sortArray, 1) > 0 Then
            QuickSort sortArray, 0, UBound(sortArray, 1)
        End If
    End If
    
    ' Заголовки
    ws.Cells(outputRow, 1).value = "Наименование товара"
    ws.Cells(outputRow, 2).value = "Количество"
    
    ' Выводим отсортированные данные
    Dim currentRow As Long
    currentRow = outputRow + 1
    
    Dim totalSum As Double
    totalSum = 0
    
    ' Выводим данные из отсортированного массива
    If finalDict.count > 0 Then
        For i = 0 To UBound(sortArray, 1)
            ws.Cells(currentRow, 1).value = sortArray(i, 0)
            ws.Cells(currentRow, 2).value = sortArray(i, 1)
            totalSum = totalSum + sortArray(i, 1)
            currentRow = currentRow + 1
        Next i
    End If
    
    ' Итоговая строка
    ws.Cells(currentRow, 1).value = "ИТОГО:"
    ws.Cells(currentRow, 2).value = totalSum
    
    ' Форматирование
    With ws.Range(ws.Cells(outputRow, 1), ws.Cells(currentRow, 2))
        .Borders.Weight = xlThin
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .columns.AutoFit
        
        ' Форматирование заголовков
        ws.Range(ws.Cells(outputRow, 1), ws.Cells(outputRow, 2)).Interior.Color = RGB(200, 200, 200)
        
        ' Форматирование итоговой строки
        With ws.Range(ws.Cells(currentRow, 1), ws.Cells(currentRow, 2))
            .Interior.Color = RGB(255, 255, 0)
            .Font.Size = 12
        End With
    End With
    
    Debug.Print "Результаты выведены. Всего товаров: " & finalDict.count & ", общая сумма: " & totalSum
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при выводе результатов: " & Err.description, vbCritical
    Debug.Print "ОШИБКА при выводе результатов: " & Err.description
End Sub

' Добавляем процедуры для сортировки
Private Sub QuickSort(arr As Variant, ByVal low As Long, ByVal high As Long)
    On Error Resume Next
    
    Dim pivot As Variant
    Dim tmp As Variant
    Dim i As Long
    Dim j As Long
    
    If low < high Then
        i = low
        j = high
        pivot = arr((low + high) \ 2, 0)
        
        Do While i <= j
            Do While StrComp(arr(i, 0), pivot, vbTextCompare) < 0 And i < high
                i = i + 1
            Loop
            
            Do While StrComp(arr(j, 0), pivot, vbTextCompare) > 0 And j > low
                j = j - 1
            Loop
            
            If i <= j Then
                ' Обмен названий
                tmp = arr(i, 0)
                arr(i, 0) = arr(j, 0)
                arr(j, 0) = tmp
                
                ' Обмен значений
                tmp = arr(i, 1)
                arr(i, 1) = arr(j, 1)
                arr(j, 1) = tmp
                
                i = i + 1
                j = j - 1
            End If
        Loop
        
        If low < j Then QuickSort arr, low, j
        If i < high Then QuickSort arr, i, high
    End If
    
    On Error GoTo 0
End Sub

