Attribute VB_Name = "ModuleOperations"
' ===== ПОЛНЫЙ ModuleOperations (ИСПРАВЛЕННАЯ ВЕРСИЯ) =====
Option Explicit

Private dict As Object
Private ws As Worksheet

Public Sub StartWarehouseForm()
    If ActiveWorkbook Is Nothing Then
        MsgBox "Откройте книгу Excel с листами ангаров перед запуском макроса!", vbExclamation
        Exit Sub
    End If
    
    MainWarehouseForm.Show
End Sub

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
    
    For Each skipRow In RowNumbersToSkip
        If row = skipRow Then
            ShouldProcessRow = False
            Exit Function
        End If
    Next skipRow
    
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
    
    If cell.MergeCells Then
        cellValue = cell.MergeArea.Cells(1, 1).value
    Else
        cellValue = cell.value
    End If
    
    If IsNumeric(cellValue) Then
        GetValueFromCell = CDbl(cellValue)
    Else
        GetValueFromCell = 0
    End If
    
    On Error GoTo 0
End Function

Private Function CleanNumberString(numberStr As String) As String
    Dim cleanStr As String
    
    ' Убираем все пробелы
    cleanStr = Replace(numberStr, " ", "")
    
    ' Заменяем точку на запятую, если есть
    cleanStr = Replace(cleanStr, ".", ",")
    
    ' Если есть запятая, проверяем дробную часть
    If InStr(cleanStr, ",") > 0 Then
        Dim parts() As String
        parts = Split(cleanStr, ",")
        
        ' Проверяем, состоит ли дробная часть только из нулей
        Dim isAllZeros As Boolean
        isAllZeros = True
        
        Dim i As Integer
        For i = 1 To Len(parts(1))
            If Mid(parts(1), i, 1) <> "0" Then
                isAllZeros = False
                Exit For
            End If
        Next i
        
        ' Если дробная часть состоит только из нулей, берём только целую часть
        If isAllZeros Then
            cleanStr = parts(0)
        End If
    End If
    
    CleanNumberString = cleanStr
End Function

Public Function ParseInput(inputStr As String) As ModuleTypes.WarehouseOperation()
    Debug.Print "=== НАЧАЛО ПАРСИНГА ==="
    Debug.Print "Входная строка:"
    Debug.Print inputStr
    
    ' Разбиваем ввод на строки
    Dim lines() As String
    lines = Split(inputStr, vbNewLine)
    
    ' Создаем массив для хранения операций
    Dim operations() As ModuleTypes.WarehouseOperation
    ReDim operations(UBound(lines))
    
    Dim opCount As Integer
    opCount = 0
    
    ' Обрабатываем каждую строку
    Dim i As Integer
    For i = 0 To UBound(lines)
        Dim line As String
        line = Trim(lines(i))
        
        ' Пропускаем пустые строки и комментарии
        If line <> "" And Left(line, 1) <> "'" Then
            Debug.Print vbNewLine & "Обрабатываем строку " & (i + 1) & ": " & line
            
            ' Разбиваем строку на части по пробелам
            Dim parts() As String
            parts = Split(line, " ")
            
            ' Убираем пустые элементы
            Dim cleanParts() As String
            Dim cleanCount As Integer
            cleanCount = 0
            
            Dim j As Integer
            For j = 0 To UBound(parts)
                If Trim(parts(j)) <> "" Then
                    ReDim Preserve cleanParts(cleanCount)
                    cleanParts(cleanCount) = Trim(parts(j))
                    Debug.Print "  Часть " & cleanCount & ": " & cleanParts(cleanCount)
                    cleanCount = cleanCount + 1
                End If
            Next j
            
            ' Проверяем формат: минимум код локации и значение
            If cleanCount >= 2 Then
                Dim locationParts() As String
                locationParts = Split(cleanParts(0), "-")
                
                Debug.Print "  Локация разобрана на " & (UBound(locationParts) + 1) & " частей"
                
                ' Проверяем правильность формата локации (4 части: ангар-ряд-ячейка-уровень)
                If UBound(locationParts) = 3 Then
                    With operations(opCount)
                        .warehouse = locationParts(0)
                        .row = locationParts(1)
                        .Letter = UCase(locationParts(2))  ' Приводим к верхнему регистру
                        .level = locationParts(3)
                        
                        ' Инициализируем значения по умолчанию
                        .value = 0
                        .Batch = ""
                        .ProductName = ""
                        
                        Debug.Print "  Локация: Ангар=" & .warehouse & ", Ряд=" & .row & ", Ячейка=" & .Letter & ", Уровень=" & .level
                        
                        ' Парсим значение и партию
                        Dim valueFound As Boolean
                        valueFound = False
                        
                        Dim k As Integer
                        For k = 1 To UBound(cleanParts)
                            Dim part As String
                            part = cleanParts(k)
                            
                            Debug.Print "  Анализируем часть: " & part
                            
                            ' Проверяем, является ли это партией (начинается с "пар")
                            If Len(part) >= 3 And LCase(Left(part, 3)) = "пар" Then
                                .Batch = part
                                Debug.Print "  ? Найдена партия: " & part
                            ' Если это не партия и мы еще не нашли значение
                            ElseIf Not valueFound Then
                                ' Очищаем и пробуем преобразовать в число
                                Dim cleanedValue As String
                                cleanedValue = CleanNumberString(part)
                                
                                If IsNumeric(cleanedValue) Then
                                    .value = CDbl(cleanedValue)
                                    valueFound = True
                                    Debug.Print "  ? Найдено значение: " & .value & " (из строки: " & part & ")"
                                End If
                            End If
                        Next k
                        
                        ' Проверяем, что значение было найдено
                        If Not valueFound Then
                            Debug.Print "  ? ОШИБКА: Не найдено числовое значение в строке: " & line
                            GoTo NextLine
                        End If
                        
                        ' Определяем тип операции и логируем
                        If .Batch <> "" Then
                            Debug.Print "  > ДОБАВЛЕНИЕ: Ангар " & .warehouse & ", ряд " & .row & _
                                       ", ячейка " & .Letter & ", уровень " & .level & _
                                       ", значение " & .value & ", партия " & .Batch
                        Else
                            Debug.Print "  > ВЫЧИТАНИЕ: Ангар " & .warehouse & ", ряд " & .row & _
                                       ", ячейка " & .Letter & ", уровень " & .level & _
                                       ", значение " & .value
                        End If
                    End With
                    
                    opCount = opCount + 1
                Else
                    Debug.Print "  ? ОШИБКА: Неверный формат локации (ожидается ангар-ряд-ячейка-уровень): " & cleanParts(0)
                End If
            Else
                Debug.Print "  ? ОШИБКА: Недостаточно параметров в строке (минимум: локация значение)"
            End If
        Else
            If line <> "" Then
                Debug.Print "Пропущена строка (комментарий): " & line
            End If
        End If
        
NextLine:
    Next i
    
    ' Корректируем размер массива
    If opCount > 0 Then
        ReDim Preserve operations(opCount - 1)
        Debug.Print vbNewLine & "? ПАРСИНГ ЗАВЕРШЕН. Всего операций: " & opCount
    Else
        ReDim operations(0)
        Debug.Print vbNewLine & "? ПАРСИНГ ЗАВЕРШЕН. Операций не найдено!"
    End If
    
    ParseInput = operations
End Function

Public Function ProcessOperations(operations() As ModuleTypes.WarehouseOperation, opCount As Integer) As String
    Dim errorLog As String
    errorLog = ""
    
    If ActiveWorkbook Is Nothing Then
        errorLog = "Не найдена активная книга Excel!"
        ProcessOperations = errorLog
        Exit Function
    End If
    
    Debug.Print vbNewLine & "=== НАЧАЛО ОБРАБОТКИ ОПЕРАЦИЙ ==="
    Debug.Print "Всего операций к обработке: " & opCount
    
    Application.ScreenUpdating = False
    
    Dim i As Integer
    For i = 0 To opCount - 1
        Debug.Print vbNewLine & "--- Операция " & (i + 1) & " из " & opCount & " ---"
        
        ' Получаем конфигурацию ангара
        Dim config As ModuleTypes.WarehouseConfig
        config = ModuleConfig.GetWarehouseConfig(operations(i).warehouse)
        
        Debug.Print "Ангар " & operations(i).warehouse & " > лист: " & config.sheetName
        
        ' Проверяем лист
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = ActiveWorkbook.Sheets(config.sheetName)
        On Error GoTo 0
        
        If ws Is Nothing Then
            errorLog = errorLog & "Лист '" & config.sheetName & "' не найден!" & vbNewLine
            Debug.Print "? ОШИБКА: Лист не найден - " & config.sheetName
            GoTo NextOperation
        End If
        
        Debug.Print "? Лист найден: " & ws.Name
        
        ' Ищем ячейку с помощью улучшенной функции
        Dim foundCell As Range
        Set foundCell = ModuleConfig.FindCellInWarehouseSection(ws, operations(i).Letter, _
            operations(i).level, operations(i).warehouse)
        
        If foundCell Is Nothing Then
            errorLog = errorLog & "Не найдена ячейка " & operations(i).Letter & _
                      " уровень " & operations(i).level & " в ангаре " & _
                      operations(i).warehouse & vbNewLine
            Debug.Print "? ОШИБКА: Ячейка не найдена - " & operations(i).Letter & "/" & operations(i).level
            GoTo NextOperation
        End If
        
        Debug.Print "? Найдена ячейка: " & foundCell.address & " = '" & foundCell.value & "'"
        
        ' Определяем столбец для значения (C=3, E=5, G=7, I=9...)
        Dim valueColumn As Long
        valueColumn = 3 + (CLng(operations(i).row) - 1) * 2
        
        Dim valueCell As Range, batchCell As Range
        Set valueCell = ws.Cells(foundCell.row, valueColumn)
        Set batchCell = ws.Cells(foundCell.row, valueColumn + 1)
        
        Debug.Print "Ячейка значения: " & valueCell.address & " = '" & valueCell.value & "'"
        Debug.Print "Ячейка партии: " & batchCell.address & " = '" & batchCell.value & "'"
        
        ' ОСНОВНАЯ ЛОГИКА ОПЕРАЦИЙ
        If operations(i).Batch = "" Then
            ' === ОПЕРАЦИЯ ВЫЧИТАНИЯ ===
            Debug.Print "> Выполняется ВЫЧИТАНИЕ " & operations(i).value
            
            Dim currentValue As Double
            If IsNumeric(valueCell.value) And valueCell.value <> "" Then
                currentValue = CDbl(valueCell.value)
            Else
                currentValue = 0
            End If
            
            Dim newValue As Double
            newValue = currentValue - operations(i).value
            
            Debug.Print "  Текущее: " & currentValue & " - Вычитаем: " & operations(i).value & " = Результат: " & newValue
            
            If newValue <= 0 Then
                ' Очищаем ячейки при нулевом или отрицательном результате
                valueCell.ClearContents
                batchCell.ClearContents
                Debug.Print "  ? Ячейки очищены (результат <= 0)"
                
                If newValue < 0 Then
                    errorLog = errorLog & "Внимание: В ячейке " & operations(i).Letter & _
                              " (ряд " & operations(i).row & ") получилось отрицательное значение (" & _
                              Format(newValue, "0.##") & "). Ячейка очищена." & vbNewLine
                    Debug.Print "  ? Внимание: отрицательный результат"
                End If
            Else
                valueCell.value = newValue
                Debug.Print "  ? Установлено новое значение: " & newValue
            End If
            
        Else
            ' === ОПЕРАЦИЯ ДОБАВЛЕНИЯ ===
            Debug.Print "> Выполняется ДОБАВЛЕНИЕ " & operations(i).value & " (партия: " & operations(i).Batch & ")"
            
            Dim oldValue As Double
            If IsNumeric(valueCell.value) And valueCell.value <> "" Then
                oldValue = CDbl(valueCell.value)
            Else
                oldValue = 0
            End If
            
            Dim addedValue As Double
            addedValue = oldValue + operations(i).value
            
            valueCell.value = addedValue
            batchCell.value = operations(i).Batch
            
            Debug.Print "  Было: " & oldValue & " + Добавляем: " & operations(i).value & " = Стало: " & addedValue
            Debug.Print "  ? Записана партия: " & operations(i).Batch
        End If
        
NextOperation:
    Next i
    
    Application.ScreenUpdating = True
    Debug.Print vbNewLine & "=== ОБРАБОТКА ОПЕРАЦИЙ ЗАВЕРШЕНА ==="
    
    If errorLog = "" Then
        Debug.Print "? Все операции выполнены успешно!"
    Else
        Debug.Print "? Есть предупреждения:"
        Debug.Print errorLog
    End If
    
    ProcessOperations = errorLog
End Function
