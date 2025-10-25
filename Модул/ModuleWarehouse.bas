Attribute VB_Name = "ModuleWarehouse"
' ===== КОД ДЛЯ ОСНОВНОГО МОДУЛЯ (Module1) =====
Option Explicit

' Функция запуска формы
Public Sub StartWarehouseForm()
    ' Проверяем наличие активной книги
    If ActiveWorkbook Is Nothing Then
        MsgBox "Откройте книгу Excel с листами ангаров перед запуском макроса!", vbExclamation
        Exit Sub
    End If
    
    MainWarehouseForm.Show
End Sub

' Функция очистки числового значения
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
        ' Пропускаем пустые строки
        If Trim(lines(i)) <> "" Then
            ' Разбиваем строку на части
            Dim parts() As String
            parts = Split(Trim(lines(i)))
            
            ' Проверяем формат строки
            If UBound(parts) >= 1 Then  ' Минимум два элемента: код и значение
                Dim locationParts() As String
                locationParts = Split(parts(0), "-")
                
                ' Проверяем правильность формата локации
                If UBound(locationParts) = 3 Then
                    With operations(opCount)
                        .warehouse = locationParts(0)
                        .row = locationParts(1)
                        .Letter = locationParts(2)
                        .level = locationParts(3)
                        
                        ' Обрабатываем числовое значение
                        Dim valueStr As String
                        valueStr = ""
                        
                        ' Собираем все части числа до встречи с "пар"
                        Dim j As Integer
                        For j = 1 To UBound(parts)
                            If Left(parts(j), 3) <> "пар" Then
                                If valueStr <> "" Then valueStr = valueStr & " "
                                valueStr = valueStr & parts(j)
                            Else
                                Exit For
                            End If
                        Next j
                        
                        ' Очищаем и преобразуем числовое значение
                        If valueStr <> "" Then
                            valueStr = CleanNumberString(valueStr)
                            On Error Resume Next
                            .value = CDbl(valueStr)
                            If Err.Number <> 0 Then
                                MsgBox "Ошибка в преобразовании числа: " & valueStr, vbExclamation
                                ReDim operations(0)
                                ParseInput = operations
                                Exit Function
                            End If
                            On Error GoTo 0
                        End If
                        
                        .Batch = ""
                        
                        ' Проверяем наличие номера партии
                        Dim k As Integer
                        For k = 1 To UBound(parts)
                            If Left(parts(k), 3) = "пар" Then
                                .Batch = parts(k)
                                Exit For
                            End If
                        Next k
                    End With
                    
                    opCount = opCount + 1
                End If
            End If
        End If
    Next i
    
    ' Если есть операции, корректируем размер массива
    If opCount > 0 Then
        ReDim Preserve operations(opCount - 1)
    Else
        ReDim operations(0)
    End If
    
    ParseInput = operations
End Function

' Функция для разъединения объединенных ячеек в диапазоне
Private Function UnmergeCellIfRequired(cell As Range) As Boolean
    On Error Resume Next
    If cell.MergeCells Then
        ' Запоминаем значение перед разъединением
        Dim cellValue As Variant
        cellValue = cell.value
        
        ' Разъединяем ячейки
        cell.UnMerge
        
        ' Возвращаем значение в первую ячейку диапазона
        cell.value = cellValue
        
        UnmergeCellIfRequired = True
    Else
        UnmergeCellIfRequired = False
    End If
    On Error GoTo 0
End Function

Public Function ProcessOperations(operations() As ModuleTypes.WarehouseOperation, opCount As Integer) As String
    Dim errorLog As String
    errorLog = ""
    
    ' Получаем целевую книгу
    If ActiveWorkbook Is Nothing Then
        errorLog = "Не найдена активная книга Excel!"
        ProcessOperations = errorLog
        Exit Function
    End If
    
    Dim i As Integer
    For i = 0 To opCount - 1
        ' Получаем конфигурацию для текущего ангара
        Dim sheetName As String
        Dim upperStart As Integer, upperEnd As Integer
        Dim lowerStart As Integer, lowerEnd As Integer
        
        Call ConfigModule.GetWarehouseRanges(operations(i).warehouse, sheetName, _
            upperStart, upperEnd, lowerStart, lowerEnd)
        
        ' Проверяем, что мы на правильном листе
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = ActiveWorkbook.Sheets(sheetName)
        On Error GoTo 0
        
        If ws Is Nothing Then
            errorLog = errorLog & "Лист '" & sheetName & "' не найден в активной книге!" & vbNewLine
            GoTo NextOperation
        End If
        
        ' Ищем ячейку с учетом конфигурации ангара
        Dim foundCell As Range
        Set foundCell = ConfigModule.FindCellInWarehouse(ws, operations(i).Letter, _
            operations(i).level, upperStart, upperEnd, lowerStart, lowerEnd)
        
        If foundCell Is Nothing Then
            errorLog = errorLog & "Не найдена ячейка " & operations(i).Letter & " уровень " & _
                      operations(i).level & " в ангаре " & operations(i).warehouse & vbNewLine
            GoTo NextOperation
        End If
        
        ' Определяем столбец для значения (C, E, G, ...)
        Dim valueColumn As Long
        valueColumn = 3 + (CLng(operations(i).row) - 1) * 2
        
        ' Проверяем и обновляем значение
        Dim valueCell As Range
        Dim batchCell As Range
        Set valueCell = ws.Cells(foundCell.row, valueColumn)
        Set batchCell = ws.Cells(foundCell.row, valueColumn + 1)
        
        ' Проверяем на объединенные ячейки и разъединяем при необходимости
        Dim wasUnmerged As Boolean
        wasUnmerged = UnmergeCellIfRequired(valueCell)
        If wasUnmerged Then
            errorLog = errorLog & "Внимание: Была разъединена объединенная ячейка в ангаре " & _
                      operations(i).warehouse & ", ряд " & operations(i).row & vbNewLine
        End If
        
        ' Разъединяем ячейку с партией, если нужно
        UnmergeCellIfRequired batchCell
        
        If Not IsNumeric(valueCell.value) Then valueCell.value = 0
        
        ' Выполняем операцию
        If operations(i).Batch = "" Then
            ' Вычитание
            Dim newValue As Double
            newValue = valueCell.value - operations(i).value
            
            ' Проверяем, станет ли значение отрицательным или нулевым
            If newValue <= 0 Then
                ' Очищаем ячейки значения и партии
                valueCell.ClearContents
                batchCell.ClearContents
                
                ' Добавляем информацию в лог
                If newValue < 0 Then
                    errorLog = errorLog & "Внимание: В ячейке " & operations(i).Letter & _
                              " (ряд " & operations(i).row & ") получилось отрицательное значение. " & _
                              "Ячейка очищена." & vbNewLine
                End If
            Else
                ' Если значение положительное, обновляем ячейку
                valueCell.value = newValue
            End If
        Else
            ' Добавление
            If Not IsNumeric(valueCell.value) Then valueCell.value = 0
            valueCell.value = valueCell.value + operations(i).value
            ' Записываем номер партии
            batchCell.value = operations(i).Batch
        End If
        
NextOperation:
    Next i
    
    ProcessOperations = errorLog
End Function

