Attribute VB_Name = "Отсортировать"
Option Explicit

' Функция для правильного форматирования текста
Private Function NormalizeText(txt As String) As String
    Dim words() As String
    Dim i As Long
    Dim result As String
    
    ' Разбиваем строку на слова
    words = Split(LCase(txt), " ")
    
    ' Обрабатываем каждое слово
    For i = 0 To UBound(words)
        If Len(words(i)) > 0 Then
            ' Делаем первую букву заглавной, остальные строчные
            words(i) = UCase(Left(words(i), 1)) & Mid(words(i), 2)
        End If
    Next i
    
    ' Соединяем слова обратно
    NormalizeText = Join(words, " ")
End Function

Sub SortWarehouseData()
    Dim ws As Worksheet
    Dim wbTarget As Workbook
    Dim dict As Object
    Dim cell As Range
    Dim i As Long, writeRow As Long
    Dim item As Variant
    Dim startRow As Long
    Dim wsName As String
    Dim normalizedName As String
    
    ' Указываем активную книгу
    Set wbTarget = ActiveWorkbook
    
    ' Создаем словарь для хранения товаров и их количества
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Массив с информацией об ангарах: имя листа и начальная строка
    Dim warehouseInfo
    warehouseInfo = Array( _
        Array("Ангар 5", 72), _
        Array("Ангар 6", 74), _
        Array("Ангар 7", 60), _
        Array("Ангар 8", 59), _
        Array("Ангар 9", 83), _
        Array("Ангар 10", 82), _
        Array("Ангар 11", 69), _
        Array("Ангар 12", 70) _
    )
    
    ' Обрабатываем каждый ангар
    For i = 0 To UBound(warehouseInfo)
        ' Очищаем словарь для нового ангара
        dict.RemoveAll
        
        wsName = warehouseInfo(i)(0)
        
        ' Проверяем существование листа в активной книге
        On Error Resume Next
        Set ws = wbTarget.Worksheets(wsName)
        On Error GoTo 0
        
        If ws Is Nothing Then
            MsgBox "Лист '" & wsName & "' не найден в книге '" & wbTarget.Name & "'. Пропускаем.", vbExclamation
            GoTo NextWarehouse
        End If
        
        startRow = warehouseInfo(i)(1)
        
        ' Собираем данные в словарь, проверяя все ячейки до 1000 строки
        Dim currentRow As Long
        For currentRow = startRow To 1000
            If Not ws.Cells(currentRow, "A").MergeCells Then
                ' Проверяем, есть ли данные в обоих столбцах
                If Not isEmpty(ws.Cells(currentRow, "A")) And _
                   Not isEmpty(ws.Cells(currentRow, "B")) Then
                    
                    ' Нормализуем название препарата
                    normalizedName = NormalizeText(ws.Cells(currentRow, "A").value)
                    
                    If dict.exists(normalizedName) Then
                        dict(normalizedName) = _
                            dict(normalizedName) + _
                            ws.Cells(currentRow, "B").value
                    Else
                        dict.Add normalizedName, _
                                ws.Cells(currentRow, "B").value
                    End If
                End If
            End If
        Next currentRow
        
        ' Очищаем весь диапазон от startRow до 1000
        ws.Range("A" & startRow & ":B1000").Clear
        
        ' Создаем временный массив для сортировки
        If dict.count > 0 Then
            Dim sortArray()
            ReDim sortArray(0 To dict.count - 1)
            Dim j As Long
            j = 0
            
            ' Заполняем массив для сортировки
            For Each item In dict.keys
                sortArray(j) = item
                j = j + 1
            Next item
            
            ' Сортируем массив по алфавиту
            Call QuickSort(sortArray, 0, dict.count - 1)
        
            ' Записываем заголовки точно на startRow
            With ws
                If Not .Cells(startRow, "A").MergeCells Then
                    .Cells(startRow, "A").value = "Наименование"
                    .Cells(startRow, "B").value = "Количество"
                    
                    ' Форматируем заголовки
                    With .Range(.Cells(startRow, "A"), .Cells(startRow, "B"))
                        .Font.Bold = True
                        .HorizontalAlignment = xlCenter
                        .Interior.Color = RGB(255, 255, 255)
                        .Font.Color = RGB(255, 255, 255) 'ak swet' Светло-синий цвет
                    End With
                End If
            End With
            
            ' Записываем отсортированные данные, начиная со следующей строки после заголовков
            writeRow = startRow + 1
            For j = 0 To UBound(sortArray)
                ' Проверяем, не объединена ли ячейка перед записью
                If Not ws.Cells(writeRow, "A").MergeCells Then
                    With ws.Cells(writeRow, "A")
                        .value = sortArray(j) ' Уже нормализованное название
                        .Font.Name = "Bradley Hand ITC"
                        .Font.Size = 11
                        .Font.Color = RGB(255, 255, 255) 'ak swet
                    End With
                    With ws.Cells(writeRow, "B")
                        .value = dict(sortArray(j))
                        .Font.Name = "Bradley Hand ITC"
                        .Font.Size = 11
                        .Font.Color = RGB(255, 255, 255) 'ak swet
                    End With
                    writeRow = writeRow + 1
                End If
            Next j
            
            ' Автоподбор ширины столбцов
            ws.columns("A:B").AutoFit
        End If
        
NextWarehouse:
    Next i
    
    MsgBox "Обработка завершена!", vbInformation
End Sub

' Процедура быстрой сортировки
Private Sub QuickSort(ByRef arr() As Variant, ByVal first As Long, ByVal last As Long)
    Dim pivot As Variant
    Dim low As Long
    Dim high As Long
    Dim temp As Variant
    
    low = first
    high = last
    pivot = arr((first + last) \ 2)
    
    Do While low <= high
        Do While arr(low) < pivot
            low = low + 1
        Loop
        
        Do While arr(high) > pivot
            high = high - 1
        Loop
        
        If low <= high Then
            temp = arr(low)
            arr(low) = arr(high)
            arr(high) = temp
            low = low + 1
            high = high - 1
        End If
    Loop
    
    If first < high Then QuickSort arr, first, high
    If low < last Then QuickSort arr, low, last
End Sub

