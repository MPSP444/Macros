Attribute VB_Name = "ModuleFormat"
' Структура конфигурации для ангара
Private Type WarehouseConfig
    UpperDataStart As Long
    UpperDataEnd As Long
    LowerDataStart As Long
    LowerDataEnd As Long
End Type

Public Sub НастройкаВсехАнгаров()
    ' Отключаем обновление экрана
    Application.ScreenUpdating = False
    
    ' Проверяем наличие активной книги
    If ActiveWorkbook Is Nothing Then
        MsgBox "Не найдена активная книга Excel!", vbExclamation
        Exit Sub
    End If
    
    ' Инициализируем переменные
    Dim ws As Worksheet
    Dim errorLog As String
    errorLog = ""
    
    ' Обрабатываем каждый лист в активной книге
    For Each ws In ActiveWorkbook.Worksheets
        On Error Resume Next
        Select Case ws.Name
            Case "Ангар 12"
                ФорматироватьАнгар12 ws
                If Err.Number <> 0 Then
                    errorLog = errorLog & "Ошибка при форматировании Ангара 12: " & Err.description & vbNewLine
                End If
                
            Case "Ангар 11"
                ФорматироватьАнгар11 ws
                If Err.Number <> 0 Then
                    errorLog = errorLog & "Ошибка при форматировании Ангара 11: " & Err.description & vbNewLine
                End If
                
            Case "Ангар 10"
                ФорматироватьАнгар10 ws
                If Err.Number <> 0 Then
                    errorLog = errorLog & "Ошибка при форматировании Ангара 10: " & Err.description & vbNewLine
                End If
                
            Case "Ангар 9"
                ФорматироватьАнгар9 ws
                If Err.Number <> 0 Then
                    errorLog = errorLog & "Ошибка при форматировании Ангара 9: " & Err.description & vbNewLine
                End If
                
            Case "Ангар 8"
                ФорматироватьАнгар8 ws
                If Err.Number <> 0 Then
                    errorLog = errorLog & "Ошибка при форматировании Ангара 8: " & Err.description & vbNewLine
                End If
                
            Case "Ангар 7"
                ФорматироватьАнгар7 ws
                If Err.Number <> 0 Then
                    errorLog = errorLog & "Ошибка при форматировании Ангара 7: " & Err.description & vbNewLine
                End If
                
            Case "Ангар 6"
                ФорматироватьАнгар6 ws
                If Err.Number <> 0 Then
                    errorLog = errorLog & "Ошибка при форматировании Ангара 6: " & Err.description & vbNewLine
                End If
                
            Case "Ангар 5"
                ФорматироватьАнгар5 ws
                If Err.Number <> 0 Then
                    errorLog = errorLog & "Ошибка при форматировании Ангара 5: " & Err.description & vbNewLine
                End If
                
            Case "мал 10"
                ФорматироватьМал10 ws
                If Err.Number <> 0 Then
                    errorLog = errorLog & "Ошибка при форматировании мал 10: " & Err.description & vbNewLine
                End If
        End Select
        On Error GoTo 0
    Next ws
    
    ' Включаем обновление экрана
    Application.ScreenUpdating = True
    
    ' Показываем результат
    If errorLog = "" Then
        MsgBox "Форматирование всех ангаров завершено успешно!", vbInformation
    Else
        MsgBox "Форматирование завершено с ошибками:" & vbNewLine & vbNewLine & errorLog, vbExclamation
    End If
End Sub

Private Sub FormatMergedCells(ws As Worksheet, config As WarehouseConfig)
    Dim lastColumn As String
    lastColumn = "DF"
    
    ' Обработка объединенных ячеек в верхней секции
    ProcessMergedSection ws, config.UpperDataStart, config.UpperDataEnd, lastColumn
    
    ' Обработка объединенных ячеек в нижней секции (если есть)
    If config.LowerDataEnd > 0 Then
        ProcessMergedSection ws, config.LowerDataStart, config.LowerDataEnd, lastColumn
    End If
End Sub

Private Sub ProcessMergedSection(ws As Worksheet, startRow As Long, endRow As Long, lastColumn As String)
    Dim i As Long
    Dim currentCell As Range
    
    For i = startRow To endRow
        Set currentCell = ws.Cells(i, 1) ' Столбец A
        
        If currentCell.MergeCells Then
            Dim mergeHeight As Long
            mergeHeight = currentCell.MergeArea.rows.count
            
            ' Добавляем жирную линию НАД группой
            With ws.Range("A" & i & ":" & lastColumn & i).Borders(xlEdgeTop)
                .LineStyle = xlSlantDashDot
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            
            ' Добавляем жирную линию ПОД группой
            With ws.Range("A" & (i + mergeHeight - 1) & ":" & lastColumn & (i + mergeHeight - 1)).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            
            i = i + mergeHeight - 1  ' Пропускаем объединенные строки
        End If
    Next i
End Sub
Private Sub FormatGridLines(ws As Worksheet, startRow As Long, endRow As Long)
    ' Определяем полный диапазон для форматирования
    Dim fullRange As Range
    Set fullRange = ws.Range("C" & startRow & ":DF" & endRow)
    Dim abRange As Range
    Set abRange = ws.Range("A" & startRow & ":B" & endRow)
    
    ' Очищаем ВСЕ границы в основном диапазоне
    With fullRange
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
    
    ' Устанавливаем все внутренние линии как точечные для основного диапазона
    With fullRange.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    With fullRange.Borders(xlInsideVertical)
        .LineStyle = xlDot
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    ' Устанавливаем все внутренние линии как точечные для столбцов A и B
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
    
    ' Проходим по каждой строке и проверяем объединенные ячейки
    Dim i As Long
    For i = startRow To endRow
        If ws.Cells(i, 1).MergeCells Then
            Dim mergeHeight As Long
            mergeHeight = ws.Cells(i, 1).MergeArea.rows.count
            
            ' Добавляем жирную линию НАД группой
            With ws.Range("A" & i & ":DF" & i).Borders(xlEdgeTop)
                .LineStyle = xlSlantDashDot
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            
            ' Добавляем жирную линию ПОД группой
            With ws.Range("A" & (i + mergeHeight - 1) & ":DF" & (i + mergeHeight - 1)).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            
            i = i + mergeHeight - 1
        End If
    Next i
End Sub
Sub ФорматироватьАнгар12(ws As Worksheet)
    ' Создаем конфигурацию для ангара
    Dim config As WarehouseConfig
    config.UpperDataStart = 4
    config.UpperDataEnd = 66
    config.LowerDataStart = 0
    config.LowerDataEnd = 0
    
    ' Настройка строк
    ws.Range("2:2").RowHeight = 35
    ws.Range("3:3").RowHeight = 15
    ws.Range("67:67").RowHeight = 15
    ws.Range("68:68").RowHeight = 35
    
    ' Настройка столбцов A и B
    ws.Range("A:A").ColumnWidth = 8
    ws.Range("B:B").ColumnWidth = 15
    
    ' Настройка столбцов
    For Each col In ws.Range("C4:CT66").columns
        If Not col.Hidden Then
            col.ColumnWidth = 10
        End If
    Next col
    
    ' Настройка строк
    For Each row In ws.Range("C4:CT66").rows
        If Not row.Hidden Then
            row.RowHeight = 20
        End If
    Next row
    
    ' Форматирование объединенных ячеек
    FormatMergedCells ws, config
    FormatGridLines ws, 4, 66
End Sub

Sub ФорматироватьАнгар11(ws As Worksheet)
    ' Создаем конфигурацию для ангара
    Dim config As WarehouseConfig
    config.UpperDataStart = 4
    config.UpperDataEnd = 65
    config.LowerDataStart = 0
    config.LowerDataEnd = 0
    
    ' Настройка строк
    ws.Range("2:2").RowHeight = 35
    ws.Range("3:3").RowHeight = 15
    ws.Range("66:66").RowHeight = 15
    ws.Range("67:67").RowHeight = 35
    
    ' Настройка столбцов A и B
    ws.Range("A:A").ColumnWidth = 8
    ws.Range("B:B").ColumnWidth = 15
    
    ' Настройка столбцов
    For Each col In ws.Range("C4:CT65").columns
        If Not col.Hidden Then
            col.ColumnWidth = 10
        End If
    Next col
    
    ' Настройка строк
    For Each row In ws.Range("C4:CT65").rows
        If Not row.Hidden Then
            row.RowHeight = 20
        End If
    Next row
    
    ' Форматирование объединенных ячеек
    FormatMergedCells ws, config
    FormatGridLines ws, 4, 65
End Sub

Sub ФорматироватьАнгар10(ws As Worksheet)
    ' Создаем конфигурацию для ангара
    Dim config As WarehouseConfig
    config.UpperDataStart = 4
    config.UpperDataEnd = 78
    config.LowerDataStart = 0
    config.LowerDataEnd = 0
    
    ' Настройка строк
    ws.Range("2:2").RowHeight = 35
    ws.Range("3:3").RowHeight = 15
    ws.Range("79:79").RowHeight = 15
    ws.Range("80:80").RowHeight = 35
    
    ' Настройка столбцов A и B
    ws.Range("A:A").ColumnWidth = 8
    ws.Range("B:B").ColumnWidth = 15
    
    ' Настройка столбцов
    For Each col In ws.Range("C4:CT78").columns
        If Not col.Hidden Then
            col.ColumnWidth = 10
        End If
    Next col
    
    ' Настройка строк
    For Each row In ws.Range("C4:CT78").rows
        If Not row.Hidden Then
            row.RowHeight = 20
        End If
    Next row
    
    ' Форматирование объединенных ячеек
    FormatMergedCells ws, config
    FormatGridLines ws, 4, 78
End Sub

Sub ФорматироватьАнгар9(ws As Worksheet)
    ' Создаем конфигурацию для ангара
    Dim config As WarehouseConfig
    config.UpperDataStart = 4
    config.UpperDataEnd = 79
    config.LowerDataStart = 0
    config.LowerDataEnd = 0
    
    ' Настройка строк
    ws.Range("2:2").RowHeight = 35
    ws.Range("3:3").RowHeight = 15
    ws.Range("80:80").RowHeight = 15
    ws.Range("81:81").RowHeight = 35
    
    ' Настройка столбцов A и B
    ws.Range("A:A").ColumnWidth = 8
    ws.Range("B:B").ColumnWidth = 15
    
    ' Настройка столбцов
    For Each col In ws.Range("C4:CT79").columns
        If Not col.Hidden Then
            col.ColumnWidth = 10
        End If
    Next col
    
    ' Настройка строк
    For Each row In ws.Range("C4:CT79").rows
        If Not row.Hidden Then
            row.RowHeight = 20
        End If
    Next row
    
    ' Форматирование объединенных ячеек
    FormatMergedCells ws, config
    FormatGridLines ws, 4, 79
End Sub

Sub ФорматироватьАнгар8(ws As Worksheet)
    ' Создаем конфигурацию для ангара
    Dim config As WarehouseConfig
    config.UpperDataStart = 4
    config.UpperDataEnd = 55
    config.LowerDataStart = 0
    config.LowerDataEnd = 0
    
    ' Настройка строк
    ws.Range("2:2").RowHeight = 35
    ws.Range("3:3").RowHeight = 15
    ws.Range("56:56").RowHeight = 15
    ws.Range("57:57").RowHeight = 35
    
    ' Настройка столбцов A и B
    ws.Range("A:A").ColumnWidth = 8
    ws.Range("B:B").ColumnWidth = 15
    
    ' Настройка столбцов
    For Each col In ws.Range("C4:CT55").columns
        If Not col.Hidden Then
            col.ColumnWidth = 10
        End If
    Next col
    
    ' Настройка строк
    For Each row In ws.Range("C4:CT55").rows
        If Not row.Hidden Then
            row.RowHeight = 20
        End If
    Next row
    
    ' Форматирование объединенных ячеек
    FormatMergedCells ws, config
    FormatGridLines ws, 4, 55
End Sub

Sub ФорматироватьАнгар7(ws As Worksheet)
    ' Создаем конфигурацию для ангара
    Dim config As WarehouseConfig
    config.UpperDataStart = 4
    config.UpperDataEnd = 56
    config.LowerDataStart = 0
    config.LowerDataEnd = 0
    
    ' Настройка строк
    ws.Range("2:2").RowHeight = 35
    ws.Range("3:3").RowHeight = 15
    ws.Range("57:57").RowHeight = 15
    ws.Range("58:58").RowHeight = 35
    
    ' Настройка столбцов A и B
    ws.Range("A:A").ColumnWidth = 8
    ws.Range("B:B").ColumnWidth = 15
    
    ' Настройка столбцов
    For Each col In ws.Range("C4:CT56").columns
        If Not col.Hidden Then
            col.ColumnWidth = 10
        End If
    Next col
    
    ' Настройка строк
    For Each row In ws.Range("C4:CT56").rows
        If Not row.Hidden Then
            row.RowHeight = 20
        End If
    Next row
    
    ' Форматирование объединенных ячеек
    FormatMergedCells ws, config
    FormatGridLines ws, 4, 56
End Sub

Sub ФорматироватьАнгар6(ws As Worksheet)
    ' Создаем конфигурацию для ангара
    Dim config As WarehouseConfig
    config.UpperDataStart = 4
    config.UpperDataEnd = 70
    config.LowerDataStart = 0
    config.LowerDataEnd = 0
    
    ' Настройка строк
    ws.Range("2:2").RowHeight = 35
    ws.Range("3:3").RowHeight = 15
    ws.Range("71:71").RowHeight = 15
    ws.Range("72:72").RowHeight = 35
    
    ' Настройка столбц
    ' Настройка столбцов A и B
    ws.Range("A:A").ColumnWidth = 8
    ws.Range("B:B").ColumnWidth = 15
    
    ' Настройка столбцов
    For Each col In ws.Range("C4:CT70").columns
        If Not col.Hidden Then
            col.ColumnWidth = 10
        End If
    Next col
    
    ' Настройка строк
    For Each row In ws.Range("C4:CT70").rows
        If Not row.Hidden Then
            row.RowHeight = 20
        End If
    Next row
    
    ' Форматирование объединенных ячеек
    FormatMergedCells ws, config
    FormatGridLines ws, 4, 70
End Sub

Sub ФорматироватьАнгар5(ws As Worksheet)
    ' Создаем конфигурацию для ангара
    Dim config As WarehouseConfig
    config.UpperDataStart = 4
    config.UpperDataEnd = 68
    config.LowerDataStart = 0
    config.LowerDataEnd = 0
    
    ' Настройка строк
    ws.Range("2:2").RowHeight = 35
    ws.Range("3:3").RowHeight = 15
    ws.Range("69:69").RowHeight = 15
    ws.Range("70:70").RowHeight = 35
    
    ' Настройка столбцов A и B
    ws.Range("A:A").ColumnWidth = 8
    ws.Range("B:B").ColumnWidth = 15
    
    ' Настройка столбцов
    For Each col In ws.Range("C4:CT69").columns
        If Not col.Hidden Then
            col.ColumnWidth = 10
        End If
    Next col
    
    ' Настройка строк
    For Each row In ws.Range("C4:CT69").rows
        If Not row.Hidden Then
            row.RowHeight = 20
        End If
    Next row
    
    ' Форматирование объединенных ячеек
    FormatMergedCells ws, config
    FormatGridLines ws, 4, 68
End Sub

Sub ФорматироватьМал10(ws As Worksheet)
    ' Создаем конфигурацию для ангара
    Dim config As WarehouseConfig
    config.UpperDataStart = 4
    config.UpperDataEnd = 78
    config.LowerDataStart = 0
    config.LowerDataEnd = 0
    
    ' Настройка строк
    ws.Range("2:2").RowHeight = 35
    ws.Range("3:3").RowHeight = 15
    ws.Range("79:79").RowHeight = 15
    ws.Range("80:80").RowHeight = 35
    
    ' Настройка столбцов A и B
    ws.Range("A:A").ColumnWidth = 8
    ws.Range("B:B").ColumnWidth = 15
    
    ' Настройка столбцов
    For Each col In ws.Range("C4:CT78").columns
        If Not col.Hidden Then
            col.ColumnWidth = 10
        End If
    Next col
    
    ' Настройка строк
    For Each row In ws.Range("C4:CT78").rows
        If Not row.Hidden Then
            row.RowHeight = 20
        End If
    Next row
    
    ' Форматирование объединенных ячеек
    FormatMergedCells ws, config
    FormatGridLines ws, 4, 78
End Sub
