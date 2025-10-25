Attribute VB_Name = "CreateWarehouseComparisonChart"
Option Explicit

Public Sub CreateWarehouseComparisonChart()
    ' Объявление переменных
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim tempSheet As Worksheet
    Dim ProductName As String
    Dim maxValue As Double
    Dim maxWarehouse As String
    Dim row As Long
    Dim chartSheet As Worksheet
    Dim chartObj As ChartObject
    
    ' Показываем пользователю список листов для выбора
    Dim selectedSheet As String
    Dim sheetsList As String
    sheetsList = "Выберите номер листа:" & vbNewLine
    
    Dim sheetCount As Long
    sheetCount = 1
    For Each ws In ThisWorkbook.Sheets
        sheetsList = sheetsList & sheetCount & ". " & ws.Name & vbNewLine
        sheetCount = sheetCount + 1
    Next ws
    
    Dim selectedNumber As String
    selectedNumber = InputBox(sheetsList, "Выберите лист с данными")
    
    If selectedNumber = "" Then
        MsgBox "Отменено пользователем"
        Exit Sub
    End If
    
    ' Проверяем корректность ввода
    If IsNumeric(selectedNumber) Then
        If CLng(selectedNumber) > 0 And CLng(selectedNumber) <= ThisWorkbook.Sheets.count Then
            Set ws = ThisWorkbook.Sheets(CLng(selectedNumber))
        Else
            MsgBox "Неверный номер листа!"
            Exit Sub
        End If
    Else
        MsgBox "Введите число!"
        Exit Sub
    End If
    
    ' Определяем последнюю строку
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    
    ' Создаем временный лист для подготовки данных для диаграммы
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("TempChartData").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Создаем новый временный лист
    Set tempSheet = ThisWorkbook.Sheets.Add(After:=ws)
    tempSheet.Name = "TempChartData"
    
    ' Подготавливаем данные для диаграммы
    row = 1
    
    ' Добавляем заголовки
    tempSheet.Cells(row, 1) = "Наименование товара"
    tempSheet.Cells(row, 2) = "Количество"
    tempSheet.Cells(row, 3) = "Ангар"
    row = row + 1
    
    ' Находим максимальные значения для каждого товара
    For i = 2 To lastRow
        maxValue = 0
        maxWarehouse = ""
        ProductName = ws.Cells(i, 1).value
        
        ' Проверяем все ангары
        For j = 2 To 10
            If Not isEmpty(ws.Cells(i, j)) Then
                Dim currentValue As Double
                If IsNumeric(ws.Cells(i, j).value) Then
                    currentValue = CDbl(ws.Cells(i, j).value)
                    
                    If currentValue > maxValue Then
                        maxValue = currentValue
                        Select Case j
                            Case 2: maxWarehouse = "Ангар 5"
                            Case 3: maxWarehouse = "Ангар 6"
                            Case 4: maxWarehouse = "Ангар 7"
                            Case 5: maxWarehouse = "Ангар 8"
                            Case 6: maxWarehouse = "Ангар 9"
                            Case 7: maxWarehouse = "Ангар 10"
                            Case 8: maxWarehouse = "Ангар 11"
                            Case 9: maxWarehouse = "Ангар 12"
                            Case 10: maxWarehouse = "Ангар 10М"
                        End Select
                    End If
                End If
            End If
        Next j
        
        ' Записываем данные, только если есть значение
        If maxValue > 0 Then
            tempSheet.Cells(row, 1) = ProductName
            tempSheet.Cells(row, 2) = maxValue
            tempSheet.Cells(row, 3) = maxWarehouse
            row = row + 1
        End If
    Next i
    
    ' Создаем лист для диаграммы
    Set chartSheet = ThisWorkbook.Sheets.Add(After:=tempSheet)
    chartSheet.Name = "Анализ по ангарам"
    
    ' Создаем диаграмму
    Set chartObj = chartSheet.ChartObjects.Add( _
        Left:=50, _
        Width:=900, _
        Top:=25, _
        Height:=500)
    
    With chartObj.Chart
        .SetSourceData Source:=tempSheet.Range("A1:C" & row - 1)
        .ChartType = xlColumnClustered
        
        ' Настраиваем заголовок
        With .ChartTitle
            .text = "Максимальное количество товаров по ангарам"
            .Font.Size = 14
            .Font.Bold = True
        End With
        
        ' Настраиваем оси
        With .Axes(xlCategory)
            .TickLabelPosition = xlTickLabelPositionLow
            .TickLabels.Orientation = 45
        End With
        
        ' Добавляем подписи значений
        .SeriesCollection(1).HasDataLabels = True
        .SeriesCollection(1).DataLabels.ShowValue = True
        
        ' Настраиваем легенду
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
    End With
    
    ' Удаляем временный лист
    Application.DisplayAlerts = False
    tempSheet.Delete
    Application.DisplayAlerts = True
    
    MsgBox "Диаграмма создана на новом листе 'Анализ по ангарам'!", vbInformation
End Sub
