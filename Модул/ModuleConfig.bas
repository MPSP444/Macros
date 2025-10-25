Attribute VB_Name = "ModuleConfig"

' ===== ModuleConfig =====
Option Explicit

Public Function GetWarehouseConfig(warehouseNumber As String) As ModuleTypes.WarehouseConfig
    Dim config As ModuleTypes.WarehouseConfig
    
    Select Case warehouseNumber
        Case "5"
            With config
                .sheetName = "Ангар 5"
                .UpperDataStart = 4
                .UpperDataEnd = 30
                .UpperHeaderRow = 2    ' Добавлено
                .LowerDataStart = 42
                .LowerDataEnd = 68
                .LowerHeaderRow = 70   ' Добавлено
                .outputStartRow = 115  ' Добавлено
                .RowNumbersToSkip = Array(3, 69)
            End With
            
        Case "6"
            With config
                .sheetName = "Ангар 6"
                .UpperDataStart = 4
                .UpperDataEnd = 30
                .UpperHeaderRow = 2
                .LowerDataStart = 44
                .LowerDataEnd = 70
                .LowerHeaderRow = 72
                .outputStartRow = 74
                .RowNumbersToSkip = Array(3, 71)
            End With
            
        Case "7"
            With config
                .sheetName = "Ангар 7"
                .UpperDataStart = 4
                .UpperDataEnd = 24
                .UpperHeaderRow = 2
                .LowerDataStart = 36
                .LowerDataEnd = 56
                .LowerHeaderRow = 58
                .outputStartRow = 60
                .RowNumbersToSkip = Array(3, 57)
            End With
            
        Case "8"
            With config
                .sheetName = "Ангар 8"
                .UpperDataStart = 4
                .UpperDataEnd = 24
                .UpperHeaderRow = 2
                .LowerDataStart = 35
                .LowerDataEnd = 55
                .LowerHeaderRow = 57
                .outputStartRow = 115
                .RowNumbersToSkip = Array(3, 56)
            End With
            
        Case "9"
            With config
                .sheetName = "Ангар 9"
                .UpperDataStart = 4
                .UpperDataEnd = 33
                .UpperHeaderRow = 2
                .LowerDataStart = 50
                .LowerDataEnd = 79
                .LowerHeaderRow = 81
                .outputStartRow = 83
                .RowNumbersToSkip = Array(3, 80)
            End With
            
        Case "10"
            With config
                .sheetName = "Ангар 10"
                .UpperDataStart = 4
                .UpperDataEnd = 34
                .UpperHeaderRow = 2
                .LowerDataStart = 48
                .LowerDataEnd = 78
                .LowerHeaderRow = 80
                .outputStartRow = 82
                .RowNumbersToSkip = Array(3, 79)
            End With
            
        Case "11"
            With config
                .sheetName = "Ангар 11"
                .UpperDataStart = 4
                .UpperDataEnd = 27
                .UpperHeaderRow = 2
                .LowerDataStart = 42
                .LowerDataEnd = 65
                .LowerHeaderRow = 67
                .outputStartRow = 69
                .RowNumbersToSkip = Array(3, 66)
            End With
            
        Case "12"
            With config
                .sheetName = "Ангар 12"
                .UpperDataStart = 4
                .UpperDataEnd = 27
                .UpperHeaderRow = 2
                .LowerDataStart = 43
                .LowerDataEnd = 66
                .LowerHeaderRow = 68
                .outputStartRow = 115
                .RowNumbersToSkip = Array(3, 67)
            End With
    End Select
    
    GetWarehouseConfig = config
End Function

Public Function GetLevelText(Letter As String, level As String) As String
    Letter = UCase(Letter)
    Select Case level
        Case "1": GetLevelText = Letter & "/ниж.ряд"
        Case "2": GetLevelText = Letter & "/2ряд"
        Case "3": GetLevelText = Letter & "/3ряд"
        Case Else: GetLevelText = Letter & "/" & level
    End Select
End Function

Public Sub GetWarehouseRanges(ByVal warehouseNumber As String, _
                            ByRef sheetName As String, _
                            ByRef upperStart As Integer, _
                            ByRef upperEnd As Integer, _
                            ByRef lowerStart As Integer, _
                            ByRef lowerEnd As Integer)
    Dim config As ModuleTypes.WarehouseConfig  ' Исправляем здесь
    config = GetWarehouseConfig(warehouseNumber)
    
    sheetName = config.sheetName
    upperStart = config.UpperDataStart
    upperEnd = config.UpperDataEnd
    lowerStart = config.LowerDataStart
    lowerEnd = config.LowerDataEnd
End Sub

Public Function FindCellInWarehouse(ws As Worksheet, _
                                  Letter As String, _
                                  level As String, _
                                  upperStart As Integer, _
                                  upperEnd As Integer, _
                                  lowerStart As Integer, _
                                  lowerEnd As Integer) As Range
    Dim foundCell As Range
    Dim upperRange As Range
    Dim lowerRange As Range
    Dim searchText As String
    
    searchText = GetLevelText(Letter, level)
    
    Set upperRange = ws.Range("A" & upperStart & ":B" & upperEnd)
    Set lowerRange = ws.Range("A" & lowerStart & ":B" & lowerEnd)
    
    Set foundCell = upperRange.Find(What:=searchText, _
                                  LookIn:=xlValues, _
                                  LookAt:=xlWhole, _
                                  SearchOrder:=xlByRows, _
                                  SearchDirection:=xlNext, _
                                  MatchCase:=False)
                                  
    If foundCell Is Nothing Then
        Set foundCell = lowerRange.Find(What:=searchText, _
                                      LookIn:=xlValues, _
                                      LookAt:=xlWhole, _
                                      SearchOrder:=xlByRows, _
                                      SearchDirection:=xlNext, _
                                      MatchCase:=False)
    End If
    
    If foundCell Is Nothing Then
        searchText = GetLevelText(LCase(Letter), level)
        
        Set foundCell = upperRange.Find(What:=searchText, _
                                      LookIn:=xlValues, _
                                      LookAt:=xlWhole, _
                                      SearchOrder:=xlByRows, _
                                      SearchDirection:=xlNext, _
                                      MatchCase:=False)
                                      
        If foundCell Is Nothing Then
            Set foundCell = lowerRange.Find(What:=searchText, _
                                          LookIn:=xlValues, _
                                          LookAt:=xlWhole, _
                                          SearchOrder:=xlByRows, _
                                          SearchDirection:=xlNext, _
                                          MatchCase:=False)
        End If
    End If
    
    Set FindCellInWarehouse = foundCell
End Function

' ===== НОВЫЕ ФУНКЦИИ ДЛЯ РАБОТЫ С СЕКЦИЯМИ =====

Public Function DetermineSectionByShelf(Shelf As String, warehouse As String) As String
    ' Определяет секцию (UPPER/LOWER) по букве стеллажа и номеру ангара
    Dim upperShelves As String
    Dim lowerShelves As String
    
    ' Получаем списки букв для каждого ангара
    Select Case warehouse
        Case "5", "6"
            upperShelves = "А,Б,В,Г,Д,Е,Ж,З,ПРЗ"
            lowerShelves = "ПРИ,И,К,Л,М,Н,О,П,Р"
            
        Case "7", "8"
            upperShelves = "А,Б,В,Г,Д,Е,ПРЕ"
            lowerShelves = "ПРЖ,Ж,З,И,К,Л,М"
            
        Case "9"
            upperShelves = "А,Б,В,Г,Д,Е,Ж,З,И,ПРИ"
            lowerShelves = "ПРМ,М,Н,О,П,Р,С,Т,У,Ф"
            
        Case "10"
            upperShelves = "А,Б,В,Г,Д,Е,Ж,З,И,ПРИ"
            lowerShelves = "ПРК,К,Л,М,Н,О,П,Р,С,Т"
            
        Case "11"
            upperShelves = "А,Б,В,Г,Д,Е,Ж,ПРЗ"
            lowerShelves = "ПРИ,З,И,К,Л,М,Н,О"
            
        Case "12"
            upperShelves = "А,Б,В,Г,Д,Е,Ж,ПРЖ"
            lowerShelves = "ПРЗ,З,И,К,Л,М,Н,О"
            
        Case Else
            DetermineSectionByShelf = "UPPER" ' По умолчанию
            Exit Function
    End Select
    
    ' Проверяем принадлежность к секции
    If InStr("," & upperShelves & ",", "," & UCase(Shelf) & ",") > 0 Then
        DetermineSectionByShelf = "UPPER"
    ElseIf InStr("," & lowerShelves & ",", "," & UCase(Shelf) & ",") > 0 Then
        DetermineSectionByShelf = "LOWER"
    Else
        DetermineSectionByShelf = "UPPER" ' По умолчанию, если не найдено
    End If
End Function

Public Function FindCellInWarehouseSection(ws As Worksheet, _
                                         Letter As String, _
                                         level As String, _
                                         warehouse As String) As Range
    ' Улучшенная версия поиска ячейки с учетом секций
    Dim config As ModuleTypes.WarehouseConfig
    config = GetWarehouseConfig(warehouse)
    
    Dim section As String
    section = DetermineSectionByShelf(Letter, warehouse)
    
    Dim searchRange As Range
    Dim searchText As String
    searchText = GetLevelText(Letter, level)
    
    ' Ищем только в нужной секции
    If section = "UPPER" Then
        Set searchRange = ws.Range("A" & config.UpperDataStart & ":B" & config.UpperDataEnd)
    Else
        Set searchRange = ws.Range("A" & config.LowerDataStart & ":B" & config.LowerDataEnd)
    End If
    
    ' Поиск с учетом регистра
    Dim foundCell As Range
    Set foundCell = searchRange.Find(What:=searchText, _
                                   LookIn:=xlValues, _
                                   LookAt:=xlWhole, _
                                   SearchOrder:=xlByRows, _
                                   SearchDirection:=xlNext, _
                                   MatchCase:=False)
                                   
    If foundCell Is Nothing Then
        ' Пробуем другой регистр
        searchText = GetLevelText(LCase(Letter), level)
        Set foundCell = searchRange.Find(What:=searchText, _
                                       LookIn:=xlValues, _
                                       LookAt:=xlWhole, _
                                       SearchOrder:=xlByRows, _
                                       SearchDirection:=xlNext, _
                                       MatchCase:=False)
    End If
    
    Set FindCellInWarehouseSection = foundCell
End Function

Public Function GetSectionName(section As String) As String
    ' Возвращает читаемое название секции
    Select Case UCase(section)
        Case "UPPER": GetSectionName = "верхняя"
        Case "LOWER": GetSectionName = "нижняя"
        Case Else: GetSectionName = "неизвестная"
    End Select
End Function

Public Function GetWarehouseSectionsList(warehouse As String) As String
    ' Возвращает список букв в секциях для отладки
    Dim result As String
    result = "Ангар " & warehouse & ":" & vbNewLine
    
    Select Case warehouse
        Case "5", "6"
            result = result & "Верхняя: А,Б,В,Г,Д,Е,Ж,З,ПРЗ" & vbNewLine
            result = result & "Нижняя: ПРИ,И,К,Л,М,Н,О,П,Р"
            
        Case "7", "8"
            result = result & "Верхняя: А,Б,В,Г,Д,Е,ПРЕ" & vbNewLine
            result = result & "Нижняя: ПРЖ,Ж,З,И,К,Л,М"
            
        Case "9"
            result = result & "Верхняя: А,Б,В,Г,Д,Е,Ж,З,И,ПРИ" & vbNewLine
            result = result & "Нижняя: ПРМ,М,Н,О,П,Р,С,Т,У,Ф"
            
        Case "10"
            result = result & "Верхняя: А,Б,В,Г,Д,Е,Ж,З,И,ПРИ" & vbNewLine
            result = result & "Нижняя: ПРК,К,Л,М,Н,О,П,Р,С,Т"
            
        Case "11"
            result = result & "Верхняя: А,Б,В,Г,Д,Е,Ж,ПРЗ" & vbNewLine
            result = result & "Нижняя: ПРИ,З,И,К,Л,М,Н,О"
            
        Case "12"
            result = result & "Верхняя: А,Б,В,Г,Д,Е,Ж,ПРЖ" & vbNewLine
            result = result & "Нижняя: ПРЗ,З,И,К,Л,М,Н,О"
            
        Case Else
            result = result & "Неизвестный ангар!"
    End Select
    
    GetWarehouseSectionsList = result
End Function

