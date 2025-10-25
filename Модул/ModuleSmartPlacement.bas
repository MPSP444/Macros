Attribute VB_Name = "ModuleSmartPlacement"
' ===== ModuleSmartPlacement - ��������� ������������ ������ "������ �� �������" =====
' ? ���� ���������: �1>�2>�3>�1>�2>�3 (������������ ���������� ������ �����)
' ? �� �1>�1>�1>�1 (�������������� ����������)
' ? �������� ������ ������ ��������
' ? ���������� ������ �� �������
' ? ����� �� ������� � �����, ����� �� ����� � �������

Option Explicit

' ��������� ��� ���������� � ����������
Public Type PlacementInfo
    warehouse As String      ' ����� ������
    row As String           ' ����� ����
    Shelf As String         ' ����� ��������
    level As String         ' ������� (����)
    ProductName As String   ' �������� ������
    Batch As String         ' ������
    quantity As Double      ' ���������� ��� ����������
    standardVolume As Double ' ����� ����������� �������
    pieceType As String     ' ���: WHOLE (�����), MEDIUM (�������), PIECE (�����)
    weightCategory As String ' HEAVY, MEDIUM, LIGHT (��� ������)
    IsPlaced As Boolean     ' ������� ���������?
    ErrorMessage As String  ' ��������� �� ������
    placementDetails As String ' ������ ���������� ��� ������
End Type

' ��������� ��� ������� ������������� ��������
Private Type shelfAnalysis
    Letter As String        ' ����� (�, �, �...)
    totalQuantity As Double ' ����� ���������� �� ��������
    EmptyLevel As String    ' ������ ������ ������� (1, 2, 3)
    HasProduct As Boolean   ' ���� �� ��� ����� �� ���� ��������
    CanAccommodate As Boolean ' ����� �� �������� ����� ����������
    Priority As Integer     ' ��������� ���������� (1=������)
    ShelfLimit As Double    ' ����� ����� ��������
    Level1Quantity As Double ' ���������� �� ������ 1
    Level2Quantity As Double ' ���������� �� ������ 2
    Level3Quantity As Double ' ���������� �� ������ 3
    SuitableLevel As String  ' ���������� ������� ��� ������� ����
End Type

' ��������� ��� ���������� ������
Private Type FoundProductInfo
    warehouse As String     ' ����� ������
    startRow As Integer     ' ��������� ��� ������
    endRow As Integer       ' �������� ��� ������
    availableRows As String ' ������ ��������� �����
    section As String      ' ������ (UPPER/LOWER)
    totalQuantity As Double ' ����� ���������� ������ � ������
    headerRow As Long      ' ������ ���������
    productInHeader As String ' ����� � ���������
End Type

' ===== ������� ������� �� ������� � ������� =====

Private Function GetShelfLimit(warehouseNum As String, productVolume As Double) As Double
    Select Case warehouseNum
        Case "5", "6", "9", "10", "11"
            Select Case CInt(productVolume)
                Case 1000: GetShelfLimit = 2000
                Case 960:  GetShelfLimit = 1000
                Case 720:  GetShelfLimit = 1500
                Case 600:  GetShelfLimit = 1300
                Case 540:  GetShelfLimit = 1100
                Case 480:  GetShelfLimit = 980
                Case 288:  GetShelfLimit = 580
                Case 240:  GetShelfLimit = 500
                Case Else: GetShelfLimit = 1500
            End Select
            
        Case "7", "8"
            Select Case CInt(productVolume)
                Case 1000: GetShelfLimit = 2000
                Case 960:  GetShelfLimit = 1000
                Case 720:  GetShelfLimit = 1500
                Case 600:  GetShelfLimit = 1300
                Case 540:  GetShelfLimit = 1100
                Case 480:  GetShelfLimit = 980
                Case 288:  GetShelfLimit = 580
                Case 240:  GetShelfLimit = 500
                Case Else: GetShelfLimit = 1500
            End Select
            
        Case "12"
            Select Case CInt(productVolume)
                Case 1000: GetShelfLimit = 2000
                Case 960:  GetShelfLimit = 1000
                Case 720:  GetShelfLimit = 1500
                Case 600:  GetShelfLimit = 1300
                Case 540:  GetShelfLimit = 1100
                Case 480:  GetShelfLimit = 980
                Case 288:  GetShelfLimit = 580
                Case 240:  GetShelfLimit = 500
                Case Else: GetShelfLimit = 1500
            End Select
            
        Case Else
            GetShelfLimit = 1500
    End Select
End Function

Private Function GetRowLimit(warehouseNum As String, productVolume As Double) As Double
    Select Case warehouseNum
        Case "5", "6", "9", "10", "11"
            Select Case CInt(productVolume)
                Case 1000: GetRowLimit = 14000
                Case 960:  GetRowLimit = 6800
                Case 720:  GetRowLimit = 10100
                Case 600:  GetRowLimit = 8500
                Case 540:  GetRowLimit = 8700
                Case 480:  GetRowLimit = 7000
                Case 288:  GetRowLimit = 4608
                Case 240:  GetRowLimit = 3850
                Case Else: GetRowLimit = 10100
            End Select
            
        Case "7", "8"
            Select Case CInt(productVolume)
                Case 1000: GetRowLimit = 10000
                Case 960:  GetRowLimit = 5000
                Case 720:  GetRowLimit = 7200
                Case 600:  GetRowLimit = 6000
                Case 540:  GetRowLimit = 6500
                Case 480:  GetRowLimit = 5000
                Case 288:  GetRowLimit = 3800
                Case 240:  GetRowLimit = 3000
                Case Else: GetRowLimit = 7200
            End Select
            
        Case "12"
            Select Case CInt(productVolume)
                Case 1000: GetRowLimit = 16000
                Case 960:  GetRowLimit = 7700
                Case 720:  GetRowLimit = 11550
                Case 600:  GetRowLimit = 9600
                Case 540:  GetRowLimit = 8700
                Case 480:  GetRowLimit = 7000
                Case 288:  GetRowLimit = 5000
                Case 240:  GetRowLimit = 4320
                Case Else: GetRowLimit = 11550
            End Select
            
        Case Else
            GetRowLimit = 10100
    End Select
End Function

Private Function DetermineWeightCategory(quantity As Double, standardVolume As Double) As String
    Dim percentage As Double
    If standardVolume > 0 Then
        percentage = (quantity / standardVolume) * 100
    Else
        percentage = 100
    End If
    
    If percentage >= 80 Then
        DetermineWeightCategory = "HEAVY"
    ElseIf percentage >= 50 Then
        DetermineWeightCategory = "MEDIUM"
    Else
        DetermineWeightCategory = "LIGHT"
    End If
    
    ModuleLogger.LogDebug "?? ���: " & quantity & "/" & standardVolume & " = " & _
                         Format(percentage, "0.0") & "% > " & DetermineWeightCategory
End Function

' ===== ?? ������� ������� ������ ���������� (������ �� �������) =====
Public Function SmartPlaceProducts(inputText As String) As String
    On Error GoTo ErrorHandler
    
    ModuleLogger.LogMessage "=== ?? ����� ���������� '������ �� �������' ==="
    ModuleLogger.LogMessage "? ����������: �1>�2>�3>�1>�2>�3 (���������)"
    ModuleLogger.LogMessage "? ����: �1>�1>�1>�1 (�����������)"
    
    ' ������������� ������
    Call ModuleProductInfo.InitializeProductDatabase
    Call ModuleWarehouseCapacity.InitializeWarehouseCapacities
    
    Application.ScreenUpdating = False
    
    ' ��������� �������� ������
    Dim activeCount As Integer
    activeCount = ModuleWarehouseCapacity.GetActiveWarehouseCount()
    
    If activeCount = 0 Then
        SmartPlaceProducts = "? �� ������ �� ���� �������� �����!"
        Application.ScreenUpdating = True
        Exit Function
    End If
    
    ModuleLogger.LogMessage "? �������� ������: " & ModuleWarehouseCapacity.GetActiveWarehousesList()
    
    ' 1. ������� � �������� �������
    Dim placements() As PlacementInfo
    placements = ParseNewFormatInput(inputText)
    
    If UBound(placements) < 0 Then
        SmartPlaceProducts = "�� ������� ���������� ������ ��� ����������!"
        Application.ScreenUpdating = True
        Exit Function
    End If
    
    ' 2. ���������� ������ �� �������
    Dim i As Integer
    Dim successCount As Integer
    successCount = 0
    
    For i = 0 To UBound(placements)
        If placements(i).ProductName <> "" Then
            ' ���������� ��������� ����
            placements(i).weightCategory = DetermineWeightCategory(placements(i).quantity, placements(i).standardVolume)
            
            ModuleLogger.LogMessage "?? ���������� ������-��-�������: " & placements(i).ProductName & " - " & _
                                   placements(i).quantity & " (" & placements(i).standardVolume & ") - " & _
                                   placements(i).pieceType & " - ���: " & placements(i).weightCategory
            
            ' ?? ���������� ������ ������ �� �������
            If SmartPlaceWithCellByCell(placements(i)) Then
                successCount = successCount + 1
                ModuleLogger.LogSuccess "? ��������� ������-��-�������: " & placements(i).ProductName
            Else
                ModuleLogger.LogError "? ������ ����������: " & placements(i).ProductName & " - " & placements(i).ErrorMessage
            End If
        End If
    Next i
    
    ' 3. ��������� ������
    Dim report As String
    report = GenerateCellByCellReport(placements, successCount, activeCount)  ' ? ����������: ����� ������
    
    Application.ScreenUpdating = True
    
    ModuleLogger.LogMessage "=== ?? ���������� ���������� '������ �� �������' ==="
    ModuleLogger.LogSuccess "������� ���������: " & successCount & " �� " & (UBound(placements) + 1)
    
    SmartPlaceProducts = report
    Exit Function
    
ErrorHandler:
    Application.ScreenUpdating = True
    SmartPlaceProducts = "����������� ������: " & Err.description
    ModuleLogger.LogError "����������� ������: " & Err.description
End Function

' ===== ?? ���������� � ������� ������ �� ������� =====
Private Function SmartPlaceWithCellByCell(ByRef placement As PlacementInfo) As Boolean
    On Error GoTo ErrorHandler
    
    ModuleLogger.LogMessage "?? ���������� ������-��-�������: " & placement.ProductName & _
                           " (" & placement.quantity & "/" & placement.standardVolume & ") - " & _
                           placement.pieceType & " - ���: " & placement.weightCategory
    
    ' ��� 1: ����� ����� � ������������ ����������� ����� ������
    Dim bestWarehouse As String
    bestWarehouse = FindWarehouseWithMaxProduct(placement.ProductName)
    
    If bestWarehouse <> "" Then
        ModuleLogger.LogSuccess "?? ������ ����� ��� '" & placement.ProductName & "': " & bestWarehouse
        
        ' ��� 2: ����������� ���������� � ������ ������ (������ �� �������)
        If TryPlaceInBestWarehouseCellByCell(placement, bestWarehouse) Then
            placement.IsPlaced = True
            SmartPlaceWithCellByCell = True
            Exit Function
        End If
    End If
    
    ' ��� 3: ���� �� ������� � ������ ������, ���� � ������ �������� (������ �� �������)
    If TryPlaceInAnyActiveWarehouseCellByCell(placement) Then
        placement.IsPlaced = True
        SmartPlaceWithCellByCell = True
        Exit Function
    End If
    
    ' �� ������� ����������
    placement.ErrorMessage = "�� ������� ����� � ������ �������� ������ ������"
    placement.IsPlaced = False
    SmartPlaceWithCellByCell = False
    Exit Function
    
ErrorHandler:
    placement.ErrorMessage = "������ ���������� ������-��-�������: " & Err.description
    SmartPlaceWithCellByCell = False
End Function

' ===== ���������� � ������ ������ (������ �� �������) =====
Private Function TryPlaceInBestWarehouseCellByCell(ByRef placement As PlacementInfo, warehouse As String) As Boolean
    ModuleLogger.LogMessage "?? ���������� � ������ ������ " & warehouse & " (������-��-�������)"
    
    ' ���� ������������ ����� � ���� ������
    Dim existingProduct As FoundProductInfo
    existingProduct = FindProductInWarehouseHeaders(placement.ProductName, warehouse)
    
    If existingProduct.startRow > 0 Then
        ModuleLogger.LogSuccess "? ����� ������ � �����: " & existingProduct.availableRows & " - ��������� ������ ������"
        
        ' ������� ������ ������������ ���� � ������� ������ �� �������
        If TryPlaceInExistingRowsCellByCell(placement, existingProduct, warehouse) Then
            TryPlaceInBestWarehouseCellByCell = True
            Exit Function
        End If
        
        ' ���� �������� ������ ����
        ModuleLogger.LogMessage "?? ���� �������� ������ ����..."
        If FindAdjacentEmptyRowsCellByCell(placement, warehouse, existingProduct) Then
            ModuleLogger.LogSuccess "? ��������� � �������� ���� (������-��-�������)!"
            TryPlaceInBestWarehouseCellByCell = True
            Exit Function
        End If
    End If
    
    ' ���� ������ ����
    If FindEmptyRowCellByCell(placement, warehouse) Then
        TryPlaceInBestWarehouseCellByCell = True
        Exit Function
    End If
    
    TryPlaceInBestWarehouseCellByCell = False
End Function

' ===== ���������� � ������������ ����� (������ �� �������) =====
Private Function TryPlaceInExistingRowsCellByCell(ByRef placement As PlacementInfo, _
                                                 existingProduct As FoundProductInfo, _
                                                 warehouse As String) As Boolean
    
    ModuleLogger.LogMessage "?? ���������� ������-��-������� � ������������ �����"
    
    ' ��������� ������ �����
    Dim rows() As String
    rows = Split(existingProduct.availableRows, ",")
    
    ' ������� ������ ��� � ������� ������-��-�������
    Dim i As Integer
    For i = 0 To UBound(rows)
        Dim rowNum As String
        rowNum = Trim(rows(i))
        
        If TryPlaceInRowCellByCell(placement, warehouse, rowNum, existingProduct.section) Then
            TryPlaceInExistingRowsCellByCell = True
            Exit Function
        End If
    Next i
    
    TryPlaceInExistingRowsCellByCell = False
End Function

' ===== ?? ������� �������: ���������� � ���� (������ �� �������) =====
Private Function TryPlaceInRowCellByCell(ByRef placement As PlacementInfo, warehouse As String, _
                                        rowNum As String, section As String) As Boolean
    On Error GoTo ErrorHandler
    
    ModuleLogger.LogMessage "?? ���������� ������-��-������� � ���� " & rowNum & " ������ " & ModuleConfig.GetSectionName(section)
    
    ' �������� ������������ ������
    Dim config As ModuleTypes.WarehouseConfig
    config = ModuleConfig.GetWarehouseConfig(warehouse)
    
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(config.sheetName)
    
    ' ��������� ����� ����
    Dim currentRowTotal As Double
    currentRowTotal = CalculateRowTotalWithVolumes(ws, rowNum, section, warehouse)
    
    Dim rowLimit As Double
    rowLimit = GetRowLimit(warehouse, placement.standardVolume)
    
    If currentRowTotal + placement.quantity > rowLimit Then
        ModuleLogger.LogWarning "? ��� " & rowNum & " ����������: " & (currentRowTotal + placement.quantity) & " > " & rowLimit
        TryPlaceInRowCellByCell = False
        Exit Function
    End If
    
    ' ?? ���� ������ �����: �1>�2>�3>�1>�2>�3...
    Dim exactLocation As String
    exactLocation = FindExactCellForPlacement(ws, warehouse, rowNum, section, placement)
    
    If exactLocation = "" Then
        ModuleLogger.LogDebug "? ��� ��������� ����� � ���� " & rowNum
        TryPlaceInRowCellByCell = False
        Exit Function
    End If
    
    ' ������ ������ �����
    Dim locationParts() As String
    locationParts = Split(exactLocation, "-")
    
    If UBound(locationParts) < 1 Then
        ModuleLogger.LogError "? ������ �������� �����: " & exactLocation
        TryPlaceInRowCellByCell = False
        Exit Function
    End If
    
    ' ��������� ����� � ������ �����
    placement.warehouse = warehouse
    placement.row = rowNum
    placement.Shelf = locationParts(0)
    placement.level = locationParts(1)
    
    If PlaceInExactCell(placement, section) Then
        placement.placementDetails = "����� " & placement.warehouse & "-" & placement.row & "-" & _
                                   placement.Shelf & "-" & placement.level & " - " & _
                                   placement.ProductName & " (" & placement.Batch & ") - " & _
                                   placement.quantity & " "
        
        ModuleLogger.LogSuccess "?? ��������� �����: " & placement.placementDetails
        TryPlaceInRowCellByCell = True
    Else
        TryPlaceInRowCellByCell = False
    End If
    
    Exit Function
    
ErrorHandler:
    ModuleLogger.LogError "������ ���������� ������-��-�������: " & Err.description
    TryPlaceInRowCellByCell = False
End Function

' ===== ?? �������� �������: ����� ������� ����� �� ������� =====
Private Function FindExactCellForPlacement(ws As Worksheet, warehouse As String, rowNum As String, _
                                         section As String, placement As PlacementInfo) As String
    ' ���������� ������ ����� � ������� "�-2" (�����-����)
    
    ModuleLogger.LogMessage "?? ����� ����� ������-��-������� � ���� " & rowNum
    
    ' �������� ������ ���� � ���������� �������
    Dim letters() As String
    letters = GetLettersWithCorrectDirection(warehouse, section, placement.pieceType)
    
    ' �������� ������������
    Dim config As ModuleTypes.WarehouseConfig
    config = ModuleConfig.GetWarehouseConfig(warehouse)
    
    Dim dataStart As Long, dataEnd As Long
    If section = "UPPER" Then
        dataStart = config.UpperDataStart
        dataEnd = config.UpperDataEnd
    Else
        dataStart = config.LowerDataStart
        dataEnd = config.LowerDataEnd
    End If
    
    Dim valueColumn As Long
    valueColumn = 3 + (CLng(rowNum) - 1) * 2
    
    ' ?? ��������� ������ ����� � ������ ����
    Dim i As Integer
    For i = 0 To UBound(letters)
        Dim currentLetter As String
        currentLetter = letters(i)
        
        ModuleLogger.LogDebug "?? ��������� ����� " & currentLetter & " �� ������..."
        
        ' ����������� ��� ����� ��������
        Dim letterAnalysis As shelfAnalysis
        letterAnalysis = AnalyzeSingleShelfDetailed(ws, currentLetter, valueColumn, dataStart, dataEnd, _
                                                   warehouse, placement.standardVolume)
        
        ' ?? ���� ������ ����� �� ������: 1>2>3
        Dim exactCell As String
        exactCell = FindExactLevelInShelf(letterAnalysis, placement.quantity, placement.weightCategory)
        
        If exactCell <> "" Then
            ModuleLogger.LogSuccess "? ������� �����: " & currentLetter & "-" & exactCell
            FindExactCellForPlacement = currentLetter & "-" & exactCell
            Exit Function
        End If
    Next i
    
    ModuleLogger.LogWarning "? ��� ��������� ����� � ���� " & rowNum
    FindExactCellForPlacement = ""
End Function

' ===== ?? ��������� ������ ����� ����� =====
Private Function AnalyzeSingleShelfDetailed(ws As Worksheet, shelfLetter As String, valueColumn As Long, _
                                           dataStart As Long, dataEnd As Long, warehouse As String, _
                                           productVolume As Double) As shelfAnalysis
    
    Dim analysis As shelfAnalysis
    analysis.Letter = shelfLetter
    analysis.ShelfLimit = GetShelfLimit(warehouse, productVolume)
    analysis.totalQuantity = 0
    analysis.Level1Quantity = 0
    analysis.Level2Quantity = 0
    analysis.Level3Quantity = 0
    analysis.HasProduct = False
    
    ModuleLogger.LogDebug "?? ��������� ������ ����� " & shelfLetter & " (�����: " & analysis.ShelfLimit & ")"
    
    ' ��������� ������ ����: 1, 2, 3
    Dim level As Integer
    For level = 1 To 3
        Dim searchText As String
        searchText = ModuleConfig.GetLevelText(shelfLetter, CStr(level))
        
        Dim foundCell As Range
        Set foundCell = ws.Range("A" & dataStart & ":B" & dataEnd).Find( _
            What:=searchText, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not foundCell Is Nothing Then
            Dim cellValue As Double
            cellValue = GetNumericValueSafe(ws.Cells(foundCell.row, valueColumn))
            
            ' ���������� ���������� �� ������
            Select Case level
                Case 1:
                    analysis.Level1Quantity = cellValue
                    ModuleLogger.LogDebug "   ���� 1: " & cellValue
                Case 2:
                    analysis.Level2Quantity = cellValue
                    ModuleLogger.LogDebug "   ���� 2: " & cellValue
                Case 3:
                    analysis.Level3Quantity = cellValue
                    ModuleLogger.LogDebug "   ���� 3: " & cellValue
            End Select
            
            analysis.totalQuantity = analysis.totalQuantity + cellValue
            If cellValue > 0 Then analysis.HasProduct = True
        End If
    Next level
    
    ModuleLogger.LogDebug "?? ����� " & shelfLetter & " �����: " & analysis.totalQuantity & "/" & analysis.ShelfLimit
    
    AnalyzeSingleShelfDetailed = analysis
End Function

' ===== ?? ����� ������� ����� � ����� =====
Private Function FindExactLevelInShelf(shelfAnalysis As shelfAnalysis, quantity As Double, _
                                      weightCategory As String) As String
    
    ' ���������, ���������� �� ����� � ��� �����
    If shelfAnalysis.totalQuantity + quantity > shelfAnalysis.ShelfLimit Then
        ModuleLogger.LogDebug "? ����� " & shelfAnalysis.Letter & " �����������: " & _
                             (shelfAnalysis.totalQuantity + quantity) & " > " & shelfAnalysis.ShelfLimit
        FindExactLevelInShelf = ""
        Exit Function
    End If
    
    ' ?? ������������ �����: 1>2>3
    
    ' ���� 1 (���)
    If shelfAnalysis.Level1Quantity = 0 Then
        ModuleLogger.LogDebug "? ���� 1 ��������"
        FindExactLevelInShelf = "1"
        Exit Function
    End If
    
    ' ���� 2 (�������)
    If shelfAnalysis.Level2Quantity = 0 And shelfAnalysis.Level1Quantity > 0 Then
        If weightCategory = "HEAVY" Then
            ModuleLogger.LogWarning "?? ������� �� ���� 2 - ��������� ������������"
        End If
        ModuleLogger.LogDebug "? ���� 2 ��������"
        FindExactLevelInShelf = "2"
        Exit Function
    End If
    
    ' ���� 3 (����)
    If shelfAnalysis.Level3Quantity = 0 And shelfAnalysis.Level2Quantity > 0 And shelfAnalysis.Level1Quantity > 0 Then
        If weightCategory = "HEAVY" Then
            ModuleLogger.LogWarning "? ������� ������ �� ���� 3"
            FindExactLevelInShelf = ""
            Exit Function
        End If
        ModuleLogger.LogDebug "? ���� 3 ��������"
        FindExactLevelInShelf = "3"
        Exit Function
    End If
    
    ' ��� ����� ������
    ModuleLogger.LogDebug "? ��� ����� � ����� " & shelfAnalysis.Letter & " ������"
    FindExactLevelInShelf = ""
End Function

' ===== ���������� � ������ ������ =====
Private Function PlaceInExactCell(ByRef placement As PlacementInfo, section As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim config As ModuleTypes.WarehouseConfig
    config = ModuleConfig.GetWarehouseConfig(placement.warehouse)
    
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(config.sheetName)
    
    Dim dataStart As Long, dataEnd As Long
    If section = "UPPER" Then
        dataStart = config.UpperDataStart
        dataEnd = config.UpperDataEnd
    Else
        dataStart = config.LowerDataStart
        dataEnd = config.LowerDataEnd
    End If
    
    Dim searchText As String
    searchText = ModuleConfig.GetLevelText(placement.Shelf, placement.level)
    
    Dim targetCell As Range
    Set targetCell = ws.Range("A" & dataStart & ":B" & dataEnd).Find( _
        What:=searchText, LookIn:=xlValues, LookAt:=xlWhole)
    
    If targetCell Is Nothing Then
        placement.ErrorMessage = "�� ������� ������ " & searchText
        PlaceInExactCell = False
        Exit Function
    End If
    
    Dim valueColumn As Long
    valueColumn = 3 + (CLng(placement.row) - 1) * 2
    
    Dim batchColumn As Long
    batchColumn = valueColumn + 1
    
    ' ���������� ������
    ws.Cells(targetCell.row, valueColumn).value = placement.quantity
    ws.Cells(targetCell.row, batchColumn).value = placement.Batch
    
    ' ?? �������� ����� ��� ������-��-�������
    Dim highlightColor As Long
    Select Case placement.pieceType
        Case "WHOLE"
            ' ����� ������ (��������, 720 �� 720)
            highlightColor = RGB(0, 176, 80)   ' ���������� �������
    
        Case "MEDIUM"
            ' ������� ����� (��������, 650 �� 720)
            highlightColor = RGB(146, 208, 80) ' ������-�������
    
        Case "PIECE"
            ' ����� (��������, 150 �� 720)
            highlightColor = RGB(255, 255, 0)  ' ������
    
        Case Else
            ' ���� �� ���������, ���� ��� �� ���������
            highlightColor = RGB(217, 217, 217) ' �����
    End Select
    
    ' ����������� ���������
    If placement.pieceType = "PIECE" Then
        ' ������� ������� ��� ������
        With ws.Cells(targetCell.row, valueColumn).Borders
            .LineStyle = xlContinuous
            .Weight = xlThick
            .Color = RGB(255, 0, 0)
        End With
        With ws.Cells(targetCell.row, batchColumn).Borders
            .LineStyle = xlContinuous
            .Weight = xlThick
            .Color = RGB(255, 0, 0)
        End With
    Else
        ' ����� ������� ��� ����������� ���������� ������-��-�������
        With ws.Cells(targetCell.row, valueColumn).Borders
            .LineStyle = xlContinuous
            .Weight = xlThick
            .Color = RGB(0, 100, 200)
        End With
        With ws.Cells(targetCell.row, batchColumn).Borders
            .LineStyle = xlContinuous
            .Weight = xlThick
            .Color = RGB(0, 100, 200)
        End With
    End If
    
    ws.Cells(targetCell.row, valueColumn).Interior.Color = highlightColor
    ws.Cells(targetCell.row, batchColumn).Interior.Color = highlightColor
    
    ModuleLogger.LogSuccess "?? ������-��-������� � " & targetCell.address & " ���� " & placement.level & _
                          " [" & placement.pieceType & "/" & placement.weightCategory & "] " & _
                          placement.Shelf & "-" & placement.level
    
    PlaceInExactCell = True
    Exit Function
    
ErrorHandler:
    placement.ErrorMessage = "������ ���������� � ������: " & Err.description
    PlaceInExactCell = False
End Function

' ===== ��������������� ������� (��������� �� ��������� �2) =====

Private Function TryPlaceInAnyActiveWarehouseCellByCell(ByRef placement As PlacementInfo) As Boolean
    ModuleLogger.LogMessage "?? ����� ����� � ����� �������� ������ (������-��-�������)"
    
    ' �������� �� ���� �������� �������
    Dim warehouse As Integer
    For warehouse = 5 To 12
        If ModuleWarehouseCapacity.IsWarehouseActive(CStr(warehouse)) Then
            
            If TryPlaceInBestWarehouseCellByCell(placement, CStr(warehouse)) Then
                TryPlaceInAnyActiveWarehouseCellByCell = True
                Exit Function
            End If
        End If
    Next warehouse
    
    TryPlaceInAnyActiveWarehouseCellByCell = False
End Function

Private Function FindAdjacentEmptyRowsCellByCell(ByRef placement As PlacementInfo, warehouse As String, _
                                                existingProduct As FoundProductInfo) As Boolean
    On Error GoTo ErrorHandler

    ModuleLogger.LogMessage "?? ����� �������� ������ ����� (������-��-�������)"

    ' ������ ������������ ����
    Dim existingRows() As String
    existingRows = Split(existingProduct.availableRows, ",")

    ' ������� ����������� � ������������ ����
    Dim minRow As Integer, maxRow As Integer
    minRow = 999
    maxRow = 0

    Dim i As Integer
    For i = 0 To UBound(existingRows)
        Dim currentRowNum As Integer
        currentRowNum = CInt(Trim(existingRows(i)))

        If currentRowNum < minRow Then minRow = currentRowNum
        If currentRowNum > maxRow Then maxRow = currentRowNum
    Next i

    ModuleLogger.LogMessage "?? �������� ������������ �����: " & minRow & " - " & maxRow

    ' �������� ������������ ������
    Dim config As ModuleTypes.WarehouseConfig
    config = ModuleConfig.GetWarehouseConfig(warehouse)

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(config.sheetName)

    Dim targetSection As String
    targetSection = existingProduct.section

    Dim headerRow As Long
    If targetSection = "UPPER" Then
        headerRow = config.UpperHeaderRow
    Else
        headerRow = config.LowerHeaderRow
    End If

    ' ��������� �������� ����
    Dim distance As Integer
    For distance = 1 To 10
        ' ������ �� �������������
        Dim rightRow As Integer
        rightRow = maxRow + distance
        If rightRow >= 1 And rightRow <= config.maxRows Then
            If IsRowCompletelyEmpty(ws, CStr(rightRow), headerRow, targetSection, warehouse) Then
                ModuleLogger.LogSuccess "? ������ �������� ������ ���: " & rightRow
                
                If CreateNewHeaderAndPlaceCellByCell(placement, warehouse, CStr(rightRow), targetSection) Then
                    FindAdjacentEmptyRowsCellByCell = True
                    Exit Function
                End If
            End If
        End If
        
        ' ����� �� ������������
        Dim leftRow As Integer
        leftRow = minRow - distance
        If leftRow >= 1 And leftRow <= config.maxRows Then
            If IsRowCompletelyEmpty(ws, CStr(leftRow), headerRow, targetSection, warehouse) Then
                ModuleLogger.LogSuccess "? ������ �������� ������ ���: " & leftRow
                
                If CreateNewHeaderAndPlaceCellByCell(placement, warehouse, CStr(leftRow), targetSection) Then
                    FindAdjacentEmptyRowsCellByCell = True
                    Exit Function
                End If
            End If
        End If
    Next distance
    
    FindAdjacentEmptyRowsCellByCell = False
    Exit Function
    
ErrorHandler:
    ModuleLogger.LogError "������ ������ �������� �����: " & Err.description
    FindAdjacentEmptyRowsCellByCell = False
End Function

Private Function FindEmptyRowCellByCell(ByRef placement As PlacementInfo, warehouse As String) As Boolean
    ModuleLogger.LogMessage "?? ����� ������� ���� � ������ " & warehouse & " (������-��-�������)"

    ' �������� ������������ ������
    Dim config As ModuleTypes.WarehouseConfig
    config = ModuleConfig.GetWarehouseConfig(warehouse)

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(config.sheetName)

    ' ���� � ������� ������
    If config.UpperHeaderRow > 0 Then
        Dim r As Integer
        For r = 1 To config.maxRows
            If IsRowCompletelyEmpty(ws, CStr(r), config.UpperHeaderRow, "UPPER", warehouse) Then
                If CreateNewHeaderAndPlaceCellByCell(placement, warehouse, CStr(r), "UPPER") Then
                    FindEmptyRowCellByCell = True
                    Exit Function
                End If
            End If
        Next r
    End If

    ' ���� � ������ ������
    If config.LowerHeaderRow > 0 Then
        For r = 1 To config.maxRows
            If IsRowCompletelyEmpty(ws, CStr(r), config.LowerHeaderRow, "LOWER", warehouse) Then
                If CreateNewHeaderAndPlaceCellByCell(placement, warehouse, CStr(r), "LOWER") Then
                    FindEmptyRowCellByCell = True
                    Exit Function
                End If
            End If
        Next r
    End If

    FindEmptyRowCellByCell = False
End Function

Private Function CreateNewHeaderAndPlaceCellByCell(ByRef placement As PlacementInfo, warehouse As String, _
                                                  rowNum As String, section As String) As Boolean
    ModuleLogger.LogMessage "?? ������� ��������� ��� '" & placement.ProductName & "' � ���� " & rowNum & " (������-��-�������)"
    
    ' �������� ������������
    Dim config As ModuleTypes.WarehouseConfig
    config = ModuleConfig.GetWarehouseConfig(warehouse)
    
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(config.sheetName)
    
    Dim headerRow As Long
    If section = "UPPER" Then
        headerRow = config.UpperHeaderRow
    Else
        headerRow = config.LowerHeaderRow
    End If
    
    If headerRow = 0 Then
        CreateNewHeaderAndPlaceCellByCell = False
        Exit Function
    End If
    
    ' ������� ���������
    Call CreateMergedHeader(ws, placement.ProductName, rowNum, headerRow)
    
    ' ��������� ����� ������� ������-��-�������
    If TryPlaceInRowCellByCell(placement, warehouse, rowNum, section) Then
        ModuleLogger.LogSuccess "? ������ ��������� � �������� ����� (������-��-�������)"
        CreateNewHeaderAndPlaceCellByCell = True
    Else
        CreateNewHeaderAndPlaceCellByCell = False
    End If
End Function

' ===== ��������� ����������� ������� �� ��������� �2 =====

Private Function IsRowCompletelyEmpty(ws As Worksheet, rowNum As String, headerRow As Long, _
                                    section As String, warehouse As String) As Boolean
    ' ��������� ���������
    If headerRow > 0 Then
        Dim headerCol As Long
        headerCol = 3 + (CLng(rowNum) - 1) * 2
        
        Dim headerValue As String
        headerValue = GetCellValueSafe(ws.Cells(headerRow, headerCol))
        
        If Trim(headerValue) <> "" Then
            IsRowCompletelyEmpty = False
            Exit Function
        End If
    End If
    
    ' ��������� ������
    Dim totalInRow As Double
    totalInRow = CalculateRowTotalWithVolumes(ws, rowNum, section, warehouse)
    
    IsRowCompletelyEmpty = (totalInRow = 0)
End Function

Private Sub CreateMergedHeader(ws As Worksheet, ProductName As String, rowNum As String, headerRow As Long)
    On Error Resume Next
    
    Dim startCol As Long, endCol As Long
    startCol = 3 + (CLng(rowNum) - 1) * 2
    endCol = startCol + 1
    
    ws.Range(ws.Cells(headerRow, startCol), ws.Cells(headerRow, endCol)).UnMerge
    ws.Range(ws.Cells(headerRow, startCol), ws.Cells(headerRow, endCol)).ClearContents
    ws.Range(ws.Cells(headerRow, startCol), ws.Cells(headerRow, endCol)).Merge
    
    ws.Cells(headerRow, startCol).value = ProductName
    
    With ws.Cells(headerRow, startCol)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.Bold = True
        .Font.Size = 12
        .Interior.Color = RGB(200, 255, 200)
        .Font.Color = RGB(0, 100, 0)
        
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(100, 100, 100)
        End With
    End With
    
    On Error GoTo 0
End Sub

Private Function GetLettersWithCorrectDirection(warehouse As String, section As String, pieceType As String) As String()
    Dim allLetters() As String
    allLetters = GetSectionLetters(warehouse, section) ' �������� �����, �������� [���, �, �, ..., �] ��� ������

    Dim isLowerSectionPassageFirst As Boolean
    ' ���������, �������� �� ������ ����� � ������� �������� ��� ������ ������
    isLowerSectionPassageFirst = (section = "LOWER" And Left(allLetters(0), 2) = "��")

    Dim needsReversing As Boolean
    needsReversing = False ' �� ��������� ������ �� ��������������

    Select Case pieceType
        Case "WHOLE", "MEDIUM" ' ��� ����� � ������� �����
            ' ���� ��� ������ ������ � ����� ���������� � ������� (���, ���...),
            ' �� ������ ����� �����������, ����� ������ �� �����.
            If isLowerSectionPassageFirst Then
                needsReversing = True
                ModuleLogger.LogDebug "?? WHOLE/LOWER: �����>������ (������): " & Join(ReverseArray(allLetters), ">")
            Else
                ModuleLogger.LogDebug "?? WHOLE/UPPER: �����>������: " & Join(allLetters, ">")
            End If

        Case "PIECE" ' ��� ������
            ' ���� ��� ������� ������ (����� �� ���������� � �������),
            ' �� ������ ����� �����������, ����� ������ � �������.
            If Not isLowerSectionPassageFirst Then
                needsReversing = True
                ModuleLogger.LogDebug "?? PIECE/UPPER: ������>����� (������): " & Join(ReverseArray(allLetters), ">")
            Else
                ModuleLogger.LogDebug "?? PIECE/LOWER: ������>�����: " & Join(allLetters, ">")
            End If
    End Select

    If needsReversing Then
        GetLettersWithCorrectDirection = ReverseArray(allLetters)
    Else
        GetLettersWithCorrectDirection = allLetters
    End If
End Function


Private Function ReverseArray(arr() As String) As String()
    Dim reversed() As String
    ReDim reversed(UBound(arr))
    
    Dim i As Integer
    For i = 0 To UBound(arr)
        reversed(i) = arr(UBound(arr) - i)
    Next i
    
    ReverseArray = reversed
End Function

Private Function FindWarehouseWithMaxProduct(ProductName As String) As String
    ModuleLogger.LogMessage "?? ���� ����� � ������������ ����������� '" & ProductName & "'"
    
    Dim maxQuantity As Double
    maxQuantity = 0
    
    Dim bestWarehouse As String
    bestWarehouse = ""
    
    ' ��������� ��� �������� ������
    Dim warehouse As Integer
    For warehouse = 5 To 12
        If ModuleWarehouseCapacity.IsWarehouseActive(CStr(warehouse)) Then
            
            Dim quantity As Double
            quantity = CalculateTotalProductInWarehouse(ProductName, CStr(warehouse))
            
            If quantity > maxQuantity Then
                maxQuantity = quantity
                bestWarehouse = CStr(warehouse)
            End If
            
            ModuleLogger.LogDebug "����� " & warehouse & ": " & quantity & "� ������ '" & ProductName & "'"
        End If
    Next warehouse
    
    If bestWarehouse <> "" Then
        ModuleLogger.LogSuccess "?? �������� � ������ " & bestWarehouse & ": " & maxQuantity & "�"
    Else
        ModuleLogger.LogMessage "?? ����� �� ������ (����� �����)"
    End If
    
    FindWarehouseWithMaxProduct = bestWarehouse
End Function

Private Function CalculateTotalProductInWarehouse(ProductName As String, warehouse As String) As Double
    On Error Resume Next
    
    ' �������� ������������ ������
    Dim config As ModuleTypes.WarehouseConfig
    config = ModuleConfig.GetWarehouseConfig(warehouse)
    
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(config.sheetName)
    
    If ws Is Nothing Then
        CalculateTotalProductInWarehouse = 0
        Exit Function
    End If
    
    Dim totalQuantity As Double
    totalQuantity = 0
    
    ' ���� ����� � ���������� � ��������� ����������
    totalQuantity = totalQuantity + SumProductInSection(ws, ProductName, config.UpperHeaderRow, config.UpperDataStart, config.UpperDataEnd, "UPPER")
    totalQuantity = totalQuantity + SumProductInSection(ws, ProductName, config.LowerHeaderRow, config.LowerDataStart, config.LowerDataEnd, "LOWER")
    
    CalculateTotalProductInWarehouse = totalQuantity
    On Error GoTo 0
End Function

Private Function SumProductInSection(ws As Worksheet, ProductName As String, headerRow As Long, _
                                    dataStart As Long, dataEnd As Long, sectionName As String) As Double
    On Error Resume Next
    
    If headerRow = 0 Then
        SumProductInSection = 0
        Exit Function
    End If
    
    Dim totalInSection As Double
    totalInSection = 0
    
    ' �������� �� ���� ���������� � ������
    Dim col As Long
    Dim lastCheckedCol As Long
    lastCheckedCol = 0
    
    For col = 3 To 127
        If col > lastCheckedCol Then
            
            Dim headerCell As Range
            Set headerCell = ws.Cells(headerRow, col)
            
            Dim headerValue As String
            headerValue = GetCellValueSafe(headerCell)
            
            If headerValue <> "" Then
                ' ��������� ���������� ������
                If UCase(Trim(headerValue)) = UCase(Trim(ProductName)) Then
                    ' ���������� �������� ���������
                    Dim headerStartCol As Long, headerEndCol As Long
                    If headerCell.MergeCells Then
                        headerStartCol = headerCell.MergeArea.Column
                        headerEndCol = headerCell.MergeArea.Column + headerCell.MergeArea.columns.count - 1
                        lastCheckedCol = headerEndCol
                    Else
                        headerStartCol = col
                        headerEndCol = col
                        lastCheckedCol = col
                    End If
                    
                    ' ��������� ���������� � ��������������� �����
                    Dim c As Long
                    For c = headerStartCol To headerEndCol Step 2
                        totalInSection = totalInSection + SumColumnInRange(ws, c, dataStart, dataEnd)
                    Next c
                End If
            End If
        End If
    Next col
    
    SumProductInSection = totalInSection
    On Error GoTo 0
End Function

Private Function SumColumnInRange(ws As Worksheet, col As Long, startRow As Long, endRow As Long) As Double
    On Error Resume Next
    
    Dim total As Double
    total = 0
    
    Dim r As Long
    For r = startRow To endRow
        Dim cellValue As Double
        cellValue = GetNumericValueSafe(ws.Cells(r, col))
        total = total + cellValue
    Next r
    
    SumColumnInRange = total
    On Error GoTo 0
End Function

Private Function FindProductInWarehouseHeaders(ProductName As String, warehouse As String) As FoundProductInfo
    Dim result As FoundProductInfo
    
    ' �������� ������������ ������
    Dim config As ModuleTypes.WarehouseConfig
    config = ModuleConfig.GetWarehouseConfig(warehouse)
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(config.sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        FindProductInWarehouseHeaders = result
        Exit Function
    End If
    
    ' ���� � ������� ������
    result = SearchProductInSectionHeaders(ws, ProductName, config.UpperHeaderRow, "UPPER")
    If result.startRow > 0 Then
        result.warehouse = warehouse
        FindProductInWarehouseHeaders = result
        Exit Function
    End If
    
    ' ���� � ������ ������
    result = SearchProductInSectionHeaders(ws, ProductName, config.LowerHeaderRow, "LOWER")
    If result.startRow > 0 Then
        result.warehouse = warehouse
    End If
    
    FindProductInWarehouseHeaders = result
End Function

' ===== ������������ ������: ���� ��� ���� � �������, � �� ������ ���������� =====
Private Function SearchProductInSectionHeaders(ws As Worksheet, ProductName As String, _
                                              headerRow As Long, sectionName As String) As FoundProductInfo
    Dim result As FoundProductInfo
    result.section = sectionName
    result.headerRow = headerRow
    result.availableRows = "" ' �������������� ������ �������

    If headerRow = 0 Then
        SearchProductInSectionHeaders = result
        Exit Function
    End If

    Dim col As Long
    Dim lastCheckedCol As Long
    lastCheckedCol = 0

    For col = 3 To 127 ' �������� �� ���� ��������� ��������
        If col > lastCheckedCol Then
            Dim headerCell As Range
            Set headerCell = ws.Cells(headerRow, col)
            
            Dim headerValue As String
            headerValue = GetCellValueSafe(headerCell)

            If headerValue <> "" Then
                ' ��������� ���������� ������ (���������� UCase ��� ����������)
                If UCase(Trim(headerValue)) = UCase(Trim(ProductName)) Then
                    ' ���������� �������� �������� ��� ����� ���������
                    Dim headerStartCol As Long, headerEndCol As Long
                    If headerCell.MergeCells Then
                        headerStartCol = headerCell.MergeArea.Column
                        headerEndCol = headerCell.MergeArea.Column + headerCell.MergeArea.columns.count - 1
                    Else
                        headerStartCol = col
                        headerEndCol = col
                    End If

                    ' ���������� ������ �����, ������� ��������� ���� ���������
                    Dim startR As Integer, endR As Integer
                    startR = ((headerStartCol - 3) \ 2) + 1
                    endR = ((headerEndCol - 2) \ 2) ' ���� endR = ((headerEndCol - 3) \ 2) + 1, ���������� ��� ����������� ���������

                    ' ��������� ��� ��������� ���� � ����� ������
                    Dim r As Integer
                    For r = startR To endR
                        If result.availableRows <> "" Then
                            result.availableRows = result.availableRows & ","
                        End If
                        result.availableRows = result.availableRows & CStr(r)
                    Next r
                    
                    ' ������������� startRow, ����� ��������, ��� ����� ������
                    If result.startRow = 0 Then result.startRow = startR
                    
                    ' �� ������� �� ����� (Exit Function), � ���������� �����
                End If

                ' ���������� ����������� �������, ����� �� ����������� �� ������
                If headerCell.MergeCells Then
                    lastCheckedCol = headerCell.MergeArea.Column + headerCell.MergeArea.columns.count - 1
                Else
                    lastCheckedCol = col
                End If
            End If
        End If
    Next col

    ' ���� ���� ���� �������, ��������� �� ��� ��������� ������� ���������
    ' ���� ���� ���� �������, ��������� �� ��� ��������� ������� ���������
    If result.availableRows <> "" Then
        Dim rowsToSort() As String
        
        ' 1. ������� ��������� ������ �� ������
        rowsToSort = Split(result.availableRows, ",")
        
        ' 2. ����� �������� ���� ������ � ������� ����������
        Dim sortedRows() As String
        sortedRows = SortStringArrayNumerically(rowsToSort)
        
        ' 3. �������� ��������������� ������ ������� � ������
        result.availableRows = Join(sortedRows, ",")
        
        ModuleLogger.LogDebug "?? �������� ������ ����� ��� '" & ProductName & "': " & result.availableRows
    End If


    SearchProductInSectionHeaders = result
End Function

' ��������������� ������� ��� ���������� ������� ����� ��� �����
Private Function SortStringArrayNumerically(arr() As String) As String()
    Dim i As Long, j As Long
    Dim temp As String
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If CLng(arr(i)) > CLng(arr(j)) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
    SortStringArrayNumerically = arr
End Function


Private Function CalculateRowTotalWithVolumes(ws As Worksheet, rowNum As String, section As String, warehouse As String) As Double
    Dim config As ModuleTypes.WarehouseConfig
    config = ModuleConfig.GetWarehouseConfig(warehouse)
    
    Dim valueColumn As Long
    valueColumn = 3 + (CLng(rowNum) - 1) * 2
    
    Dim dataStart As Long, dataEnd As Long
    If section = "UPPER" Then
        dataStart = config.UpperDataStart
        dataEnd = config.UpperDataEnd
    Else
        dataStart = config.LowerDataStart
        dataEnd = config.LowerDataEnd
    End If
    
    Dim total As Double
    total = 0
    
    Dim row As Long
    For row = dataStart To dataEnd
        Dim cellValue As Double
        cellValue = GetNumericValueSafe(ws.Cells(row, valueColumn))
        total = total + cellValue
    Next row
    
    CalculateRowTotalWithVolumes = total
End Function

Private Function GetSectionLetters(warehouse As String, section As String) As String()
    Dim letters() As String
    
    Select Case warehouse
        Case "5", "6"
            If section = "UPPER" Then
                letters = Split("�,�,�,�,�,�,�,�,���", ",")
            Else
                letters = Split("���,�,�,�,�,�,�,�,�", ",")
            End If
            
        Case "7", "8"
            If section = "UPPER" Then
                letters = Split("�,�,�,�,�,�,���", ",")
            Else
                letters = Split("���,�,�,�,�,�,�", ",")
            End If
            
        Case "9"
            If section = "UPPER" Then
                letters = Split("�,�,�,�,�,�,�,�,�,���", ",")
            Else
                letters = Split("���,�,�,�,�,�,�,�,�,�", ",")
            End If
            
        Case "10"
            If section = "UPPER" Then
                letters = Split("�,�,�,�,�,�,�,�,�,���", ",")
            Else
                letters = Split("���,�,�,�,�,�,�,�,�,�", ",")
            End If
            
        Case "11"
            If section = "UPPER" Then
                letters = Split("�,�,�,�,�,�,�,���", ",")
            Else
                letters = Split("���,�,�,�,�,�,�,�", ",")
            End If
            
        Case "12"
            If section = "UPPER" Then
                letters = Split("�,�,�,�,�,�,�,���", ",")
            Else
                letters = Split("���,�,�,�,�,�,�,�", ",")
            End If
            
        Case Else
            letters = Split("�,�,�,�,�,�,�,�", ",")
    End Select
    
    GetSectionLetters = letters
End Function

Private Function GetCellValueSafe(cell As Range) As String
    On Error Resume Next
    If cell.MergeCells Then
        GetCellValueSafe = Trim(CStr(cell.MergeArea.Cells(1, 1).value))
    Else
        GetCellValueSafe = Trim(CStr(cell.value))
    End If
    If GetCellValueSafe = "False" Then GetCellValueSafe = ""
    On Error GoTo 0
End Function

Private Function GetNumericValueSafe(cell As Range) As Double
    On Error Resume Next
    Dim value As Variant
    
    If cell.MergeCells Then
        value = cell.MergeArea.Cells(1, 1).value
    Else
        value = cell.value
    End If
    
    If IsNumeric(value) And Not isEmpty(value) Then
        GetNumericValueSafe = CDbl(value)
    Else
        GetNumericValueSafe = 0
    End If
    On Error GoTo 0
End Function

' ===== ������� ������� ������ =====
Private Function ParseNewFormatInput(inputText As String) As PlacementInfo()
    Dim lines() As String
    lines = Split(inputText, vbNewLine)
    
    Dim results() As PlacementInfo
    ReDim results(UBound(lines))
    
    Dim validCount As Integer
    validCount = 0
    
    Dim i As Integer
    For i = 0 To UBound(lines)
        Dim line As String
        line = Trim(lines(i))
        
        If line <> "" And Not (InStr(LCase(line), "������") > 0) Then
            Dim placement As PlacementInfo
            
            ' ������: "����� - ������ - ���������� (�����)"
            If ParseNewFormatLine(line, placement) Then
                ' ���������� ��������� ����
                placement.weightCategory = DetermineWeightCategory(placement.quantity, placement.standardVolume)
                
                results(validCount) = placement
                validCount = validCount + 1
                
                ModuleLogger.LogDebug "?? ����������: " & placement.ProductName & " - " & _
                                     placement.Batch & " - " & placement.quantity & " (" & _
                                     placement.standardVolume & ") - " & placement.pieceType & _
                                     " - ���: " & placement.weightCategory
            End If
        End If
    Next i
    
    If validCount > 0 Then
        ReDim Preserve results(validCount - 1)
    Else
        ReDim results(-1)
    End If
    
    ParseNewFormatInput = results
End Function

Private Function ParseNewFormatLine(line As String, ByRef placement As PlacementInfo) As Boolean
    On Error GoTo ErrorHandler
    
    ' ������: "�������� - ���13 - 600 (720)"
    Dim parts() As String
    parts = Split(line, " - ")
    
    If UBound(parts) < 2 Then
        ParseNewFormatLine = False
        Exit Function
    End If
    
    ' �������� �������� �����
    placement.ProductName = Trim(parts(0))
    placement.Batch = Trim(parts(1))
    
    ' ������ ���������� � �����
    Dim lastPart As String
    lastPart = Trim(parts(2))
    
    ' ���� ����� � �������
    Dim openBracket As Integer, closeBracket As Integer
    openBracket = InStr(lastPart, "(")
    closeBracket = InStr(lastPart, ")")
    
    If openBracket > 0 And closeBracket > openBracket Then
        ' ��������� ����������
        Dim quantityStr As String
        quantityStr = Trim(Left(lastPart, openBracket - 1))
        
        ' ��������� �����
        Dim volumeStr As String
        volumeStr = Trim(Mid(lastPart, openBracket + 1, closeBracket - openBracket - 1))
        
        If IsNumeric(quantityStr) And IsNumeric(volumeStr) Then
            placement.quantity = CDbl(quantityStr)
            placement.standardVolume = CDbl(volumeStr)
            
            ' ���������� ��� ������
            placement.pieceType = DeterminePieceType(placement.quantity, placement.standardVolume)
            
            ParseNewFormatLine = True
        End If
    Else
        ' ������ ��� ������
        If IsNumeric(lastPart) Then
            placement.quantity = CDbl(lastPart)
            placement.standardVolume = 720 ' �� ���������
            placement.pieceType = DeterminePieceType(placement.quantity, placement.standardVolume)
            
            ParseNewFormatLine = True
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    ParseNewFormatLine = False
End Function

Private Function DeterminePieceType(quantity As Double, standardVolume As Double) As String
    If standardVolume = 0 Then
        DeterminePieceType = "WHOLE"
        Exit Function
    End If
    
    Dim percentage As Double
    percentage = (quantity / standardVolume) * 100
    
    If percentage >= 90 Then
        DeterminePieceType = "WHOLE"   ' ����� �������
    ElseIf percentage >= 60 Then
        DeterminePieceType = "MEDIUM"  ' �������
    Else
        DeterminePieceType = "PIECE"   ' �����
    End If
End Function

' ===== ��������� ������ "������ �� �������" (���������� �������� �������) =====
Private Function GenerateCellByCellReport(placements() As PlacementInfo, successCount As Integer, activeCount As Integer) As String
    Dim report As String
    report = "?? ����� ���������� '������ �� �������'" & vbNewLine
    report = report & "===============================================" & vbNewLine
    report = report & "?? " & Format(Now, "dd.mm.yyyy hh:mm:ss") & vbNewLine
    report = report & "?? �������� ������: " & ModuleWarehouseCapacity.GetActiveWarehousesList() & vbNewLine & vbNewLine
    
    report = report & "? �������: " & successCount & vbNewLine
    report = report & "? ������: " & (UBound(placements) + 1 - successCount) & vbNewLine & vbNewLine
    
    If successCount > 0 Then
        report = report & "?? ���������� ������-��-�������:" & vbNewLine
        report = report & "-----------------------------------------" & vbNewLine
        
        Dim i As Integer
        For i = 0 To UBound(placements)
            If placements(i).IsPlaced And placements(i).placementDetails <> "" Then
                report = report & placements(i).placementDetails & vbNewLine
            End If
        Next i
        report = report & vbNewLine
    End If
    
    Dim errorCount As Integer
    errorCount = (UBound(placements) + 1 - successCount)
    
    If errorCount > 0 Then
        report = report & "? ������:" & vbNewLine
        report = report & "-------------" & vbNewLine
        
        Dim j As Integer
        For j = 0 To UBound(placements)
            If Not placements(j).IsPlaced Then
                report = report & "� " & placements(j).ProductName & " (" & _
                        placements(j).Batch & ") - " & _
                        placements(j).quantity & " - " & _
                        placements(j).ErrorMessage & vbNewLine
            End If
        Next j
        report = report & vbNewLine
    End If
    
    report = report & "?? �����������:" & vbNewLine
    report = report & "=========================" & vbNewLine
    report = report & "? �������� ������ ������:" & vbNewLine
    report = report & "   � �1>�2>�3>�1>�2>�3 (���������!)" & vbNewLine
    report = report & "   � �� �1>�1>�1>�1 (���� �����������)" & vbNewLine & vbNewLine
    
    report = report & "? ����� ������ �� ������:" & vbNewLine
    report = report & "   � 720�: �� 2 ���� (1440/1500)" & vbNewLine
    report = report & "   � 960�: ������ 1 ����� (960/1000)" & vbNewLine
    report = report & "   � 240��: �� 2 ���� (480/500)" & vbNewLine & vbNewLine
    
    report = report & "?? �������� ���������:" & vbNewLine
    report = report & "   � ?? �����-������� - ���� 1" & vbNewLine
    report = report & "   � ?? ����-������� - ���� 2" & vbNewLine
    report = report & "   � ?? �����-������� - ���� 3" & vbNewLine
    report = report & "   � ?? ����� ������� - ���������� ����������" & vbNewLine
    report = report & "   � ?? ������� ������� - �����" & vbNewLine & vbNewLine
    
    report = report & "? ������� '������-��-�������' �������� ���������!"
    
    GenerateCellByCellReport = report
End Function

' ===== �������� ������� =====
Public Sub TestCellByCell()
    Call ModuleLogger.InitializeLogger(True, ModuleLogger.LOG_LEVEL_DEBUG)
    
    ModuleLogger.LogMessage "=== ?? ���� ������� '������ �� �������' ==="
    
    ' �������� ������: 3 ������� �� 720�
    Dim testData As String
    testData = "�������� - ���01 - 720 (720)" & vbNewLine & _
               "�������� - ���02 - 720 (720)" & vbNewLine & _
               "�������� - ���03 - 720 (720)"
    
    ModuleLogger.LogMessage "������� ������: 3?720�"
    ModuleLogger.LogMessage "��������� ���������:"
    ModuleLogger.LogMessage "  � �-1: 720 (������ ������)"
    ModuleLogger.LogMessage "  � �-2: 720 (������ ������ � ��� �� ����� �)"
    ModuleLogger.LogMessage "  � �-1: 720 (������ ������ � ����� ����� �, ����� � ��������)"
    
    Dim result As String
    result = SmartPlaceProducts(testData)
    
    MsgBox "?? ���� '������ �� �������' ��������!" & vbNewLine & vbNewLine & _
           "��������� ���������:" & vbNewLine & _
           "? �-1: 720 (�����-������� + ����� �������)" & vbNewLine & _
           "? �-2: 720 (����-������� + ����� �������)" & vbNewLine & _
           "? �-1: 720 (�����-������� + ����� �������)" & vbNewLine & vbNewLine & _
           "?? ���������: " & vbNewLine & result, _
           vbInformation, "������� '������-��-�������' ��������!"
End Sub

