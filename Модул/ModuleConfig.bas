Attribute VB_Name = "ModuleConfig"

' ===== ModuleConfig =====
Option Explicit

Public Function GetWarehouseConfig(warehouseNumber As String) As ModuleTypes.WarehouseConfig
    Dim config As ModuleTypes.WarehouseConfig
    
    Select Case warehouseNumber
        Case "5"
            With config
                .sheetName = "����� 5"
                .UpperDataStart = 4
                .UpperDataEnd = 30
                .UpperHeaderRow = 2    ' ���������
                .LowerDataStart = 42
                .LowerDataEnd = 68
                .LowerHeaderRow = 70   ' ���������
                .outputStartRow = 115  ' ���������
                .RowNumbersToSkip = Array(3, 69)
                .maxRows = 45          ' Максимум 45 рядов (до CN)
            End With
            
        Case "6"
            With config
                .sheetName = "����� 6"
                .UpperDataStart = 4
                .UpperDataEnd = 30
                .UpperHeaderRow = 2
                .LowerDataStart = 44
                .LowerDataEnd = 70
                .LowerHeaderRow = 72
                .outputStartRow = 74
                .RowNumbersToSkip = Array(3, 71)
                .maxRows = 48          ' Максимум 48 рядов (до CT)
            End With
            
        Case "7"
            With config
                .sheetName = "����� 7"
                .UpperDataStart = 4
                .UpperDataEnd = 24
                .UpperHeaderRow = 2
                .LowerDataStart = 36
                .LowerDataEnd = 56
                .LowerHeaderRow = 58
                .outputStartRow = 60
                .RowNumbersToSkip = Array(3, 57)
                .maxRows = 50          ' Максимум 50 рядов (до CX)
            End With
            
        Case "8"
            With config
                .sheetName = "����� 8"
                .UpperDataStart = 4
                .UpperDataEnd = 24
                .UpperHeaderRow = 2
                .LowerDataStart = 35
                .LowerDataEnd = 55
                .LowerHeaderRow = 57
                .outputStartRow = 115
                .RowNumbersToSkip = Array(3, 56)
                .maxRows = 51          ' Максимум 51 ряд (до CZ)
            End With
            
        Case "9"
            With config
                .sheetName = "����� 9"
                .UpperDataStart = 4
                .UpperDataEnd = 33
                .UpperHeaderRow = 2
                .LowerDataStart = 50
                .LowerDataEnd = 79
                .LowerHeaderRow = 81
                .outputStartRow = 83
                .RowNumbersToSkip = Array(3, 80)
                .maxRows = 48          ' Максимум 48 рядов (до CT)
            End With
            
        Case "10"
            With config
                .sheetName = "����� 10"
                .UpperDataStart = 4
                .UpperDataEnd = 34
                .UpperHeaderRow = 2
                .LowerDataStart = 48
                .LowerDataEnd = 78
                .LowerHeaderRow = 80
                .outputStartRow = 82
                .RowNumbersToSkip = Array(3, 79)
                .maxRows = 40          ' Максимум 40 рядов (до CD)
            End With
            
        Case "11"
            With config
                .sheetName = "����� 11"
                .UpperDataStart = 4
                .UpperDataEnd = 27
                .UpperHeaderRow = 2
                .LowerDataStart = 42
                .LowerDataEnd = 65
                .LowerHeaderRow = 67
                .outputStartRow = 69
                .RowNumbersToSkip = Array(3, 66)
                .maxRows = 48          ' Максимум 48 рядов (до CT)
            End With
            
        Case "12"
            With config
                .sheetName = "����� 12"
                .UpperDataStart = 4
                .UpperDataEnd = 27
                .UpperHeaderRow = 2
                .LowerDataStart = 43
                .LowerDataEnd = 66
                .LowerHeaderRow = 68
                .outputStartRow = 115
                .RowNumbersToSkip = Array(3, 67)
                .maxRows = 53          ' Максимум 53 ряда (до DD)
            End With
    End Select
    
    GetWarehouseConfig = config
End Function

Public Function GetLevelText(Letter As String, level As String) As String
    Letter = UCase(Letter)
    Select Case level
        Case "1": GetLevelText = Letter & "/���.���"
        Case "2": GetLevelText = Letter & "/2���"
        Case "3": GetLevelText = Letter & "/3���"
        Case Else: GetLevelText = Letter & "/" & level
    End Select
End Function

Public Sub GetWarehouseRanges(ByVal warehouseNumber As String, _
                            ByRef sheetName As String, _
                            ByRef upperStart As Integer, _
                            ByRef upperEnd As Integer, _
                            ByRef lowerStart As Integer, _
                            ByRef lowerEnd As Integer)
    Dim config As ModuleTypes.WarehouseConfig  ' ���������� �����
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

' ===== ����� ������� ��� ������ � �������� =====

Public Function DetermineSectionByShelf(Shelf As String, warehouse As String) As String
    ' ���������� ������ (UPPER/LOWER) �� ����� �������� � ������ ������
    Dim upperShelves As String
    Dim lowerShelves As String
    
    ' �������� ������ ���� ��� ������� ������
    Select Case warehouse
        Case "5", "6"
            upperShelves = "�,�,�,�,�,�,�,�,���"
            lowerShelves = "���,�,�,�,�,�,�,�,�"
            
        Case "7", "8"
            upperShelves = "�,�,�,�,�,�,���"
            lowerShelves = "���,�,�,�,�,�,�"
            
        Case "9"
            upperShelves = "�,�,�,�,�,�,�,�,�,���"
            lowerShelves = "���,�,�,�,�,�,�,�,�,�"
            
        Case "10"
            upperShelves = "�,�,�,�,�,�,�,�,�,���"
            lowerShelves = "���,�,�,�,�,�,�,�,�,�"
            
        Case "11"
            upperShelves = "�,�,�,�,�,�,�,���"
            lowerShelves = "���,�,�,�,�,�,�,�"
            
        Case "12"
            upperShelves = "�,�,�,�,�,�,�,���"
            lowerShelves = "���,�,�,�,�,�,�,�"
            
        Case Else
            DetermineSectionByShelf = "UPPER" ' �� ���������
            Exit Function
    End Select
    
    ' ��������� �������������� � ������
    If InStr("," & upperShelves & ",", "," & UCase(Shelf) & ",") > 0 Then
        DetermineSectionByShelf = "UPPER"
    ElseIf InStr("," & lowerShelves & ",", "," & UCase(Shelf) & ",") > 0 Then
        DetermineSectionByShelf = "LOWER"
    Else
        DetermineSectionByShelf = "UPPER" ' �� ���������, ���� �� �������
    End If
End Function

Public Function FindCellInWarehouseSection(ws As Worksheet, _
                                         Letter As String, _
                                         level As String, _
                                         warehouse As String) As Range
    ' ���������� ������ ������ ������ � ������ ������
    Dim config As ModuleTypes.WarehouseConfig
    config = GetWarehouseConfig(warehouse)
    
    Dim section As String
    section = DetermineSectionByShelf(Letter, warehouse)
    
    Dim searchRange As Range
    Dim searchText As String
    searchText = GetLevelText(Letter, level)
    
    ' ���� ������ � ������ ������
    If section = "UPPER" Then
        Set searchRange = ws.Range("A" & config.UpperDataStart & ":B" & config.UpperDataEnd)
    Else
        Set searchRange = ws.Range("A" & config.LowerDataStart & ":B" & config.LowerDataEnd)
    End If
    
    ' ����� � ������ ��������
    Dim foundCell As Range
    Set foundCell = searchRange.Find(What:=searchText, _
                                   LookIn:=xlValues, _
                                   LookAt:=xlWhole, _
                                   SearchOrder:=xlByRows, _
                                   SearchDirection:=xlNext, _
                                   MatchCase:=False)
                                   
    If foundCell Is Nothing Then
        ' ������� ������ �������
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
    ' ���������� �������� �������� ������
    Select Case UCase(section)
        Case "UPPER": GetSectionName = "�������"
        Case "LOWER": GetSectionName = "������"
        Case Else: GetSectionName = "�����������"
    End Select
End Function

Public Function GetWarehouseSectionsList(warehouse As String) As String
    ' ���������� ������ ���� � ������� ��� �������
    Dim result As String
    result = "����� " & warehouse & ":" & vbNewLine
    
    Select Case warehouse
        Case "5", "6"
            result = result & "�������: �,�,�,�,�,�,�,�,���" & vbNewLine
            result = result & "������: ���,�,�,�,�,�,�,�,�"
            
        Case "7", "8"
            result = result & "�������: �,�,�,�,�,�,���" & vbNewLine
            result = result & "������: ���,�,�,�,�,�,�"
            
        Case "9"
            result = result & "�������: �,�,�,�,�,�,�,�,�,���" & vbNewLine
            result = result & "������: ���,�,�,�,�,�,�,�,�,�"
            
        Case "10"
            result = result & "�������: �,�,�,�,�,�,�,�,�,���" & vbNewLine
            result = result & "������: ���,�,�,�,�,�,�,�,�,�"
            
        Case "11"
            result = result & "�������: �,�,�,�,�,�,�,���" & vbNewLine
            result = result & "������: ���,�,�,�,�,�,�,�"
            
        Case "12"
            result = result & "�������: �,�,�,�,�,�,�,���" & vbNewLine
            result = result & "������: ���,�,�,�,�,�,�,�"
            
        Case Else
            result = result & "����������� �����!"
    End Select
    
    GetWarehouseSectionsList = result
End Function

