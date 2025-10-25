Attribute VB_Name = "ModuleTypes"
' ===== ModuleTypes - ������������ ������ ��� ���������� =====
Option Explicit

' ===== �������� ���� ������ ������� =====

' ������������ ������ (���������������� �����������)
Public Type WarehouseConfig
    sheetName As String
    UpperDataStart As Long
    UpperDataEnd As Long
    UpperHeaderRow As Long
    LowerDataStart As Long
    LowerDataEnd As Long
    LowerHeaderRow As Long
    outputStartRow As Long
    RowNumbersToSkip As Variant
    maxRows As Integer         ' Максимальное количество рядов в ангаре
End Type

' �������� � ������
Public Type WarehouseOperation
    warehouse As String  ' ����� ������
    row As String       ' ����� ����
    Letter As String    ' ����� ������
    level As String     ' ������� (����)
    value As Double     ' ��������
    Batch As String     ' ����� ������
    ProductName As String ' �������� ������ (��������� ��� �������������)
End Type

' �������������� � ������
Public Type WarehouseLocation
    warehouse As String  ' ����� ������
    row As String       ' ����� ����
    Letter As String    ' ����� ������
    level As String     ' ������� (����)
End Type

' ��������� ��� ����������� ����������
Public Type HeaderGroup
    ProductName As String      ' �������� ������
    startRow As Integer       ' ��������� ��� ������
    endRow As Integer         ' �������� ��� ������
    StartColumn As Integer    ' ��������� ������� (C, E, G...)
    EndColumn As Integer      ' �������� ������� (D, F, H...)
    section As String         ' ������ (UPPER/LOWER)
    IsConflict As Boolean     ' ���� ��������� �������
    RowsInGroup As String     ' ������ ����� ��� ������� (��������: "3,5,6")
End Type

' ��������� ��� ������� ������������ ����������
Public Type ExistingHeader
    ProductName As String     ' �������� ������
    row As Integer           ' ����� ����
    StartColumn As Integer   ' ��������� �������
    EndColumn As Integer     ' �������� �������
    section As String        ' ������
    IsConflict As Boolean    ' ���� �� �������� (�������� "+")
End Type

' ===== �������������� ���� ��� ������������ ����������� =====

' ���������� � ������ (��� ������ ProductInfo)
Public Type ProductInfo
    ProductName As String
    volumeGroup As String    ' ������ ������ (1-7)
    description As String
    standardQuantity As Double
End Type

' ���������� � ���������� (��� ������ ����������)
Public Type PlacementResult
    warehouse As String      ' ����� ������
    row As String           ' ����� ����
    Shelf As String         ' ����� ��������
    level As String         ' �������
    ProductName As String   ' �������� ������
    Batch As String         ' ������
    quantity As Double      ' ����������
    IsPlaced As Boolean     ' ������� ���������?
    ErrorMessage As String  ' ��������� �� ������
    cellAddress As String   ' ����� ������ (��������: "D15")
End Type

' ������������ ������� ������
Public Type CapacityConfig
    warehouseNumber As String
    maxItemsPerShelf As Integer    ' �������� ������� �� �������
    maxQuantityPerRow As Double    ' �������� ���������� � ����
    isActive As Boolean           ' ������� �� �����
    Notes As String              ' �������������� �������
End Type

' ��������� ���������
Public Type ValidationResult
    IsValid As Boolean           ' ������ �� ���������
    ErrorMessage As String       ' ��������� �� ������
    WarningMessage As String     ' ��������������
    ValidatedData As String      ' ����������� ������
End Type

' ���������� ������
Public Type WarehouseStats
    warehouseNumber As String
    totalItems As Integer        ' ����� ���������� �������
    totalQuantity As Double      ' ����� ���������� ������
    occupiedShelves As Integer   ' ������� ��������
    emptyPositions As Integer    ' ��������� �������
    lastUpdateTime As Date       ' ����� ���������� ����������
End Type

' ===== ������� ��� ������ � ������ =====

' ������� ������ �������� ������
Public Function CreateEmptyOperation() As WarehouseOperation
    Dim emptyOp As WarehouseOperation
    emptyOp.warehouse = ""
    emptyOp.row = ""
    emptyOp.Letter = ""
    emptyOp.level = ""
    emptyOp.value = 0
    emptyOp.Batch = ""
    emptyOp.ProductName = ""
    CreateEmptyOperation = emptyOp
End Function

' ������� ������ ��������� ����������
Public Function CreateEmptyPlacement() As PlacementResult
    Dim emptyPlace As PlacementResult
    emptyPlace.warehouse = ""
    emptyPlace.row = ""
    emptyPlace.Shelf = ""
    emptyPlace.level = ""
    emptyPlace.ProductName = ""
    emptyPlace.Batch = ""
    emptyPlace.quantity = 0
    emptyPlace.IsPlaced = False
    emptyPlace.ErrorMessage = ""
    emptyPlace.cellAddress = ""
    CreateEmptyPlacement = emptyPlace
End Function

' ������� ������ ������������ ������
Public Function CreateEmptyConfig() As WarehouseConfig
    Dim emptyConfig As WarehouseConfig
    emptyConfig.sheetName = ""
    emptyConfig.UpperDataStart = 0
    emptyConfig.UpperDataEnd = 0
    emptyConfig.UpperHeaderRow = 0
    emptyConfig.LowerDataStart = 0
    emptyConfig.LowerDataEnd = 0
    emptyConfig.LowerHeaderRow = 0
    emptyConfig.outputStartRow = 0
    emptyConfig.RowNumbersToSkip = Array()
    emptyConfig.maxRows = 0
    CreateEmptyConfig = emptyConfig
End Function

' ���������, �������� �� �������� ��������
Public Function IsValidOperation(op As WarehouseOperation) As Boolean
    IsValidOperation = (op.warehouse <> "" And op.row <> "" And _
                       op.Letter <> "" And op.level <> "" And _
                       op.value > 0 And op.Batch <> "")
End Function

' ����������� �������� � �������� ���
Public Function FormatOperation(op As WarehouseOperation) As String
    FormatOperation = "����� " & op.warehouse & ", ��� " & op.row & _
                     ", ������� " & op.Letter & ", ������� " & op.level & _
                     " - " & op.ProductName & " (������: " & op.Batch & _
                     ", ����������: " & op.value & ")"
End Function

' ===== ��������� ������� =====

' �������������� ������
Public Const MIN_WAREHOUSE As Integer = 5
Public Const MAX_WAREHOUSE As Integer = 12

' ������������ ������ ���������
Public Const MAX_SHELF_LEVEL As Integer = 3

' ����������� ������ ������
Public Const UPPER_SECTION As String = "UPPER"
Public Const LOWER_SECTION As String = "LOWER"

' ����������� �����������
Public Const MAIN_DELIMITER As String = " - "
Public Const LOCATION_DELIMITER As String = "-"

' ===== ������� ��������� ����� =====

Public Function IsValidWarehouseNumber(warehouse As String) As Boolean
    Dim num As Integer
    On Error Resume Next
    num = CInt(warehouse)
    On Error GoTo 0
    IsValidWarehouseNumber = (num >= MIN_WAREHOUSE And num <= MAX_WAREHOUSE)
End Function

Public Function IsValidShelfLevel(level As String) As Boolean
    Dim num As Integer
    On Error Resume Next
    num = CInt(level)
    On Error GoTo 0
    IsValidShelfLevel = (num >= 1 And num <= MAX_SHELF_LEVEL)
End Function

Public Function IsValidBatchFormat(Batch As String) As Boolean
    ' ������ ������ ���������� � "���" � ��������� �����
    IsValidBatchFormat = (Len(Batch) >= 4 And Left(LCase(Batch), 3) = "���")
End Function

