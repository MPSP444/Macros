Attribute VB_Name = "ModuleColor"

' ===== ModuleColor =====
Option Explicit

Public Sub StartColorWarehouseForm()
    If ActiveWorkbook Is Nothing Then
        MsgBox "Откройте книгу Excel с листами ангаров перед запуском макроса!", vbExclamation
        Exit Sub
    End If
    
    ColorWarehouseForm.Show
End Sub

Public Function ParseColorInput(inputStr As String) As WarehouseLocation()
    Dim lines() As String
    lines = Split(inputStr, vbNewLine)
    
    Dim locations() As WarehouseLocation
    ReDim locations(UBound(lines))
    
    Dim locCount As Integer
    locCount = 0
    
    Dim i As Integer
    For i = 0 To UBound(lines)
        If Trim(lines(i)) <> "" And InStr(lines(i), "-") > 0 Then
Dim parts() As String
            parts = Split(Trim(lines(i)), "-")
            
            If UBound(parts) = 3 Then
                If IsNumeric(parts(0)) Then
                    With locations(locCount)
                        .warehouse = parts(0)
                        .row = parts(1)
                        .Letter = parts(2)
                        .level = parts(3)
                    End With
                    
                    locCount = locCount + 1
                End If
            End If
        End If
    Next i
    
    If locCount > 0 Then
        ReDim Preserve locations(locCount - 1)
    Else
        ReDim locations(0)
    End If
    
    ParseColorInput = locations
End Function

Public Function ColorCells(locations() As WarehouseLocation, locCount As Integer) As String
    Dim errorLog As String
    errorLog = ""
    
    If ActiveWorkbook Is Nothing Then
        errorLog = "Не найдена активная книга Excel!"
        ColorCells = errorLog
        Exit Function
    End If
    
    Application.ScreenUpdating = False
    
    Dim lilacColor As Long
    lilacColor = RGB(198, 70, 186) ' Лиловый цвет
    
    Dim i As Integer
    For i = 0 To locCount - 1
        Dim sheetName As String
        Dim upperStart As Integer, upperEnd As Integer
        Dim lowerStart As Integer, lowerEnd As Integer
        
        Call GetWarehouseRanges(locations(i).warehouse, sheetName, _
            upperStart, upperEnd, lowerStart, lowerEnd)
        
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = ActiveWorkbook.Sheets(sheetName)
        On Error GoTo 0
        
        If ws Is Nothing Then
            errorLog = errorLog & "Лист '" & sheetName & "' не найден!" & vbNewLine
            GoTo NextLocation
        End If
        
        Dim foundCell As Range
        Set foundCell = FindCellInWarehouse(ws, locations(i).Letter, _
            locations(i).level, upperStart, upperEnd, lowerStart, lowerEnd)
        
        If foundCell Is Nothing Then
            errorLog = errorLog & "Не найдена ячейка " & locations(i).Letter & " уровень " & _
                      locations(i).level & " в ангаре " & locations(i).warehouse & vbNewLine
            GoTo NextLocation
        End If
        
        Dim valueColumn As Long
        valueColumn = 3 + (CLng(locations(i).row) - 1) * 2
        
        Dim valueCell As Range
        Dim batchCell As Range
        Set valueCell = ws.Cells(foundCell.row, valueColumn)
        Set batchCell = ws.Cells(foundCell.row, valueColumn + 1)
        
        ' Красим текст в лиловый цвет
        valueCell.Font.Color = lilacColor
        If batchCell.value <> "" Then
            batchCell.Font.Color = lilacColor
        End If
        
NextLocation:
    Next i
    
    Application.ScreenUpdating = True
    ColorCells = errorLog
End Function

