Attribute VB_Name = "ЗакрепитьАнгары"
Option Explicit

' Процедуры форматирования для каждого ангара
Private Sub ФорматироватьАнгар5(ws As Worksheet)
    With ws
        .Activate
        ActiveWindow.FreezePanes = False
        .Range("C4").Select  ' Для закрепления строк 3,4 и столбцов A,B
        ActiveWindow.FreezePanes = True
    End With
End Sub

Private Sub ФорматироватьАнгар6(ws As Worksheet)
    With ws
        .Activate
        ActiveWindow.FreezePanes = False
        .Range("C4").Select  ' Для закрепления строк 2,3 и столбцов A,B
        ActiveWindow.FreezePanes = True
    End With
End Sub

Private Sub ФорматироватьАнгар7(ws As Worksheet)
    With ws
        .Activate
        ActiveWindow.FreezePanes = False
        .Range("C4").Select  ' Для закрепления строк 2,3 и столбцов A,B
        ActiveWindow.FreezePanes = True
    End With
End Sub

Private Sub ФорматироватьАнгар8(ws As Worksheet)
    With ws
        .Activate
        ActiveWindow.FreezePanes = False
        .Range("C4").Select  ' Для закрепления строк 2,3 и столбцов A,B
        ActiveWindow.FreezePanes = True
    End With
End Sub

Private Sub ФорматироватьАнгар9(ws As Worksheet)
    With ws
        .Activate
        ActiveWindow.FreezePanes = False
        .Range("C4").Select  ' Для закрепления строк 2,3 и столбцов A,B
        ActiveWindow.FreezePanes = True
    End With
End Sub

Private Sub ФорматироватьАнгар10(ws As Worksheet)
    With ws
        .Activate
        ActiveWindow.FreezePanes = False
        .Range("C4").Select  ' Для закрепления строк 2,3 и столбцов A,B
        ActiveWindow.FreezePanes = True
    End With
End Sub

Private Sub ФорматироватьАнгар11(ws As Worksheet)
    With ws
        .Activate
        ActiveWindow.FreezePanes = False
        .Range("C4").Select  ' Для закрепления строк 2,3 и столбцов A,B
        ActiveWindow.FreezePanes = True
    End With
End Sub

Private Sub ФорматироватьАнгар12(ws As Worksheet)
    With ws
        .Activate
        ActiveWindow.FreezePanes = False
        .Range("C4").Select  ' Для закрепления строк 2,3 и столбцов A,B
        ActiveWindow.FreezePanes = True
    End With
End Sub

Private Sub ФорматироватьМал10(ws As Worksheet)
    With ws
        .Activate
        ActiveWindow.FreezePanes = False
        .Range("C4").Select  ' Для закрепления строк 2,3 и столбцов A,B
        ActiveWindow.FreezePanes = True
    End With
End Sub

' Основная процедура для форматирования всех ангаров
Public Sub НастройкаВсехАнгаров()
    ' Отключаем обновление экрана
    Application.ScreenUpdating = False
    
    ' Проверяем наличие активной книги
    If ActiveWorkbook Is Nothing Then
        MsgBox "Не найдена активная книга Excel!", vbExclamation
        Exit Sub
    End If
    
    ' Сохраняем текущий активный лист
    Dim currentSheet As Worksheet
    Set currentSheet = activeSheet
    
    ' Инициализируем переменные
    Dim ws As Worksheet
    Dim errorLog As String
    errorLog = ""
    
    ' Обрабатываем каждый лист в активной книге
    For Each ws In ActiveWorkbook.Worksheets
        On Error Resume Next
        Select Case ws.Name
            Case "Ангар 5"
                ФорматироватьАнгар5 ws
                If Err.Number <> 0 Then
                    errorLog = errorLog & "Ошибка при форматировании Ангара 5: " & Err.description & vbNewLine
                End If
                
            Case "Ангар 6"
                ФорматироватьАнгар6 ws
                If Err.Number <> 0 Then
                    errorLog = errorLog & "Ошибка при форматировании Ангара 6: " & Err.description & vbNewLine
                End If
                
            Case "Ангар 7"
                ФорматироватьАнгар7 ws
                If Err.Number <> 0 Then
                    errorLog = errorLog & "Ошибка при форматировании Ангара 7: " & Err.description & vbNewLine
                End If
                
            Case "Ангар 8"
                ФорматироватьАнгар8 ws
                If Err.Number <> 0 Then
                    errorLog = errorLog & "Ошибка при форматировании Ангара 8: " & Err.description & vbNewLine
                End If
                
            Case "Ангар 9"
                ФорматироватьАнгар9 ws
                If Err.Number <> 0 Then
                    errorLog = errorLog & "Ошибка при форматировании Ангара 9: " & Err.description & vbNewLine
                End If
                
            Case "Ангар 10"
                ФорматироватьАнгар10 ws
                If Err.Number <> 0 Then
                    errorLog = errorLog & "Ошибка при форматировании Ангара 10: " & Err.description & vbNewLine
                End If
                
            Case "Ангар 11"
                ФорматироватьАнгар11 ws
                If Err.Number <> 0 Then
                    errorLog = errorLog & "Ошибка при форматировании Ангара 11: " & Err.description & vbNewLine
                End If
                
            Case "Ангар 12"
                ФорматироватьАнгар12 ws
                If Err.Number <> 0 Then
                    errorLog = errorLog & "Ошибка при форматировании Ангара 12: " & Err.description & vbNewLine
                End If
                
            Case "мал 10"
                ФорматироватьМал10 ws
                If Err.Number <> 0 Then
                    errorLog = errorLog & "Ошибка при форматировании мал 10: " & Err.description & vbNewLine
                End If
        End Select
        On Error GoTo 0
    Next ws
    
    ' Возвращаемся к исходному листу
    currentSheet.Activate
    
    ' Включаем обновление экрана
    Application.ScreenUpdating = True
    
    ' Показываем результат
    If errorLog = "" Then
        MsgBox "Форматирование всех ангаров завершено успешно!", vbInformation
    Else
        MsgBox "Форматирование завершено с ошибками:" & vbNewLine & vbNewLine & errorLog, vbExclamation
    End If
End Sub

' Дополнительная процедура для форматирования одного конкретного ангара
Public Sub ЗакрепитьОбластиОдногоАнгара()
    ' Получаем имя активного листа
    Dim wsName As String
    wsName = activeSheet.Name
    
    ' Проверяем, является ли лист ангаром
    Select Case wsName
        Case "Ангар 5"
            ФорматироватьАнгар5 activeSheet
        Case "Ангар 6"
            ФорматироватьАнгар6 activeSheet
        Case "Ангар 7"
            ФорматироватьАнгар7 activeSheet
        Case "Ангар 8"
            ФорматироватьАнгар8 activeSheet
        Case "Ангар 9"
            ФорматироватьАнгар9 activeSheet
        Case "Ангар 10"
            ФорматироватьАнгар10 activeSheet
        Case "Ангар 11"
            ФорматироватьАнгар11 activeSheet
        Case "Ангар 12"
            ФорматироватьАнгар12 activeSheet
        Case "мал 10"
            ФорматироватьМал10 activeSheet
        Case Else
            MsgBox "Текущий лист не является ангаром!", vbExclamation
    End Select
End Sub
