VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SimpleSmartForm2 
   Caption         =   "UserForm1"
   ClientHeight    =   6840
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7695
   OleObjectBlob   =   "SimpleSmartForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SimpleSmartForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===== SimpleSmartForm - ИСПРАВЛЕННАЯ РАБОЧАЯ ВЕРСИЯ =====
Option Explicit

Private lastReport As String  ' Для хранения последнего отчета

Private Sub UserForm_Initialize()
    ' === РАЗМЕР ФОРМЫ ===
    Me.Width = 600
    Me.Height = 650
    Me.Caption = "Умное размещение товаров"
    
    ' === 1. ИНСТРУКЦИЯ (Label1) ===
    With Label1
        .Caption = "Введите товары БЕЗ указания места (формат: Товар - Партия - Количество):"
        .Top = 10
        .Left = 10
        .Width = 570
        .Height = 30
        .Font.Size = 12
        .Font.Bold = True
        .WordWrap = True
    End With
    
    ' === 2. ТЕКСТОВОЕ ПОЛЕ ДЛЯ ВВОДА (TextBox1) ===
    With TextBox1
        .MultiLine = True
        .ScrollBars = fmScrollBarsBoth
        .Top = 50
        .Left = 10
        .Width = 570
        .Height = 120
        .Font.Name = "Consolas"
        .Font.Size = 11
        ' Пример для пользователя
        .text = "Лерашанс - пар13 - 720" & vbNewLine & _
                "Босфор - пар07 - 500" & vbNewLine & _
                "Дикошанс - пар25 - 360"
    End With
    
    ' === 3. ЗАГОЛОВОК АНГАРОВ (Label2) ===
    With Label2
        .Caption = "?? Активные ангары для размещения:"
        .Top = 180
        .Left = 10
        .Width = 570
        .Height = 20
        .Font.Size = 11
        .Font.Bold = True
        .ForeColor = RGB(0, 100, 0)
    End With
    
    ' === 4. ГАЛОЧКИ АНГАРОВ (CheckBox1-8) ===
    ' Первый ряд: Ангары 5,6,7,8
    With CheckBox1  ' Ангар 5
        .Caption = "Ангар 5"
        .Top = 210
        .Left = 15
        .Width = 80
        .Height = 20
        .value = True
    End With
    
    With CheckBox2  ' Ангар 6
        .Caption = "Ангар 6"
        .Top = 210
        .Left = 105
        .Width = 80
        .Height = 20
        .value = True
    End With
    
    With CheckBox3  ' Ангар 7
        .Caption = "Ангар 7"
        .Top = 210
        .Left = 195
        .Width = 80
        .Height = 20
        .value = True
    End With
    
    With CheckBox4  ' Ангар 8
        .Caption = "Ангар 8"
        .Top = 210
        .Left = 285
        .Width = 80
        .Height = 20
        .value = True
    End With
    
    ' Второй ряд: Ангары 9,10,11,12
    With CheckBox5  ' Ангар 9
        .Caption = "Ангар 9"
        .Top = 240
        .Left = 15
        .Width = 80
        .Height = 20
        .value = True
    End With
    
    With CheckBox6  ' Ангар 10
        .Caption = "Ангар 10"
        .Top = 240
        .Left = 105
        .Width = 80
        .Height = 20
        .value = True
    End With
    
    With CheckBox7  ' Ангар 11
        .Caption = "Ангар 11"
        .Top = 240
        .Left = 195
        .Width = 80
        .Height = 20
        .value = True
    End With
    
    With CheckBox8  ' Ангар 12
        .Caption = "Ангар 12"
        .Top = 240
        .Left = 285
        .Width = 80
        .Height = 20
        .value = True
    End With
    
    ' === 5. КНОПКИ УПРАВЛЕНИЯ АНГАРАМИ ===
    With CommandButton4  ' Все
        .Caption = "Все"
        .Top = 240
        .Left = 380
        .Width = 50
        .Height = 20
        .BackColor = RGB(200, 255, 200)
    End With
    
    With CommandButton5  ' Нет
        .Caption = "Нет"
        .Top = 240
        .Left = 440
        .Width = 50
        .Height = 20
        .BackColor = RGB(255, 200, 200)
    End With
    
    ' === 6. ГЛАВНЫЕ КНОПКИ ===
    With CommandButton1  ' РАЗМЕСТИТЬ
        .Caption = "?? РАЗМЕСТИТЬ"
        .Top = 280
        .Left = 10
        .Width = 150
        .Height = 40
        .Font.Bold = True
        .Font.Size = 12
        .BackColor = RGB(0, 200, 0)
        .ForeColor = RGB(255, 255, 255)
    End With
    
    With CommandButton2  ' ОЧИСТИТЬ
        .Caption = "??? ОЧИСТИТЬ"
        .Top = 280
        .Left = 170
        .Width = 150
        .Height = 40
        .Font.Bold = True
        .Font.Size = 12
        .BackColor = RGB(255, 150, 0)
        .ForeColor = RGB(255, 255, 255)
    End With
    
    With CommandButton3  ' ОТЧЕТ
        .Caption = "?? ОТЧЕТ"
        .Top = 280
        .Left = 330
        .Width = 150
        .Height = 40
        .Font.Bold = True
        .Font.Size = 12
        .BackColor = RGB(0, 150, 255)
        .ForeColor = RGB(255, 255, 255)
        .Enabled = False
    End With
    
    ' === 7. СТАТУС (Label3) ===
    With Label3
        .Caption = "Готов к работе. Выберите ангары и введите товары."
        .Top = 340
        .Left = 10
        .Width = 570
        .Height = 60
        .Font.Size = 10
        .BackColor = RGB(240, 240, 240)
        .BorderStyle = fmBorderStyleSingle
        .TextAlign = fmTextAlignCenter
    End With
    
    ' === 8. ДОПОЛНИТЕЛЬНЫЕ КНОПКИ ===
    With CommandButton6  ' Настройки
        .Caption = "?? Настройки"
        .Top = 420
        .Left = 10
        .Width = 120
        .Height = 30
    End With
    
    With CommandButton7  ' Помощь
        .Caption = "? Помощь"
        .Top = 420
        .Left = 140
        .Width = 120
        .Height = 30
    End With
    
    With CommandButton8  ' Закрыть
        .Caption = "? Закрыть"
        .Top = 420
        .Left = 460
        .Width = 120
        .Height = 30
    End With
    
    ' === ИНИЦИАЛИЗАЦИЯ ===
    Call LoadWarehouseSettings
    Call UpdateWarehouseStatus
End Sub

' ===== ОСНОВНЫЕ ОБРАБОТЧИКИ =====

Private Sub CommandButton1_Click()  ' РАЗМЕСТИТЬ
    ' Проверка данных
    If Trim(TextBox1.text) = "" Then
        MsgBox "Введите товары для размещения!", vbExclamation
        TextBox1.SetFocus
        Exit Sub
    End If
    
    ' Проверка Excel
    If ActiveWorkbook Is Nothing Then
        MsgBox "Откройте файл с ангарами!", vbExclamation
        Exit Sub
    End If
    
    ' Проверка активных ангаров
    If GetActiveWarehouseCount() = 0 Then
        MsgBox "?? Не выбран ни один ангар для размещения!" & vbNewLine & _
               "Включите хотя бы один ангар.", vbExclamation
        Exit Sub
    End If
    
    ' Подтверждение с информацией об активных ангарах
    Dim activeList As String
    activeList = GetActiveWarehousesList()
    
    If MsgBox("Разместить товары в активных ангарах?" & vbNewLine & vbNewLine & _
              "Активные ангары: " & activeList, vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    
    ' Обновляем статус
    Label3.Caption = "? Выполняется размещение в " & GetActiveWarehouseCount() & " ангарах..."
    Label3.BackColor = RGB(255, 255, 200)
    DoEvents
    
    ' Инициализация систем
    Call ModuleProductInfo.InitializeProductDatabase
    Call ModuleWarehouseCapacity.InitializeWarehouseCapacities
    
    ' Синхронизируем состояние ангаров
    Call SyncWarehouseStates
    
    ' Выполняем размещение
    On Error GoTo ErrorHandler
    lastReport = ModuleSmartPlacement.SmartPlaceProducts(TextBox1.text)
    
    ' Показываем результат
    If InStr(lastReport, "УСПЕШНО") > 0 Then
        Label3.Caption = "? Размещение выполнено! Нажмите ОТЧЕТ для деталей."
        Label3.BackColor = RGB(200, 255, 200)
        CommandButton3.Enabled = True
        
        ' Показываем краткий результат
        Dim successCount As Integer
        successCount = CountSuccessful(lastReport)
        MsgBox "Успешно размещено: " & successCount & " товаров" & vbNewLine & _
               "В ангарах: " & activeList & vbNewLine & _
               "Ячейки помечены ЖЕЛТЫМ цветом", vbInformation
    Else
        Label3.Caption = "? Ошибка размещения"
        Label3.BackColor = RGB(255, 200, 200)
        MsgBox lastReport, vbExclamation
    End If
    
    Exit Sub
    
ErrorHandler:
    Label3.Caption = "? Ошибка: " & Err.description
    Label3.BackColor = RGB(255, 200, 200)
    MsgBox "Ошибка: " & Err.description, vbCritical
End Sub

Private Sub CommandButton2_Click()  ' ОЧИСТИТЬ
    TextBox1.text = ""
    Label3.Caption = "Поле очищено. Готов к вводу товаров."
    Label3.BackColor = RGB(240, 240, 240)
    CommandButton3.Enabled = False
    lastReport = ""
    TextBox1.SetFocus
End Sub

Private Sub CommandButton3_Click()  ' ОТЧЕТ
    If lastReport = "" Then
        MsgBox "Нет данных для отчета!", vbInformation
        Exit Sub
    End If
    
    ' Показываем отчет
    MsgBox lastReport, vbInformation, "Подробный отчет о размещении"
End Sub

Private Sub CommandButton4_Click()  ' ВСЕ
    CheckBox1.value = True   ' Ангар 5
    CheckBox2.value = True   ' Ангар 6
    CheckBox3.value = True   ' Ангар 7
    CheckBox4.value = True   ' Ангар 8
    CheckBox5.value = True   ' Ангар 9
    CheckBox6.value = True   ' Ангар 10
    CheckBox7.value = True   ' Ангар 11
    CheckBox8.value = True   ' Ангар 12
    Call UpdateWarehouseStatus
End Sub

Private Sub CommandButton5_Click()  ' НЕТ
    CheckBox1.value = False  ' Ангар 5
    CheckBox2.value = False  ' Ангар 6
    CheckBox3.value = False  ' Ангар 7
    CheckBox4.value = False  ' Ангар 8
    CheckBox5.value = False  ' Ангар 9
    CheckBox6.value = False  ' Ангар 10
    CheckBox7.value = False  ' Ангар 11
    CheckBox8.value = False  ' Ангар 12
    Call UpdateWarehouseStatus
End Sub

Private Sub CommandButton6_Click()  ' НАСТРОЙКИ
    Dim info As String
    info = "?? НАСТРОЙКИ АНГАРОВ" & vbNewLine & vbNewLine
    info = info & "Активные ангары: " & GetActiveWarehousesList() & vbNewLine
    info = info & "Общее количество: " & GetActiveWarehouseCount() & " из 8" & vbNewLine & vbNewLine
    info = info & "?? Настройки автоматически сохраняются" & vbNewLine
    info = info & "?? Для изменения лимитов используйте ModuleWarehouseCapacity"
    
    MsgBox info, vbInformation, "Информация о настройках"
End Sub

Private Sub CommandButton7_Click()  ' ПОМОЩЬ
    Dim help As String
    help = "? СПРАВКА ПО УМНОМУ РАЗМЕЩЕНИЮ" & vbNewLine & vbNewLine
    help = help & "?? ФОРМАТ ВВОДА:" & vbNewLine
    help = help & "Товар - Партия - Количество" & vbNewLine & vbNewLine
    help = help & "?? ПРИМЕРЫ:" & vbNewLine
    help = help & "Лерашанс - пар13 - 720" & vbNewLine
    help = help & "Босфор - пар07 - 500" & vbNewLine
    help = help & "Дикошанс - пар25 - 360" & vbNewLine & vbNewLine
    help = help & "?? ОСОБЕННОСТИ:" & vbNewLine
    help = help & "• Система сама найдет оптимальное место" & vbNewLine
    help = help & "• Учитывает индивидуальные лимиты ангаров" & vbNewLine
    help = help & "• Размещает только в выбранных ангарах" & vbNewLine
    help = help & "• Помечает новые ячейки желтым цветом"
    
    MsgBox help, vbInformation, "Помощь"
End Sub

Private Sub CommandButton8_Click()  ' ЗАКРЫТЬ
    If MsgBox("Закрыть форму умного размещения?", vbYesNo + vbQuestion) = vbYes Then
        Call SaveWarehouseSettings
        Unload Me
    End If
End Sub

' ===== ОБРАБОТЧИКИ ГАЛОЧЕК =====

Private Sub CheckBox1_Click()  ' Ангар 5
    Call UpdateWarehouseStatus
    Call ModuleWarehouseCapacity.SetWarehouseActive("5", CheckBox1.value)
End Sub

Private Sub CheckBox2_Click()  ' Ангар 6
    Call UpdateWarehouseStatus
    Call ModuleWarehouseCapacity.SetWarehouseActive("6", CheckBox2.value)
End Sub

Private Sub CheckBox3_Click()  ' Ангар 7
    Call UpdateWarehouseStatus
    Call ModuleWarehouseCapacity.SetWarehouseActive("7", CheckBox3.value)
End Sub

Private Sub CheckBox4_Click()  ' Ангар 8
    Call UpdateWarehouseStatus
    Call ModuleWarehouseCapacity.SetWarehouseActive("8", CheckBox4.value)
End Sub

Private Sub CheckBox5_Click()  ' Ангар 9
    Call UpdateWarehouseStatus
    Call ModuleWarehouseCapacity.SetWarehouseActive("9", CheckBox5.value)
End Sub

Private Sub CheckBox6_Click()  ' Ангар 10
    Call UpdateWarehouseStatus
    Call ModuleWarehouseCapacity.SetWarehouseActive("10", CheckBox6.value)
End Sub

Private Sub CheckBox7_Click()  ' Ангар 11
    Call UpdateWarehouseStatus
    Call ModuleWarehouseCapacity.SetWarehouseActive("11", CheckBox7.value)
End Sub

Private Sub CheckBox8_Click()  ' Ангар 12
    Call UpdateWarehouseStatus
    Call ModuleWarehouseCapacity.SetWarehouseActive("12", CheckBox8.value)
End Sub

' ===== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ =====

Private Function GetActiveWarehouseCount() As Integer
    Dim count As Integer
    count = 0
    
    If CheckBox1.value Then count = count + 1   ' Ангар 5
    If CheckBox2.value Then count = count + 1   ' Ангар 6
    If CheckBox3.value Then count = count + 1   ' Ангар 7
    If CheckBox4.value Then count = count + 1   ' Ангар 8
    If CheckBox5.value Then count = count + 1   ' Ангар 9
    If CheckBox6.value Then count = count + 1   ' Ангар 10
    If CheckBox7.value Then count = count + 1   ' Ангар 11
    If CheckBox8.value Then count = count + 1   ' Ангар 12
    
    GetActiveWarehouseCount = count
End Function

Private Function GetActiveWarehousesList() As String
    Dim list As String
    list = ""
    
    If CheckBox1.value Then list = list & "5, "
    If CheckBox2.value Then list = list & "6, "
    If CheckBox3.value Then list = list & "7, "
    If CheckBox4.value Then list = list & "8, "
    If CheckBox5.value Then list = list & "9, "
    If CheckBox6.value Then list = list & "10, "
    If CheckBox7.value Then list = list & "11, "
    If CheckBox8.value Then list = list & "12, "
    
    ' Убираем последнюю запятую
    If Len(list) > 2 Then
        list = Left(list, Len(list) - 2)
    End If
    
    GetActiveWarehousesList = list
End Function

Private Sub UpdateWarehouseStatus()
    Dim activeCount As Integer
    activeCount = GetActiveWarehouseCount()
    
    If activeCount = 0 Then
        Label3.Caption = "?? Не выбран ни один ангар! Выберите ангары для размещения."
        Label3.BackColor = RGB(255, 200, 200)
    ElseIf activeCount = 8 Then
        Label3.Caption = "? Все ангары активны - готов к размещению"
        Label3.BackColor = RGB(200, 255, 200)
    Else
        Label3.Caption = "?? Активно ангаров: " & activeCount & " из 8 (" & GetActiveWarehousesList() & ")"
        Label3.BackColor = RGB(200, 230, 255)
    End If
End Sub

Private Sub SyncWarehouseStates()
    Call ModuleWarehouseCapacity.SetWarehouseActive("5", CheckBox1.value)
    Call ModuleWarehouseCapacity.SetWarehouseActive("6", CheckBox2.value)
    Call ModuleWarehouseCapacity.SetWarehouseActive("7", CheckBox3.value)
    Call ModuleWarehouseCapacity.SetWarehouseActive("8", CheckBox4.value)
    Call ModuleWarehouseCapacity.SetWarehouseActive("9", CheckBox5.value)
    Call ModuleWarehouseCapacity.SetWarehouseActive("10", CheckBox6.value)
    Call ModuleWarehouseCapacity.SetWarehouseActive("11", CheckBox7.value)
    Call ModuleWarehouseCapacity.SetWarehouseActive("12", CheckBox8.value)
End Sub

Private Sub SaveWarehouseSettings()
    On Error Resume Next
    SaveSetting "SmartWarehouse", "ActiveWarehouses", "Warehouse5", IIf(CheckBox1.value, "1", "0")
    SaveSetting "SmartWarehouse", "ActiveWarehouses", "Warehouse6", IIf(CheckBox2.value, "1", "0")
    SaveSetting "SmartWarehouse", "ActiveWarehouses", "Warehouse7", IIf(CheckBox3.value, "1", "0")
    SaveSetting "SmartWarehouse", "ActiveWarehouses", "Warehouse8", IIf(CheckBox4.value, "1", "0")
    SaveSetting "SmartWarehouse", "ActiveWarehouses", "Warehouse9", IIf(CheckBox5.value, "1", "0")
    SaveSetting "SmartWarehouse", "ActiveWarehouses", "Warehouse10", IIf(CheckBox6.value, "1", "0")
    SaveSetting "SmartWarehouse", "ActiveWarehouses", "Warehouse11", IIf(CheckBox7.value, "1", "0")
    SaveSetting "SmartWarehouse", "ActiveWarehouses", "Warehouse12", IIf(CheckBox8.value, "1", "0")
End Sub

Private Sub LoadWarehouseSettings()
    On Error Resume Next
    CheckBox1.value = (GetSetting("SmartWarehouse", "ActiveWarehouses", "Warehouse5", "1") = "1")
    CheckBox2.value = (GetSetting("SmartWarehouse", "ActiveWarehouses", "Warehouse6", "1") = "1")
    CheckBox3.value = (GetSetting("SmartWarehouse", "ActiveWarehouses", "Warehouse7", "1") = "1")
    CheckBox4.value = (GetSetting("SmartWarehouse", "ActiveWarehouses", "Warehouse8", "1") = "1")
    CheckBox5.value = (GetSetting("SmartWarehouse", "ActiveWarehouses", "Warehouse9", "1") = "1")
    CheckBox6.value = (GetSetting("SmartWarehouse", "ActiveWarehouses", "Warehouse10", "1") = "1")
    CheckBox7.value = (GetSetting("SmartWarehouse", "ActiveWarehouses", "Warehouse11", "1") = "1")
    CheckBox8.value = (GetSetting("SmartWarehouse", "ActiveWarehouses", "Warehouse12", "1") = "1")
    Call SyncWarehouseStates
End Sub

Private Function CountSuccessful(report As String) As Integer
    Dim lines() As String
    lines = Split(report, vbNewLine)
    
    Dim count As Integer
    count = 0
    
    Dim i As Integer
    For i = 0 To UBound(lines)
        If InStr(lines(i), "?") > 0 Then
            count = count + 1
        End If
    Next i
    
    CountSuccessful = count
End Function
