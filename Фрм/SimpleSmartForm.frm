VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SimpleSmartForm 
   Caption         =   "UserForm4"
   ClientHeight    =   5760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11520
   OleObjectBlob   =   "SimpleSmartForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SimpleSmartForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===== SimpleSmartForm - ПОЛНАЯ ВЕРСИЯ С КНОПКОЙ ПЕЧАТИ =====
Option Explicit

Private lastReport As String  ' Для хранения последнего отчета

' Вместо пользовательского типа используем массивы строк
' Индексы: 0=ProductName, 1=Batch, 2=Quantity, 3=Location

' ===== НОВЫЙ ОБРАБОТЧИК КНОПКИ ПЕЧАТИ =====
Private Sub CommandButton10_Click()
    If lastReport = "" Then
        MsgBox "Нет данных для печати!", vbInformation
        Exit Sub
    End If
    
    ' Печатаем красивый отчет
    PrintBeautifulReport lastReport
End Sub

Private Sub CommandButton9_Click()
    If lastReport = "" Then
        MsgBox "Нет данных для отчета!", vbInformation
        Exit Sub
    End If
    
    ' Показываем полный отчет
    ShowReportForm lastReport

End Sub

Private Sub ShowReportForm(reportText As String)
    On Error GoTo ErrorHandler
    
    ' Фиксированный путь для сохранения
    Dim filePath As String
    filePath = "C:\Users\mukam\Desktop\Остатки_ШТ_ШЭ\Размещение.txt"
    
    ' Проверяем существование папки
    Dim folderPath As String
    folderPath = "C:\Users\mukam\Desktop\Остатки_ШТ_ШЭ"
    
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If
    
    ' Добавляем дату и время к отчету
    Dim fullReport As String
    fullReport = "=== ОТЧЕТ О РАЗМЕЩЕНИИ ТОВАРОВ ===" & vbNewLine
    fullReport = fullReport & "Дата: " & Format(Now, "dd.mm.yyyy hh:mm:ss") & vbNewLine
    fullReport = fullReport & String(50, "=") & vbNewLine & vbNewLine
    fullReport = fullReport & reportText
    
    ' Сохраняем в файл
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, fullReport
    Close #fileNum
    
    ' Открываем файл в Блокноте
    Shell "notepad.exe " & filePath, vbNormalFocus
    
    MsgBox "Отчет сохранен:" & vbNewLine & filePath & vbNewLine & vbNewLine & _
           "Файл открыт в Блокноте - можете скопировать нужный текст", vbInformation, "Отчет сохранен"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка сохранения отчета:" & vbNewLine & Err.description & vbNewLine & vbNewLine & _
           "Проверьте путь: " & filePath, vbExclamation, "Ошибка"
End Sub

Private Sub UserForm_Initialize()
    ' === РАЗМЕР ФОРМЫ (увеличиваем для новой кнопки) ===
    Me.Width = 600
    Me.Height = 680  ' Увеличено для новой кнопки
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
        .text = "Пропишанс - пар03 - 720 (720)" & vbNewLine & _
"Пропишанс - пар02 - 690 (720)" & vbNewLine & _
"Пропишанс - пар03 - 400 (720)" & vbNewLine & _
"Лерашанс - пар01 - 720 (720)" & vbNewLine & _
"Лерашанс - пар02 - 110 (720)" & vbNewLine & _
"Пропишанс Универсал - пар13 - 720 (720)" & vbNewLine & _
"Пропишанс Универсал - пар13 - 720 (720)" & vbNewLine & _
"Пропишанс Универсал - пар13 - 720 (720)" & vbNewLine & _
"Пропишанс Универсал - пар27 - 720 (720)" & vbNewLine & _
"Пропишанс Универсал - пар27 - 720 (720)" & vbNewLine & _
"Пропишанс Универсал - пар27 - 720 (720)"
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
        .Width = 100
        .Height = 30
    End With
    
    With CommandButton7  ' Помощь
        .Caption = "? Помощь"
        .Top = 420
        .Left = 120
        .Width = 100
        .Height = 30
    End With
    
    ' === НОВАЯ КНОПКА ПЕЧАТИ ===
    With CommandButton10  ' ПЕЧАТЬ
        .Caption = "??? ПЕЧАТЬ"
        .Top = 420
        .Left = 230
        .Width = 100
        .Height = 30
        .Font.Bold = True
        .BackColor = RGB(100, 149, 237)  ' Синий цвет
        .ForeColor = RGB(255, 255, 255)
        .Enabled = False  ' Включается после размещения
    End With
    
    With CommandButton9  ' СОХРАНИТЬ
        .Caption = "?? СОХРАНИТЬ"
        .Top = 420
        .Left = 340
        .Width = 100
        .Height = 30
        .Font.Bold = True
        .BackColor = RGB(255, 165, 0)  ' Оранжевый цвет
        .ForeColor = RGB(255, 255, 255)
        .Enabled = False  ' Включается после размещения
    End With
    
    With CommandButton8  ' Закрыть
        .Caption = "? Закрыть"
        .Top = 420
        .Left = 450
        .Width = 100
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
        CommandButton9.Enabled = True   ' Включаем кнопку СОХРАНИТЬ
        CommandButton10.Enabled = True  ' Включаем кнопку ПЕЧАТЬ
        
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
    CommandButton9.Enabled = False   ' Выключаем кнопку СОХРАНИТЬ
    CommandButton10.Enabled = False  ' Выключаем кнопку ПЕЧАТЬ
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
    help = help & "• Помечает новые ячейки желтым цветом" & vbNewLine & vbNewLine
    help = help & "??? ПЕЧАТЬ:" & vbNewLine
    help = help & "• Кнопка 'ПЕЧАТЬ' создает красивый табличный отчет" & vbNewLine
    help = help & "• Автоматическая сортировка по наименованию" & vbNewLine
    help = help & "• Наглядное отображение со стрелками" & vbNewLine & vbNewLine
    help = help & "?? СОХРАНИТЬ:" & vbNewLine
    help = help & "• Сохраняет отчет в текстовый файл" & vbNewLine
    help = help & "• Автоматически открывает в Блокноте"
    
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

' ===== ФУНКЦИИ ДЛЯ КРАСИВОЙ ПЕЧАТИ =====

' ===== ФУНКЦИЯ КРАСИВОЙ ПЕЧАТИ =====
Private Sub PrintBeautifulReport(reportText As String)
    On Error GoTo ErrorHandler
    
    ' Парсим данные из отчета
    Dim placements As Collection
    Set placements = ParseReportData(reportText)
    
    If placements.count = 0 Then
        MsgBox "Нет данных для печати!", vbInformation
        Exit Sub
    End If
    
    ' Создаем новый лист для печати
    Dim printSheet As Worksheet
    Set printSheet = CreatePrintSheet()
    
    If printSheet Is Nothing Then
        Exit Sub
    End If
    
    ' Заполняем данные
    Call FillPrintSheet(printSheet, placements)
    
    ' Форматируем таблицу
    Call FormatPrintSheet(printSheet, placements.count)
    
    ' Показываем предварительный просмотр и печать
    printSheet.Select
    
    MsgBox "?? Отчет готов к печати!" & vbNewLine & vbNewLine & _
           "• Данные отсортированы по наименованию" & vbNewLine & _
           "• Красивое табличное оформление" & vbNewLine & _
           "• Стрелки для наглядности" & vbNewLine & _
           "• Всего позиций: " & placements.count & vbNewLine & vbNewLine & _
           "Нажмите OK для печати...", vbInformation, "Готов к печати"
    
    ' Вызываем диалог печати
    Application.Dialogs(xlDialogPrint).Show
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка создания отчета для печати:" & vbNewLine & Err.description, vbCritical
End Sub

' ===== ПАРСИНГ ДАННЫХ ИЗ ОТЧЕТА =====
Private Function ParseReportData(reportText As String) As Collection
    Set ParseReportData = New Collection
    
    On Error GoTo ErrorHandler
    
    ' Разбиваем отчет на строки
    Dim lines() As String
    lines = Split(reportText, vbNewLine)
    
    ' Ищем секцию с размещениями
    Dim foundPlacements As Boolean
    foundPlacements = False
    
    Dim i As Integer
    For i = 0 To UBound(lines)
        Dim line As String
        line = Trim(lines(i))
        
        ' Ищем начало секции размещений
        If InStr(line, "РАЗМЕЩЕНИЯ ЯЧЕЙКА-ЗА-ЯЧЕЙКОЙ:") > 0 Then
            foundPlacements = True
            GoTo NextLine
        End If
        
        ' Если нашли размещения и строка содержит данные
        If foundPlacements And line <> "" And InStr(line, "---------") = 0 Then
            ' Проверяем, не закончились ли размещения
            If InStr(line, "ОШИБКИ:") > 0 Or InStr(line, "ИСПРАВЛЕНИЯ:") > 0 Then
                Exit For
            End If
            
            ' Парсим строку размещения
            Dim placement(3) As String  ' 0=ProductName, 1=Batch, 2=Quantity, 3=Location
            If ParsePlacementLine(line, placement) Then
                ParseReportData.Add placement
            End If
        End If
        
NextLine:
    Next i
    
    ' Сортируем по наименованию
    If ParseReportData.count > 1 Then
        Set ParseReportData = SortPlacementsByName(ParseReportData)
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "Ошибка парсинга данных: " & Err.description, vbCritical
    Set ParseReportData = New Collection
End Function

' ===== ПАРСИНГ ОДНОЙ СТРОКИ РАЗМЕЩЕНИЯ =====
Private Function ParsePlacementLine(line As String, ByRef placement() As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Пример строки: "Ангар 6-28-ПРЗ-1 - Галошанс (пар08) - 110 "
    
    ' Убираем "Ангар " в начале
    If InStr(line, "Ангар ") = 1 Then
        line = Mid(line, 7) ' Убираем "Ангар "
    End If
    
    ' Разбиваем по " - "
    Dim parts() As String
    parts = Split(line, " - ")
    
    If UBound(parts) < 2 Then
        ParsePlacementLine = False
        Exit Function
    End If
    
    ' Получаем место размещения
    placement(3) = Trim(parts(0))  ' Location
    
    ' Парсим продукт и партию: "Галошанс (пар08)"
    Dim productPart As String
    productPart = Trim(parts(1))
    
    ' Ищем партию в скобках
    Dim openBracket As Integer, closeBracket As Integer
    openBracket = InStrRev(productPart, "(")
    closeBracket = InStrRev(productPart, ")")
    
    If openBracket > 0 And closeBracket > openBracket Then
        placement(0) = Trim(Left(productPart, openBracket - 1))  ' ProductName
        placement(1) = Trim(Mid(productPart, openBracket + 1, closeBracket - openBracket - 1))  ' Batch
    Else
        placement(0) = productPart  ' ProductName
        placement(1) = ""  ' Batch
    End If
    
    ' Получаем количество
    placement(2) = Trim(parts(2))  ' Quantity
    
    ParsePlacementLine = True
    Exit Function
    
ErrorHandler:
    ParsePlacementLine = False
End Function

' ===== СОРТИРОВКА ПО НАИМЕНОВАНИЮ =====
Private Function SortPlacementsByName(unsortedCollection As Collection) As Collection
    On Error GoTo ErrorHandler
    
    ' Создаем массив для сортировки
    Dim tempArray() As Variant
    ReDim tempArray(1 To unsortedCollection.count)
    
    ' Копируем в массив
    Dim i As Integer
    For i = 1 To unsortedCollection.count
        tempArray(i) = unsortedCollection(i)
    Next i
    
    ' ===== ТРЕХУРОВНЕВАЯ СОРТИРОВКА =====
    Dim j As Integer
    Dim temp As Variant
    For i = 1 To UBound(tempArray) - 1
        For j = i + 1 To UBound(tempArray)
            Dim item1() As String, item2() As String
            item1 = tempArray(i)
            item2 = tempArray(j)
            
            ' Извлекаем номера ангаров из Location
            Dim warehouse1 As Integer, warehouse2 As Integer
            warehouse1 = GetWarehouseFromLocation(item1(3))
            warehouse2 = GetWarehouseFromLocation(item2(3))
            
            ' Извлекаем количества
            Dim qty1 As Double, qty2 As Double
            qty1 = Val(item1(2))
            qty2 = Val(item2(2))
            
            Dim needSwap As Boolean
            needSwap = False
            
            ' УРОВЕНЬ 1: Сортировка по ангарам
            If warehouse1 > warehouse2 Then
                needSwap = True
            ElseIf warehouse1 = warehouse2 Then
                ' УРОВЕНЬ 2: Если ангары одинаковые - сортируем по названию товара
                If UCase(item1(0)) > UCase(item2(0)) Then
                    needSwap = True
                ElseIf UCase(item1(0)) = UCase(item2(0)) Then
                    ' УРОВЕНЬ 3: Если товары одинаковые - сортируем по количеству (УБЫВАНИЕ)
                    If qty1 < qty2 Then  ' Меньшее количество идет ПОСЛЕ большего
                        needSwap = True
                    End If
                End If
            End If
            
            ' Меняем местами если нужно
            If needSwap Then
                temp = tempArray(i)
                tempArray(i) = tempArray(j)
                tempArray(j) = temp
            End If
        Next j
    Next i
    
    ' Создаем новую коллекцию
    Set SortPlacementsByName = New Collection
    For i = 1 To UBound(tempArray)
        SortPlacementsByName.Add tempArray(i)
    Next i
    
    Exit Function
    
ErrorHandler:
    Set SortPlacementsByName = unsortedCollection
End Function
Private Function GetWarehouseFromLocation(location As String) As Integer
    On Error Resume Next
    
    ' Из строки "6-38-М-1" извлекаем "6"
    Dim parts() As String
    parts = Split(location, "-")
    
    If UBound(parts) >= 0 Then
        GetWarehouseFromLocation = CInt(parts(0))
    Else
        GetWarehouseFromLocation = 999  ' Если ошибка - помещаем в конец
    End If
    
    On Error GoTo 0
End Function

' ===== СОЗДАНИЕ ЛИСТА ДЛЯ ПЕЧАТИ =====
Private Function CreatePrintSheet() As Worksheet
    On Error GoTo ErrorHandler
    
    ' Удаляем старый лист если есть
    Dim sheetName As String
    sheetName = "Отчет_Печать"
    
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name = sheetName Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
            Exit For
        End If
    Next ws
    
    ' Создаем новый лист
    Set CreatePrintSheet = ActiveWorkbook.Worksheets.Add
    CreatePrintSheet.Name = sheetName
    
    ' Настройки печати
    With CreatePrintSheet.PageSetup
        .Orientation = xlPortrait ' Книжная ориентация
        .PaperSize = xlPaperA4
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .CenterHorizontally = True
        .CenterVertically = True
    End With
    
    Exit Function
    
ErrorHandler:
    MsgBox "Ошибка создания листа: " & Err.description, vbCritical
    Set CreatePrintSheet = Nothing
End Function

' ===== ЗАПОЛНЕНИЕ ЛИСТА ДАННЫМИ =====
Private Sub FillPrintSheet(printSheet As Worksheet, placements As Collection)
    On Error GoTo ErrorHandler
    
    With printSheet
        ' Очищаем лист
        .Cells.Clear
        
        ' === ЗАГОЛОВОК ОТЧЕТА ===
        .Cells(1, 1).value = "ОТЧЕТ О РАЗМЕЩЕНИИ ТОВАРОВ"
        .Cells(2, 1).value = "Дата: " & Format(Now, "dd.mm.yyyy hh:mm:ss")
        .Cells(3, 1).value = "Активные ангары: " & GetActiveWarehousesList()
        
        ' === ЗАГОЛОВКИ ТАБЛИЦЫ ===
        Dim headerRow As Long
        headerRow = 6
        
        .Cells(headerRow, 1).value = "Наименований"
        .Cells(headerRow, 2).value = "Партия"
        .Cells(headerRow, 3).value = "Кол-во"
        .Cells(headerRow, 4).value = "Куда"
        .Cells(headerRow, 5).value = "Ангар"
        
        ' === ДАННЫЕ ===
        Dim dataRow As Long
        dataRow = headerRow + 1
        
        Dim i As Integer
        For i = 1 To placements.count
            Dim placement() As String
            placement = placements(i)
            
            .Cells(dataRow, 1).value = placement(0)  ' ProductName
            .Cells(dataRow, 2).value = "(" & placement(1) & ")"  ' Batch
            .Cells(dataRow, 3).value = placement(2)  ' Quantity
            .Cells(dataRow, 4).value = ">"  ' Стрелка
            .Cells(dataRow, 5).value = placement(3)  ' Location
            
            dataRow = dataRow + 1
        Next i
        
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка заполнения данных: " & Err.description, vbCritical
End Sub

' ===== ФОРМАТИРОВАНИЕ ТАБЛИЦЫ =====
Private Sub FormatPrintSheet(printSheet As Worksheet, dataCount As Long)
    On Error GoTo ErrorHandler
    
    With printSheet
        ' === ЗАГОЛОВОК ОТЧЕТА ===
        With .Range("A1:E1")
            .Merge
            .Font.Size = 16
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(0, 100, 200)
            .Font.Color = RGB(255, 255, 255)
            .RowHeight = 30
        End With
        
        With .Range("A2:E2")
            .Merge
            .Font.Size = 12
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(220, 230, 241)
            .RowHeight = 20
        End With
        
        With .Range("A3:E3")
            .Merge
            .Font.Size = 12
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(220, 230, 241)
            .RowHeight = 20
        End With
        
        ' === ЗАГОЛОВКИ ТАБЛИЦЫ ===
        Dim headerRange As String
        headerRange = "A6:E6"
        
        With .Range(headerRange)
            .Font.Bold = True
            .Font.Size = 12
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = RGB(79, 129, 189)
            .Font.Color = RGB(255, 255, 255)
            .RowHeight = 30
        End With
        
        ' === ДАННЫЕ ТАБЛИЦЫ ===
        If dataCount > 0 Then
            Dim dataRange As String
            dataRange = "A7:E" & (6 + dataCount)
            
            With .Range(dataRange)
                .Font.Size = 11
                .VerticalAlignment = xlCenter
                .RowHeight = 25
            End With
            
            ' Выравнивание колонок
            .columns("A").HorizontalAlignment = xlLeft     ' Наименования - влево
            .columns("B").HorizontalAlignment = xlCenter   ' Партия - по центру
            .columns("C").HorizontalAlignment = xlCenter   ' Количество - по центру
            .columns("D").HorizontalAlignment = xlCenter   ' Стрелка - по центру
            .columns("E").HorizontalAlignment = xlCenter   ' Ангар - по центру
            
            ' Цвет стрелок
            .columns("D").Font.Color = RGB(255, 0, 0)
            .columns("D").Font.Size = 14
            .columns("D").Font.Bold = True
            
            ' Чередующиеся цвета строк
            Dim row As Long
            For row = 7 To 6 + dataCount
                If row Mod 2 = 0 Then
                    .Range("A" & row & ":E" & row).Interior.Color = RGB(242, 242, 242)
                End If
            Next row
        End If
        
        ' === ГРАНИЦЫ ТАБЛИЦЫ ===
        Dim tableRange As String
        tableRange = "A6:E" & (6 + dataCount)
        
        With .Range(tableRange).Borders
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(0, 0, 0)
        End With
        
        ' === ШИРИНА КОЛОНОК ===
        .columns("A").ColumnWidth = 25  ' Наименования
        .columns("B").ColumnWidth = 12  ' Партия
        .columns("C").ColumnWidth = 10  ' Количество
        .columns("D").ColumnWidth = 8   ' Стрелка
        .columns("E").ColumnWidth = 20  ' Ангар
        
        ' === АВТОПОДГОНКА ВЫСОТЫ ===
        .rows("1:3").AutoFit
        
        ' === ОБЛАСТЬ ПЕЧАТИ ===
        .PageSetup.PrintArea = "A1:E" & (6 + dataCount + 2)
        
        ' === ЗАГОЛОВОК И ПОДВАЛ ===
        .PageSetup.CenterHeader = "&B&14Отчет о размещении товаров"
        .PageSetup.RightFooter = "&D &T - Стр. &P из &N"
        
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка форматирования: " & Err.description, vbCritical
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

