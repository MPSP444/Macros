Attribute VB_Name = "ModuleProductConfig"

' ===== ModuleProductConfig - НАСТРОЙКА ГРУПП ТОВАРОВ =====
Option Explicit

' ===== БЫСТРАЯ НАСТРОЙКА ВСЕХ ГРУПП =====
Public Sub ConfigureAllGroups()
    ' Инициализация
    Call ModuleProductInfo.InitializeProductDatabase
    
    MsgBox "НАСТРОЙКА ГРУПП ПРЕПАРАТОВ" & vbNewLine & vbNewLine & _
           "Будут настроены группы:" & vbNewLine & _
           "• 1000 л - Виташанс, Фумишанс" & vbNewLine & _
           "• 960 л - Глифошанс, Дикошанс" & vbNewLine & _
           "• 600 л - Гоплит, Фасшанс и др." & vbNewLine & _
           "• 540 л - Антимышин, Меташанс" & vbNewLine & _
           "• 288 кг - Эллада, Эльшанс и др." & vbNewLine & _
           "• 240 кг - Шантус, Шанстар и др." & vbNewLine & _
           "• 720 - ВСЕ ОСТАЛЬНЫЕ (Лерашанс и т.д.)", _
           vbInformation, "Информация"
    
    ' Добавляем препараты в группы
    Call SetupGroup1000
    Call SetupGroup960
    Call SetupGroup600
    Call SetupGroup540
    Call SetupGroup288
    Call SetupGroup240
    
    MsgBox "? Все группы препаратов настроены!" & vbNewLine & _
           "Всего: 28 препаратов в группах" & vbNewLine & _
           "Остальные будут использовать 720 (стандарт)", vbInformation, "Готово"
End Sub

' ===== ГРУППА 1000 литров =====
Private Sub SetupGroup1000()
    Call ModuleProductInfo.AddProductToGroup("ВИТАШАНС", 1000)
    Call ModuleProductInfo.AddProductToGroup("ФУМИШАНС", 1000)
End Sub

' ===== ГРУППА 960 литров =====
Private Sub SetupGroup960()
    Call ModuleProductInfo.AddProductToGroup("ГЛИФОШАНС СУПЕР", 960)
    Call ModuleProductInfo.AddProductToGroup("ДИКОШАНС", 960)
End Sub

' ===== ГРУППА 600 литров =====
Private Sub SetupGroup600()
    Call ModuleProductInfo.AddProductToGroup("ГОПЛИТ", 600)
    Call ModuleProductInfo.AddProductToGroup("ФАСШАНС", 600)
    Call ModuleProductInfo.AddProductToGroup("КАЛИНА", 600)
    Call ModuleProductInfo.AddProductToGroup("КАРАТОШАНС", 600)
    Call ModuleProductInfo.AddProductToGroup("КРЕПОШАНС", 600)
    Call ModuleProductInfo.AddProductToGroup("СЕКТОР", 600)
    Call ModuleProductInfo.AddProductToGroup("ШАНС-90", 600)
End Sub

' ===== ГРУППА 540 литров =====
Private Sub SetupGroup540()
    Call ModuleProductInfo.AddProductToGroup("АНТИМЫШИН", 540)
    Call ModuleProductInfo.AddProductToGroup("МЕТАШАНС", 540)
End Sub

' ===== ГРУППА 288 кг =====
Private Sub SetupGroup288()
    Call ModuleProductInfo.AddProductToGroup("ЭЛЛАДА", 288)
    Call ModuleProductInfo.AddProductToGroup("ЭЛЬШАНС", 288)
    Call ModuleProductInfo.AddProductToGroup("ДЕЛАТОН", 288)
    Call ModuleProductInfo.AddProductToGroup("ПЕНТАГОН", 288)
    Call ModuleProductInfo.AddProductToGroup("ПОЛИДОК", 288)
End Sub

' ===== ГРУППА 240 кг =====
Private Sub SetupGroup240()
    Call ModuleProductInfo.AddProductToGroup("ШАНТУС", 240)
    Call ModuleProductInfo.AddProductToGroup("ШАНСПРОФИ", 240)
    Call ModuleProductInfo.AddProductToGroup("ШАНСТАР", 240)
    Call ModuleProductInfo.AddProductToGroup("ШАНСТАР ПЛЮС", 240)
    Call ModuleProductInfo.AddProductToGroup("ШАНСТИ", 240)
    Call ModuleProductInfo.AddProductToGroup("ЗНАТОК", 240)
    Call ModuleProductInfo.AddProductToGroup("ХОРИСТ", 240)
    Call ModuleProductInfo.AddProductToGroup("ШАНСИЛИН", 240)
End Sub

' ===== ДОБАВИТЬ ОДИН ТОВАР =====
Public Sub AddSingleProduct()
    Dim ProductName As String
    ProductName = InputBox("Введите название препарата:", "Добавление препарата")
    
    If ProductName = "" Then Exit Sub
    
    Dim capacity As String
    capacity = InputBox("Введите вместимость:" & vbNewLine & _
                       "1000, 960, 600, 540, 288, 240" & vbNewLine & _
                       "(или оставьте пустым для 720)", _
                       "Вместимость", "720")
    
    If capacity = "" Or capacity = "720" Then
        MsgBox "Препарат '" & ProductName & "' будет использовать стандартную вместимость 720", vbInformation
    Else
        Call ModuleProductInfo.AddProductToGroup(ProductName, CDbl(capacity))
        MsgBox "? Препарат '" & ProductName & "' добавлен в группу " & capacity, vbInformation
    End If
End Sub

' ===== ДОБАВИТЬ НЕСКОЛЬКО ТОВАРОВ В ГРУППУ =====
Public Sub AddMultipleProducts()
    Dim capacity As String
    capacity = InputBox("Выберите группу по вместимости:" & vbNewLine & _
                       "1000 - литры (большие)" & vbNewLine & _
                       "960 - литры (Глифошанс)" & vbNewLine & _
                       "600 - литры (средние)" & vbNewLine & _
                       "540 - литры (специальные)" & vbNewLine & _
                       "288 - кг (порошки)" & vbNewLine & _
                       "240 - кг (мелкие)", _
                       "Выбор группы", "600")
    
    If capacity = "" Then Exit Sub
    
    Dim products As String
    products = InputBox("Введите препараты через запятую:" & vbNewLine & _
                       "Пример: Препарат1, Препарат2, Препарат3", _
                       "Добавление препаратов в группу " & capacity)
    
    If products = "" Then Exit Sub
    
    ' Разбиваем и добавляем
    Dim productArray() As String
    productArray = Split(products, ",")
    
    Dim i As Integer
    Dim count As Integer
    count = 0
    
    For i = 0 To UBound(productArray)
        Dim ProductName As String
        ProductName = Trim(productArray(i))
        
        If ProductName <> "" Then
            Call ModuleProductInfo.AddProductToGroup(ProductName, CDbl(capacity))
            count = count + 1
        End If
    Next i
    
    MsgBox "? Добавлено препаратов: " & count & " в группу " & capacity, vbInformation
End Sub

' ===== ПОКАЗАТЬ ТЕКУЩИЕ ГРУППЫ =====
Public Sub ShowCurrentGroups()
    Call ModuleProductInfo.InitializeProductDatabase
    
    Dim result As String
    result = ModuleProductInfo.ShowAllGroups()
    
    ' Показываем в MsgBox если текст не очень большой
    If Len(result) < 1000 Then
        MsgBox result, vbInformation, "Группы товаров"
    Else
        ' Иначе выводим в Immediate Window
        Debug.Print result
        MsgBox "Список групп выведен в Immediate Window (Ctrl+G)", vbInformation
    End If
End Sub

' ===== ПРОВЕРИТЬ ТОВАР =====
Public Sub CheckProduct()
    Dim ProductName As String
    ProductName = InputBox("Введите название препарата для проверки:", "Проверка препарата")
    
    If ProductName = "" Then Exit Sub
    
    Dim capacity As Double
    capacity = ModuleProductInfo.GetCellCapacity(ProductName)
    
    Dim message As String
    message = "Препарат: " & ProductName & vbNewLine & _
              "Вместимость палеты: " & capacity
    
    If capacity >= 540 Then
        message = message & " литров"
    ElseIf capacity <= 288 Then
        message = message & " кг"
    End If
    
    message = message & vbNewLine & "Максимум на букву: 1500" & vbNewLine & vbNewLine
    
    If capacity = 720 Then
        message = message & "?? Используется СТАНДАРТНАЯ вместимость"
    Else
        message = message & "?? Препарат в группе " & capacity
    End If
    
    MsgBox message, vbInformation, "Информация о препарате"
End Sub

' ===== ЭКСПОРТ В EXCEL =====
Public Sub ExportGroups()
    Call ModuleProductInfo.ExportGroupsToExcel
End Sub

' ===== ПРИМЕР НАСТРОЙКИ ДЛЯ ВАШИХ ТОВАРОВ =====
Public Sub SetupYourProducts()
    ' Здесь вы можете добавить ВАШИ препараты
    ' Просто раскомментируйте и измените названия
    
    ' Группа 1000 литров
    ' Call ModuleProductInfo.AddProductToGroup("ВАШ_ПРЕПАРАТ_1", 1000)
    
    ' Группа 960 литров
    ' Call ModuleProductInfo.AddProductToGroup("ВАШ_ПРЕПАРАТ_2", 960)
    
    ' Группа 600 литров
    ' Call ModuleProductInfo.AddProductToGroup("ВАШ_ПРЕПАРАТ_3", 600)
    
    ' Группа 540 литров
    ' Call ModuleProductInfo.AddProductToGroup("ВАШ_ПРЕПАРАТ_4", 540)
    
    ' Группа 288 кг
    ' Call ModuleProductInfo.AddProductToGroup("ВАШ_ПРЕПАРАТ_5", 288)
    
    ' Группа 240 кг
    ' Call ModuleProductInfo.AddProductToGroup("ВАШ_ПРЕПАРАТ_6", 240)
    
    ' Все остальные автоматически будут 720!
    
    MsgBox "Добавьте ваши препараты в эту процедуру!", vbInformation
End Sub

' ===== ПОКАЗАТЬ СВОДНУЮ ТАБЛИЦУ =====
Public Sub ShowSummaryTable()
    Dim summary As String
    summary = "?? СВОДНАЯ ТАБЛИЦА ПРЕПАРАТОВ" & vbNewLine
    summary = summary & String(40, "=") & vbNewLine & vbNewLine
    
    summary = summary & "ГРУППА 1000 л: Виташанс, Фумишанс" & vbNewLine
    summary = summary & "ГРУППА 960 л:  Глифошанс Супер, Дикошанс" & vbNewLine
    summary = summary & "ГРУППА 600 л:  Гоплит, Фасшанс, Калина," & vbNewLine
    summary = summary & "               Каратошанс, Крепошанс," & vbNewLine
    summary = summary & "               Сектор, Шанс-90" & vbNewLine
    summary = summary & "ГРУППА 540 л:  Антимышин, Меташанс" & vbNewLine
    summary = summary & "ГРУППА 288 кг: Эллада, Эльшанс, Делатон," & vbNewLine
    summary = summary & "               Пентагон, Полидок" & vbNewLine
    summary = summary & "ГРУППА 240 кг: Шантус, Шанспрофи, Шанстар," & vbNewLine
    summary = summary & "               Шанстар Плюс, Шансти," & vbNewLine
    summary = summary & "               Знаток, Хорист, Шансилин" & vbNewLine & vbNewLine
    
    summary = summary & "СТАНДАРТ 720:  Лерашанс, Керемет и все остальные"
    
    MsgBox summary, vbInformation, "Сводная таблица"
End Sub
