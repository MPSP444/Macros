Attribute VB_Name = "ModuleLauncher"


' ===== ModuleLauncher - Простой модуль запуска системы массового ввода =====
Option Explicit

' ===== ОСНОВНЫЕ ФУНКЦИИ ЗАПУСКА =====

Public Sub ЗапуститьСистему()
    ' Главная функция запуска системы (русское название для удобства)
    Call StartSystem
End Sub

Public Sub StartMassInputSystem()
    On Error GoTo ErrorHandler
    
    ' Инициализируем систему логирования
    Call ModuleLogger.InitializeLogger(True, ModuleLogger.LOG_LEVEL_INFO)
    
    ' Логируем информацию о системе
    Call ModuleLogger.LogSystemInfo
    
    ' Проверяем готовность системы
    If Not CheckSystemReadiness() Then
        Exit Sub
    End If
    
    ' ИСПРАВЛЕНО: Показываем форму вместо InputBox
    ModuleLogger.LogMessage "Запуск формы массового ввода товаров"
    MassProductInputForm.Show
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка запуска системы: " & Err.description, vbCritical, "Ошибка системы"
    ModuleLogger.LogError "Критическая ошибка запуска: " & Err.description
End Sub

Public Sub ОчиститьВсе()
    ' Функция быстрой очистки всех ангаров (русское название)
    Call ClearAllData
End Sub

Public Sub ClearAllData()
    ' Функция быстрой очистки всех данных
    On Error GoTo ErrorHandler
    
    ' Проверяем готовность
    If Not CheckSystemReady() Then
        Exit Sub
    End If
    
    ' Показываем предупреждение
    Dim confirmResult As VbMsgBoxResult
    confirmResult = ShowClearWarning()
    
    If confirmResult = vbYes Then
        ' Инициализируем логирование
        Call ModuleLogger.InitializeLogger(True, ModuleLogger.LOG_LEVEL_INFO)
        
        ' Запускаем очистку
        Call MainLauncher.QuickClearAllData
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка очистки данных: " & Err.description, vbCritical, "Ошибка очистки"
End Sub
Public Sub ShowSmartForm()  ' Изменено имя процедуры
    On Error GoTo ErrorHandler
    
    ' Инициализируем систему логирования
    Call ModuleLogger.InitializeLogger(True, ModuleLogger.LOG_LEVEL_INFO)
    
    ' Логируем информацию о системе
    Call ModuleLogger.LogSystemInfo
    
    ' Показываем форму
    ModuleLogger.LogMessage "Запуск формы массового ввода товаров"
    SimpleSmartForm.Show  ' Теперь это однозначно форма
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка запуска системы: " & Err.description, vbCritical, "Ошибка системы"
    ModuleLogger.LogError "Критическая ошибка запуска: " & Err.description
End Sub
Public Sub ПоказатьСправку()
    ' Показать справку по системе (русское название)
    Call ShowHelp
End Sub

Public Sub ShowHelp()
    ' Показать подробную справку по использованию системы
    Call ShowSystemManual
End Sub

Public Sub ТестСистемы()
    ' Быстрый тест системы (русское название)
    Call TestSystem
End Sub

Public Sub TestSystem()
    ' Функция тестирования системы
    On Error GoTo ErrorHandler
    
    ' Инициализируем логирование для теста
    Call ModuleLogger.InitializeLogger(True, ModuleLogger.LOG_LEVEL_DEBUG)
    
    ' Запускаем тест
    Call MainLauncher.TestSystem
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка тестирования: " & Err.description, vbCritical, "Ошибка теста"
End Sub

' ===== ФУНКЦИИ ПРОВЕРКИ СИСТЕМЫ =====

Private Function CheckSystemReady() As Boolean
    ' Проверяет готовность системы к работе
    CheckSystemReady = False
    
    ' Проверяем активную книгу
    If ActiveWorkbook Is Nothing Then
        MsgBox "? ОШИБКА: Не найдена активная книга Excel!" & vbNewLine & vbNewLine & _
               "Пожалуйста:" & vbNewLine & _
               "1. Откройте книгу с ангарами" & vbNewLine & _
               "2. Убедитесь, что есть листы 'Ангар 5', 'Ангар 6' и т.д." & vbNewLine & _
               "3. Попробуйте снова", _
               vbExclamation, "Нет активной книги"
        Exit Function
    End If
    
    ' Проверяем наличие основных листов
    Dim missingSheets As String
    missingSheets = CheckRequiredSheets()
    
    If missingSheets <> "" Then
        Dim continueWork As VbMsgBoxResult
        continueWork = MsgBox("?? ПРЕДУПРЕЖДЕНИЕ: Не найдены следующие листы:" & vbNewLine & _
                              missingSheets & vbNewLine & vbNewLine & _
                              "Система может работать некорректно." & vbNewLine & _
                              "Продолжить работу?", _
                              vbYesNo + vbQuestion, "Отсутствуют листы")
        
        If continueWork = vbNo Then
            Exit Function
        End If
    End If
    
    CheckSystemReady = True
End Function

Private Function CheckRequiredSheets() As String
    ' Проверяет наличие необходимых листов ангаров
    Dim missingSheets As String
    missingSheets = ""
    
    Dim warehouseNumbers As Variant
    warehouseNumbers = Array(5, 6, 7, 8, 9, 10, 11, 12)
    
    Dim i As Integer
    For i = 0 To UBound(warehouseNumbers)
        Dim sheetName As String
        sheetName = "Ангар " & warehouseNumbers(i)
        
        On Error Resume Next
        Dim ws As Worksheet
        Set ws = ActiveWorkbook.Worksheets(sheetName)
        
        If Err.Number <> 0 Then
            If missingSheets <> "" Then missingSheets = missingSheets & ", "
            missingSheets = missingSheets & sheetName
        End If
        On Error GoTo 0
    Next i
    
    CheckRequiredSheets = missingSheets
End Function

' ===== ФУНКЦИИ ПОЛЬЗОВАТЕЛЬСКОГО ИНТЕРФЕЙСА =====

Private Sub ShowWelcomeMessage()
    ' Показывает приветственное сообщение
    Dim welcomeMsg As String
    welcomeMsg = "?? СИСТЕМА МАССОВОГО ВВОДА ТОВАРОВ НА СКЛАД" & vbNewLine & vbNewLine & _
                 "?? Возможности системы:" & vbNewLine & _
                 "• Массовый ввод товаров в ангары" & vbNewLine & _
                 "• Автоматическое создание заголовков" & vbNewLine & _
                 "• Правильная работа с секциями" & vbNewLine & _
                 "• Обнаружение конфликтов товаров" & vbNewLine & _
                 "• Суммирование дублирующихся позиций" & vbNewLine & _
                 "• Подробное логирование операций" & vbNewLine & vbNewLine & _
                 "? Система готова к запуску!"
    
    MsgBox welcomeMsg, vbInformation, "Добро пожаловать!"
End Sub

Private Function ShowClearWarning() As VbMsgBoxResult
    ' Показывает предупреждение перед очисткой
    Dim warningMsg As String
    warningMsg = "?? ВНИМАНИЕ! ПОЛНАЯ ОЧИСТКА ВСЕХ АНГАРОВ!" & vbNewLine & vbNewLine & _
                 "? БУДЕТ УДАЛЕНО:" & vbNewLine & _
                 "• Все заголовки товаров" & vbNewLine & _
                 "• Все количества товаров" & vbNewLine & _
                 "• Все номера партий" & vbNewLine & _
                 "• Все результаты подсчетов" & vbNewLine & vbNewLine & _
                 "? ОСТАНЕТСЯ:" & vbNewLine & _
                 "• Структура таблиц" & vbNewLine & _
                 "• Названия стеллажей (А, Б, В...)" & vbNewLine & _
                 "• Границы и форматирование" & vbNewLine & vbNewLine & _
                 "?? Вы ДЕЙСТВИТЕЛЬНО хотите очистить ВСЕ данные?"
    
    ShowClearWarning = MsgBox(warningMsg, vbYesNo + vbCritical, "Подтверждение очистки")
End Function

Private Sub ShowSystemManual()
    ' Показывает подробное руководство по системе
    Dim manual As String
    manual = "?? РУКОВОДСТВО ПО СИСТЕМЕ МАССОВОГО ВВОДА" & vbNewLine & vbNewLine & _
             "?? ФОРМАТ ВВОДА ДАННЫХ:" & vbNewLine & _
             "Ангар-Ряд-Стеллаж-Уровень - Товар - Партия - Количество" & vbNewLine & vbNewLine & _
             "?? ПРИМЕРЫ:" & vbNewLine & _
             "6-5-Е-2 - Лерашанс - пар13 - 720" & vbNewLine & _
             "7-1-А-1 - Босфор - пар03 - 480" & vbNewLine & _
             "8-3-П-2 - Дикошанс - пар25 - 360" & vbNewLine & vbNewLine & _
             "?? ПОДДЕРЖИВАЕМЫЕ АНГАРЫ: 5, 6, 7, 8, 9, 10, 11, 12" & vbNewLine & _
             "?? УРОВНИ СТЕЛЛАЖЕЙ: 1, 2, 3" & vbNewLine & _
             "?? СТЕЛЛАЖИ: А, Б, В, Г... (зависит от ангара)" & vbNewLine & vbNewLine & _
             "? БЫСТРЫЕ КОМАНДЫ:" & vbNewLine & _
             "• ЗапуститьСистему() - запуск основной формы" & vbNewLine & _
             "• ОчиститьВсе() - очистка всех данных" & vbNewLine & _
             "• ТестСистемы() - проверка работоспособности" & vbNewLine & _
             "• ПоказатьСправку() - эта справка" & vbNewLine & vbNewLine & _
             "?? Для запуска нажмите Alt+F8 и выберите нужную команду"
    
    MsgBox manual, vbInformation, "Руководство пользователя"
End Sub

' ===== ДОПОЛНИТЕЛЬНЫЕ УТИЛИТЫ =====

Public Sub ПоказатьСтатистику()
    ' Показать статистику системы (русское название)
    Call ShowSystemStatistics
End Sub

Public Sub ShowSystemStatistics()
    ' Показывает статистику системы
    On Error GoTo ErrorHandler
    
    Dim stats As String
    stats = "?? СТАТИСТИКА СИСТЕМЫ" & vbNewLine & vbNewLine
    
    ' Информация о книге
    If ActiveWorkbook Is Nothing Then
        stats = stats & "? Активная книга: НЕ НАЙДЕНА" & vbNewLine
    Else
        stats = stats & "?? Активная книга: " & ActiveWorkbook.Name & vbNewLine
        stats = stats & "?? Путь: " & ActiveWorkbook.Path & vbNewLine
    End If
    
    ' Подсчет доступных ангаров
    Dim availableCount As Integer
    Dim totalCount As Integer
    availableCount = 0
    totalCount = 8
    
    Dim warehouseList As String
    warehouseList = ""
    
    Dim i As Integer
    For i = 5 To 12
        On Error Resume Next
        Dim ws As Worksheet
        Set ws = ActiveWorkbook.Worksheets("Ангар " & i)
        
        If Err.Number = 0 Then
            availableCount = availableCount + 1
            If warehouseList <> "" Then warehouseList = warehouseList & ", "
            warehouseList = warehouseList & i
        End If
        On Error GoTo ErrorHandler
    Next i
    
    stats = stats & "?? Доступные ангары: " & availableCount & " из " & totalCount & vbNewLine
    
    If availableCount > 0 Then
        stats = stats & "?? Список ангаров: " & warehouseList & vbNewLine
    End If
    
    ' Информация о системе
    stats = stats & vbNewLine & "?? ИНФОРМАЦИЯ О СИСТЕМЕ:" & vbNewLine
    stats = stats & "?? Дата: " & Format(Date, "dd.mm.yyyy") & vbNewLine
    stats = stats & "? Время: " & Format(Time, "hh:mm:ss") & vbNewLine
    stats = stats & "??? Excel: " & Application.version & vbNewLine
    
    MsgBox stats, vbInformation, "Статистика системы"
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка получения статистики: " & Err.description, vbExclamation, "Ошибка"
End Sub

Public Sub СоздатьТестовыеДанные()
    ' Создать образец тестовых данных (русское название)
    Call CreateSampleData
End Sub

Public Sub CreateSampleData()
    ' Создает образец данных для тестирования
    Dim sampleData As String
    sampleData = "?? ОБРАЗЕЦ ДАННЫХ ДЛЯ ТЕСТИРОВАНИЯ" & vbNewLine & vbNewLine & _
                 "Скопируйте эти строки в форму массового ввода:" & vbNewLine & vbNewLine & _
                 "6-5-Е-2 - Лерашанс - пар13 - 720" & vbNewLine & _
                 "6-5-П-2 - Дикошанс - пар25 - 50" & vbNewLine & _
                 "6-6-А-1 - Лерашанс - пар13 - 720" & vbNewLine & _
                 "6-6-А-2 - Лерашанс - пар14 - 360" & vbNewLine & _
                 "7-1-А-1 - Босфор - пар03 - 720" & vbNewLine & _
                 "7-1-А-1 - Босфор - пар12 - 360" & vbNewLine & _
                 "8-2-В-2 - Чистосад - пар07 - 480" & vbNewLine & _
                 "9-3-М-1 - Микрошанс - пар15 - 240" & vbNewLine & vbNewLine & _
                 "?? ЧТО ПРОДЕМОНСТРИРУЕТ СИСТЕМА:" & vbNewLine & _
                 "• Разные секции: Е (верх) и П (низ) в ряду 5" & vbNewLine & _
                 "• Один товар в ряду: Лерашанс в ряду 6" & vbNewLine & _
                 "• Суммирование дублей: Босфор 720+360=1080" & vbNewLine & _
                 "• Работа с разными ангарами" & vbNewLine & _
                 "• Правильные заголовки в секциях"
    
    MsgBox sampleData, vbInformation, "Образец тестовых данных"
End Sub

Public Sub ПоказатьРаспределениеСекций()
    ' Показать распределение стеллажей по секциям для всех ангаров
    Call ShowSectionMapping
End Sub

Public Sub ShowSectionMapping()
    ' Показывает распределение стеллажей по секциям
    Dim warehouseNum As String
    warehouseNum = InputBox("Введите номер ангара для просмотра секций (5-12):", _
                           "Распределение по секциям", "6")
    
    If warehouseNum <> "" And IsNumeric(warehouseNum) Then
        If CInt(warehouseNum) >= 5 And CInt(warehouseNum) <= 12 Then
            Call ModuleHeaderManager.ShowSectionMappingForWarehouse(warehouseNum)
        Else
            MsgBox "Неверный номер ангара! Используйте числа от 5 до 12.", vbExclamation, "Ошибка"
        End If
    End If
End Sub

' ===== ФУНКЦИИ БЫСТРОГО ДОСТУПА =====

Public Sub СоздатьКнопкиБыстрогоДоступа()
    ' Инструкция по созданию кнопок быстрого доступа
    Call CreateQuickAccessButtons
End Sub

Public Sub CreateQuickAccessButtons()
    ' Показывает инструкцию по созданию кнопок быстрого доступа
    Dim instruction As String
    instruction = "?? СОЗДАНИЕ КНОПОК БЫСТРОГО ДОСТУПА" & vbNewLine & vbNewLine & _
                  "Для удобства работы создайте кнопки со следующими макросами:" & vbNewLine & vbNewLine & _
                  "?? ОСНОВНЫЕ КНОПКИ:" & vbNewLine & _
                  "• ModuleLauncher.ЗапуститьСистему" & vbNewLine & _
                  "  (Запуск системы массового ввода)" & vbNewLine & vbNewLine & _
                  "?? СЛУЖЕБНЫЕ КНОПКИ:" & vbNewLine & _
                  "• ModuleLauncher.ОчиститьВсе" & vbNewLine & _
                  "  (Очистка всех данных)" & vbNewLine & vbNewLine & _
                  "?? ИНФОРМАЦИОННЫЕ КНОПКИ:" & vbNewLine & _
                  "• ModuleLauncher.ПоказатьСправку" & vbNewLine & _
                  "  (Справка по системе)" & vbNewLine & _
                  "• ModuleLauncher.ТестСистемы" & vbNewLine & _
                  "  (Проверка работоспособности)" & vbNewLine & vbNewLine & _
                  "?? КАК СОЗДАТЬ:" & vbNewLine & _
                  "1. Правой кнопкой на панели инструментов" & vbNewLine & _
                  "2. Настройка > Команды > Макросы" & vbNewLine & _
                  "3. Перетащите нужные макросы на панель" & vbNewLine & _
                  "4. Настройте названия и значки кнопок"
    
    MsgBox instruction, vbInformation, "Создание кнопок быстрого доступа"
End Sub

' ===== ОТЛАДОЧНЫЕ ФУНКЦИИ =====

Public Sub ПоказатьВерсию()
    ' Показать версию системы
    Call ShowVersion
End Sub

Public Sub ShowVersion()
    ' Показывает информацию о версии системы
    Dim version As String
    version = "?? ИНФОРМАЦИЯ О ВЕРСИИ" & vbNewLine & vbNewLine & _
              "?? Система: Массовый ввод товаров на склад" & vbNewLine & _
              "??? Версия: 2.0 (с поддержкой секций)" & vbNewLine & _
              "?? Дата: " & Format(Date, "dd.mm.yyyy") & vbNewLine & vbNewLine & _
              "? НОВЫЕ ВОЗМОЖНОСТИ В ЭТОЙ ВЕРСИИ:" & vbNewLine & _
              "• Правильная работа с секциями ангаров" & vbNewLine & _
              "• Исправлены ложные конфликты товаров" & vbNewLine & _
              "• Улучшенное логирование операций" & vbNewLine & _
              "• Точный поиск ячеек по секциям" & vbNewLine & _
              "• Оптимизированные заголовки товаров" & vbNewLine & vbNewLine & _
              "?? ПОДДЕРЖИВАЕМЫЕ АНГАРЫ: 5, 6, 7, 8, 9, 10, 11, 12" & vbNewLine & _
              "?? МАКСИМАЛЬНОЕ КОЛИЧЕСТВО РЯДОВ: до 10" & vbNewLine & _
              "?? УРОВНИ СТЕЛЛАЖЕЙ: 1, 2, 3"
    
    MsgBox version, vbInformation, "Версия системы"
End Sub

' ===== АВАРИЙНЫЕ ФУНКЦИИ =====

Public Sub АварийнаяОстановка()
    ' Аварийная остановка всех процессов
    Call EmergencyStop
End Sub

Public Sub EmergencyStop()
    ' Функция аварийной остановки системы
    On Error Resume Next
    
    ' Останавливаем все процессы
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    ' Показываем сообщение
    MsgBox "?? АВАРИЙНАЯ ОСТАНОВКА ВЫПОЛНЕНА" & vbNewLine & vbNewLine & _
           "Все процессы остановлены:" & vbNewLine & _
           "• Обновление экрана включено" & vbNewLine & _
           "• Автоматические вычисления включены" & vbNewLine & _
           "• События Excel включены" & vbNewLine & vbNewLine & _
           "Система готова к нормальной работе.", _
           vbInformation, "Аварийная остановка"
End Sub

