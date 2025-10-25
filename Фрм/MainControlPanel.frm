VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainControlPanel 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "MainControlPanel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainControlPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===== MainControlPanel - ГЛАВНАЯ ПАНЕЛЬ УПРАВЛЕНИЯ СКЛАДОМ =====
Option Explicit

' Переменные для отслеживания состояния
Private lastBackupDate As Date
Private operationCount As Long

Private Sub UserForm_Initialize()
    ' Настройка размеров формы
    Me.Width = 900
    Me.Height = 700
    
    ' Заголовок формы
    Me.Caption = "?? СИСТЕМА УПРАВЛЕНИЯ СКЛАДОМ v2.0 - Главная панель"
    
    ' ===== ЗАГОЛОВОК =====
    With lblMainTitle
        .Caption = "СИСТЕМА УПРАВЛЕНИЯ СКЛАДОМ"
        .Top = 10
        .Left = 250
        .Width = 400
        .Height = 30
        .Font.Size = 18
        .Font.Bold = True
        .ForeColor = RGB(0, 50, 100)
        .TextAlign = fmTextAlignCenter
    End With
    
    ' ===== ИНФОРМАЦИОННАЯ ПАНЕЛЬ =====
    With lblInfoPanel
        .Caption = GetSystemInfo()
        .Top = 45
        .Left = 20
        .Width = 860
        .Height = 40
        .Font.Size = 9
        .BackColor = RGB(240, 248, 255)
        .BorderStyle = fmBorderStyleSingle
    End With
    
    ' ===== РАЗДЕЛ 1: ОСНОВНЫЕ ОПЕРАЦИИ =====
    With frameMainOperations
        .Caption = "?? ОСНОВНЫЕ ОПЕРАЦИИ"
        .Top = 95
        .Left = 20
        .Width = 280
        .Height = 200
        .Font.Bold = True
        .ForeColor = RGB(0, 100, 0)
    End With
    
    ' Кнопка: Массовый ввод с координатами
    With btnMassInput
        .Caption = "?? Массовый ввод" & vbNewLine & "(с координатами)"
        .Top = 115
        .Left = 30
        .Width = 125
        .Height = 50
        .BackColor = RGB(100, 200, 100)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 10
    End With
    
    ' Кнопка: Умное размещение
    With btnSmartPlacement
        .Caption = "?? Умное размещение" & vbNewLine & "(автоматическое)"
        .Top = 115
        .Left = 165
        .Width = 125
        .Height = 50
        .BackColor = RGB(255, 200, 50)
        .ForeColor = RGB(0, 0, 0)
        .Font.Bold = True
        .Font.Size = 10
    End With
    
    ' Кнопка: Поиск товара
    With btnSearchProduct
        .Caption = "?? Поиск товара"
        .Top = 175
        .Left = 30
        .Width = 125
        .Height = 40
        .BackColor = RGB(100, 150, 200)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    ' Кнопка: Перемещение
    With btnMoveProduct
        .Caption = "?? Перемещение"
        .Top = 175
        .Left = 165
        .Width = 125
        .Height = 40
        .BackColor = RGB(150, 100, 200)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    ' Кнопка: Очистить все
    With btnClearAll
        .Caption = "??? Очистить все"
        .Top = 225
        .Left = 30
        .Width = 260
        .Height = 35
        .BackColor = RGB(220, 50, 50)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    ' Кнопка: Обработка спец.рядов
    With btnSpecialRows
        .Caption = "?? Обработка спец.рядов"
        .Top = 270
        .Left = 30
        .Width = 260
        .Height = 35
        .BackColor = RGB(200, 150, 100)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    ' ===== РАЗДЕЛ 2: БЫСТРЫЕ ОПЕРАЦИИ =====
    With frameQuickOperations
        .Caption = "? БЫСТРЫЕ ОПЕРАЦИИ"
        .Top = 95
        .Left = 310
        .Width = 280
        .Height = 200
        .Font.Bold = True
        .ForeColor = RGB(200, 100, 0)
    End With
    
    ' Кнопка: Быстрое добавление
    With btnQuickAdd
        .Caption = "? Добавить быстро"
        .Top = 115
        .Left = 320
        .Width = 125
        .Height = 40
        .BackColor = RGB(50, 200, 50)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    ' Кнопка: Быстрое списание
    With btnQuickRemove
        .Caption = "? Списать быстро"
        .Top = 115
        .Left = 455
        .Width = 125
        .Height = 40
        .BackColor = RGB(200, 50, 50)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    ' Кнопка: Быстрая инвентаризация
    With btnQuickInventory
        .Caption = "?? Инвентаризация"
        .Top = 165
        .Left = 320
        .Width = 125
        .Height = 40
        .BackColor = RGB(100, 100, 200)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    ' Кнопка: Проверка партии
    With btnCheckBatch
        .Caption = "??? Проверка партии"
        .Top = 165
        .Left = 455
        .Width = 125
        .Height = 40
        .BackColor = RGB(200, 100, 200)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    ' Кнопка: Корректировка
    With btnCorrection
        .Caption = "?? Корректировка остатков"
        .Top = 215
        .Left = 320
        .Width = 260
        .Height = 35
        .BackColor = RGB(200, 150, 50)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    ' Кнопка: Перенос между секциями
    With btnTransferSections
        .Caption = "¦? Перенос между секциями"
        .Top = 260
        .Left = 320
        .Width = 260
        .Height = 35
        .BackColor = RGB(150, 150, 200)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    ' ===== РАЗДЕЛ 3: ОТЧЕТЫ И АНАЛИЗ =====
    With frameReports
        .Caption = "?? ОТЧЕТЫ И АНАЛИЗ"
        .Top = 95
        .Left = 600
        .Width = 280
        .Height = 200
        .Font.Bold = True
        .ForeColor = RGB(0, 50, 150)
    End With
    
    ' Кнопка: Итоги по ангарам
    With btnWarehouseTotals
        .Caption = "?? Итоги по ангарам"
        .Top = 115
        .Left = 610
        .Width = 125
        .Height = 40
        .BackColor = RGB(100, 150, 255)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    ' Кнопка: Анализ заполненности
    With btnOccupancyAnalysis
        .Caption = "?? Заполненность"
        .Top = 115
        .Left = 745
        .Width = 125
        .Height = 40
        .BackColor = RGB(150, 200, 255)
        .ForeColor = RGB(0, 0, 0)
        .Font.Bold = True
    End With
    
    ' Кнопка: Остатки по товарам
    With btnProductStock
        .Caption = "?? Остатки товаров"
        .Top = 165
        .Left = 610
        .Width = 125
        .Height = 40
        .BackColor = RGB(200, 150, 255)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    ' Кнопка: История операций
    With btnOperationHistory
        .Caption = "?? История"
        .Top = 165
        .Left = 745
        .Width = 125
        .Height = 40
        .BackColor = RGB(255, 200, 150)
        .ForeColor = RGB(0, 0, 0)
        .Font.Bold = True
    End With
    
    ' Кнопка: Полный отчет
    With btnFullReport
        .Caption = "?? Полный отчет по складу"
        .Top = 215
        .Left = 610
        .Width = 260
        .Height = 35
        .BackColor = RGB(50, 100, 200)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    ' Кнопка: Аналитика движения
    With btnMovementAnalytics
        .Caption = "?? Аналитика движения товаров"
        .Top = 260
        .Left = 610
        .Width = 260
        .Height = 35
        .BackColor = RGB(100, 50, 200)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    ' ===== РАЗДЕЛ 4: РАБОТА С ДАННЫМИ =====
    With frameDataOperations
        .Caption = "?? РАБОТА С ДАННЫМИ"
        .Top = 310
        .Left = 20
        .Width = 430
        .Height = 140
        .Font.Bold = True
        .ForeColor = RGB(100, 0, 100)
    End With
    
    ' Кнопка: Резервная копия
    With btnBackup
        .Caption = "?? Создать резервную копию"
        .Top = 330
        .Left = 30
        .Width = 200
        .Height = 40
        .BackColor = RGB(0, 150, 0)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 10
    End With
    
    ' Кнопка: Восстановление
    With btnRestore
        .Caption = "?? Восстановить из копии"
        .Top = 330
        .Left = 240
        .Width = 200
        .Height = 40
        .BackColor = RGB(150, 100, 0)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 10
    End With
    
    ' Кнопка: Импорт
    With btnImport
        .Caption = "?? Импорт из Excel"
        .Top = 380
        .Left = 30
        .Width = 130
        .Height = 35
        .BackColor = RGB(100, 200, 200)
        .ForeColor = RGB(0, 0, 0)
        .Font.Bold = True
    End With
    
    ' Кнопка: Экспорт
    With btnExport
        .Caption = "?? Экспорт в Excel"
        .Top = 380
        .Left = 170
        .Width = 130
        .Height = 35
        .BackColor = RGB(200, 200, 100)
        .ForeColor = RGB(0, 0, 0)
        .Font.Bold = True
    End With
    
    ' Кнопка: Синхронизация
    With btnSync
        .Caption = "?? Синхронизация"
        .Top = 380
        .Left = 310
        .Width = 130
        .Height = 35
        .BackColor = RGB(150, 150, 150)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    ' ===== РАЗДЕЛ 5: НАСТРОЙКИ И СПРАВОЧНИКИ =====
    With frameSettings
        .Caption = "?? НАСТРОЙКИ И СПРАВОЧНИКИ"
        .Top = 310
        .Left = 460
        .Width = 420
        .Height = 140
        .Font.Bold = True
        .ForeColor = RGB(50, 50, 50)
    End With
    
    ' Кнопка: База товаров
    With btnProductDatabase
        .Caption = "?? База товаров"
        .Top = 330
        .Left = 470
        .Width = 130
        .Height = 40
        .BackColor = RGB(100, 150, 200)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    ' Кнопка: Настройки системы
    With btnSystemSettings
        .Caption = "?? Настройки"
        .Top = 330
        .Left = 610
        .Width = 130
        .Height = 40
        .BackColor = RGB(150, 150, 150)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    ' Кнопка: Пользователи
    With btnUsers
        .Caption = "?? Пользователи"
        .Top = 330
        .Left = 750
        .Width = 120
        .Height = 40
        .BackColor = RGB(200, 150, 100)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    ' Кнопка: Логи системы
    With btnSystemLogs
        .Caption = "?? Логи системы"
        .Top = 380
        .Left = 470
        .Width = 130
        .Height = 35
        .BackColor = RGB(100, 100, 100)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    ' Кнопка: Справка
    With btnHelp
        .Caption = "? Справка"
        .Top = 380
        .Left = 610
        .Width = 130
        .Height = 35
        .BackColor = RGB(100, 200, 100)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    ' Кнопка: О программе
    With btnAbout
        .Caption = "?? О программе"
        .Top = 380
        .Left = 750
        .Width = 120
        .Height = 35
        .BackColor = RGB(200, 100, 200)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    ' ===== РАЗДЕЛ 6: СПЕЦИАЛЬНЫЕ ФУНКЦИИ =====
    With frameSpecial
        .Caption = "?? СПЕЦИАЛЬНЫЕ ФУНКЦИИ"
        .Top = 460
        .Left = 20
        .Width = 860
        .Height = 80
        .Font.Bold = True
        .ForeColor = RGB(150, 0, 0)
    End With
    
    ' Кнопка: Оптимизация размещения
    With btnOptimize
        .Caption = "?? Оптимизация размещения"
        .Top = 480
        .Left = 30
        .Width = 200
        .Height = 35
        .BackColor = RGB(255, 150, 0)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    ' Кнопка: Консолидация партий
    With btnConsolidate
        .Caption = "?? Консолидация партий"
        .Top = 480
        .Left = 240
        .Width = 200
        .Height = 35
        .BackColor = RGB(0, 150, 255)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    ' Кнопка: Проверка целостности
    With btnIntegrityCheck
        .Caption = "?? Проверка целостности"
        .Top = 480
        .Left = 450
        .Width = 200
        .Height = 35
        .BackColor = RGB(150, 0, 150)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    ' Кнопка: Форматирование таблиц
    With btnFormatTables
        .Caption = "?? Форматирование таблиц"
        .Top = 480
        .Left = 660
        .Width = 200
        .Height = 35
        .BackColor = RGB(100, 200, 150)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    ' ===== СТАТУСНАЯ СТРОКА =====
    With lblStatus
        .Caption = "? Система готова к работе | Последняя операция: нет | Операций сегодня: 0"
        .Top = 550
        .Left = 20
        .Width = 860
        .Height = 25
        .Font.Size = 9
        .BackColor = RGB(240, 240, 240)
        .BorderStyle = fmBorderStyleSingle
        .ForeColor = RGB(0, 100, 0)
    End With
    
    ' ===== КНОПКИ УПРАВЛЕНИЯ =====
    ' Кнопка: Выход
    With btnExit
        .Caption = "? ВЫХОД"
        .Top = 590
        .Left = 750
        .Width = 120
        .Height = 40
        .BackColor = RGB(200, 0, 0)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    ' Кнопка: Свернуть
    With btnMinimize
        .Caption = "? Свернуть"
        .Top = 590
        .Left = 620
        .Width = 120
        .Height = 40
        .BackColor = RGB(150, 150, 150)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    ' ===== ДОПОЛНИТЕЛЬНЫЕ ЭЛЕМЕНТЫ =====
    ' Индикатор резервного копирования
    With lblBackupIndicator
        .Caption = GetBackupStatus()
        .Top = 590
        .Left = 20
        .Width = 300
        .Height = 20
        .Font.Size = 8
        .ForeColor = GetBackupColor()
    End With
    
    ' Индикатор активности
    With lblActivityIndicator
        .Caption = "?? Онлайн"
        .Top = 590
        .Left = 330
        .Width = 100
        .Height = 20
        .Font.Size = 8
        .ForeColor = RGB(0, 150, 0)
    End With
    
    ' Версия программы
    With lblVersion
        .Caption = "v2.0.1"
        .Top = 610
        .Left = 20
        .Width = 50
        .Height = 20
        .Font.Size = 8
        .ForeColor = RGB(100, 100, 100)
    End With
    
    ' Инициализация переменных
    operationCount = 0
    lastBackupDate = Date - 7 ' Предполагаем, что давно не было резервной копии
    
    ' Проверка при запуске
    CheckSystemState
End Sub

' ===== ОБРАБОТЧИКИ ОСНОВНЫХ ОПЕРАЦИЙ =====

Private Sub btnMassInput_Click()
    UpdateStatus "Открытие формы массового ввода..."
    MassProductInputForm.Show
    IncrementOperationCount
End Sub

Private Sub btnSmartPlacement_Click()
    UpdateStatus "Открытие умного размещения..."
    SmartPlacementForm.Show
    IncrementOperationCount
End Sub

Private Sub btnSearchProduct_Click()
    Dim ProductName As String
    ProductName = InputBox("Введите название товара для поиска:", "Поиск товара")
    
    If ProductName <> "" Then
        UpdateStatus "Поиск товара: " & ProductName
        ' Здесь должен быть вызов модуля поиска
        MsgBox "Функция поиска товара '" & ProductName & "' будет реализована", vbInformation
        IncrementOperationCount
    End If
End Sub

Private Sub btnMoveProduct_Click()
    UpdateStatus "Открытие формы перемещения..."
    MsgBox "Форма перемещения товаров будет реализована", vbInformation
    IncrementOperationCount
End Sub

Private Sub btnClearAll_Click()
    If MsgBox("Вы действительно хотите очистить ВСЕ ангары?", _
              vbYesNo + vbExclamation, "Подтверждение") = vbYes Then
        UpdateStatus "Очистка всех ангаров..."
        Dim result As String
        result = ModuleDataCleaner.ClearAllWarehouses()
        If result = "" Then
            UpdateStatus "? Все ангары очищены"
        Else
            UpdateStatus "?? Ошибка при очистке"
        End If
        IncrementOperationCount
    End If
End Sub

Private Sub btnSpecialRows_Click()
    UpdateStatus "Обработка специальных рядов..."
    Call ModuleSpecialCases.ProcessSpecialRows
    UpdateStatus "? Специальные ряды обработаны"
    IncrementOperationCount
End Sub

' ===== ОБРАБОТЧИКИ БЫСТРЫХ ОПЕРАЦИЙ =====

Private Sub btnQuickAdd_Click()
    Dim input As String
    input = InputBox("Быстрое добавление:" & vbNewLine & _
                    "Формат: Товар - Партия - Количество", _
                    "Быстрое добавление")
    
    If input <> "" Then
        UpdateStatus "Быстрое добавление: " & input
        ' Вызов модуля умного размещения
        Dim result As String
        result = ModuleSmartPlacement.SmartPlaceProducts(input)
        MsgBox result, vbInformation
        IncrementOperationCount
    End If
End Sub

Private Sub btnQuickRemove_Click()
    UpdateStatus "Быстрое списание..."
    MsgBox "Функция быстрого списания будет реализована", vbInformation
End Sub

Private Sub btnQuickInventory_Click()
    UpdateStatus "Запуск инвентаризации..."
    MsgBox "Модуль инвентаризации будет реализован", vbInformation
End Sub

Private Sub btnCheckBatch_Click()
    Dim Batch As String
    Batch = InputBox("Введите номер партии для проверки:", "Проверка партии")
    
    If Batch <> "" Then
        UpdateStatus "Проверка партии: " & Batch
        MsgBox "Функция проверки партии '" & Batch & "' будет реализована", vbInformation
        IncrementOperationCount
    End If
End Sub

' ===== ОБРАБОТЧИКИ ОТЧЕТОВ =====

Private Sub btnWarehouseTotals_Click()
    UpdateStatus "Расчет итогов по ангарам..."
    Call CalculateAllWarehouses.CalculateAllWarehouses
    UpdateStatus "? Итоги рассчитаны"
    IncrementOperationCount
End Sub

Private Sub btnOccupancyAnalysis_Click()
    UpdateStatus "Анализ заполненности..."
    MsgBox "Модуль анализа заполненности будет реализован", vbInformation
End Sub

Private Sub btnProductStock_Click()
    UpdateStatus "Формирование отчета по остаткам..."
    MsgBox "Отчет по остаткам товаров будет реализован", vbInformation
End Sub

Private Sub btnOperationHistory_Click()
    UpdateStatus "Загрузка истории операций..."
    Call ModuleLogger.ShowLogBuffer
End Sub

' ===== ОБРАБОТЧИКИ РАБОТЫ С ДАННЫМИ =====

Private Sub btnBackup_Click()
    UpdateStatus "Создание резервной копии..."
    
    ' Сохраняем текущую книгу с новым именем
    Dim backupPath As String
    backupPath = ThisWorkbook.Path & "\Backup_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsm"
    
    On Error Resume Next
    ThisWorkbook.SaveCopyAs backupPath
    
    If Err.Number = 0 Then
        lastBackupDate = Date
        UpdateStatus "? Резервная копия создана: " & backupPath
        MsgBox "Резервная копия успешно создана!" & vbNewLine & backupPath, vbInformation
    Else
        UpdateStatus "? Ошибка создания резервной копии"
        MsgBox "Ошибка при создании резервной копии!", vbExclamation
    End If
    On Error GoTo 0
    
    lblBackupIndicator.Caption = GetBackupStatus()
    lblBackupIndicator.ForeColor = GetBackupColor()
    IncrementOperationCount
End Sub

Private Sub btnRestore_Click()
    UpdateStatus "Восстановление из резервной копии..."
    MsgBox "Функция восстановления будет реализована", vbInformation
End Sub

Private Sub btnImport_Click()
    UpdateStatus "Импорт данных..."
    MsgBox "Функция импорта из Excel будет реализована", vbInformation
End Sub

Private Sub btnExport_Click()
    UpdateStatus "Экспорт данных..."
    Call ModuleProductInfo.ExportDatabaseToExcel
    UpdateStatus "? Данные экспортированы"
    IncrementOperationCount
End Sub

' ===== ОБРАБОТЧИКИ НАСТРОЕК =====

Private Sub btnProductDatabase_Click()
    UpdateStatus "Открытие базы товаров..."
    Call ModuleProductInfo.ShowAllProducts
    MsgBox "База товаров выведена в Immediate Window (Ctrl+G)", vbInformation
    IncrementOperationCount
End Sub

Private Sub btnSystemSettings_Click()
    UpdateStatus "Открытие настроек..."
    MsgBox "Форма настроек системы будет реализована", vbInformation
End Sub

Private Sub btnSystemLogs_Click()
    UpdateStatus "Просмотр логов..."
    Call ModuleLogger.ShowLogBuffer
End Sub

Private Sub btnHelp_Click()
    ShowHelp
End Sub

Private Sub btnAbout_Click()
    ShowAbout
End Sub

' ===== ОБРАБОТЧИКИ СПЕЦИАЛЬНЫХ ФУНКЦИЙ =====

Private Sub btnOptimize_Click()
    UpdateStatus "Оптимизация размещения..."
    MsgBox "Модуль оптимизации размещения будет реализован", vbInformation
End Sub

Private Sub btnConsolidate_Click()
    UpdateStatus "Консолидация партий..."
    MsgBox "Модуль консолидации партий будет реализован", vbInformation
End Sub

Private Sub btnIntegrityCheck_Click()
    UpdateStatus "Проверка целостности данных..."
    MsgBox "Модуль проверки целостности будет реализован", vbInformation
End Sub

Private Sub btnFormatTables_Click()
    UpdateStatus "Форматирование таблиц..."
    MsgBox "Модуль форматирования таблиц будет реализован", vbInformation
End Sub

' ===== ОБРАБОТЧИКИ УПРАВЛЕНИЯ =====

Private Sub btnExit_Click()
    If operationCount > 0 Then
        If MsgBox("Выполнено операций: " & operationCount & vbNewLine & _
                  "Вы действительно хотите выйти?", _
                  vbYesNo + vbQuestion, "Подтверждение выхода") = vbNo Then
            Exit Sub
        End If
    End If
    
    ' Предложение создать резервную копию
    If Date - lastBackupDate > 3 Then
        If MsgBox("Резервная копия не создавалась более 3 дней." & vbNewLine & _
                  "Создать резервную копию перед выходом?", _
                  vbYesNo + vbQuestion, "Резервная копия") = vbYes Then
            btnBackup_Click
        End If
    End If
    
    Unload Me
End Sub

Private Sub btnMinimize_Click()
    Me.Hide
    MsgBox "Панель свернута. Для возврата запустите макрос ShowMainPanel", vbInformation
End Sub

' ===== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ =====

Private Function GetSystemInfo() As String
    Dim info As String
    info = "?? " & Format(Date, "dd.mm.yyyy") & " | "
    info = info & "?? Ангары: 5-12 | "
    info = info & "?? Пользователь: " & Application.UserName & " | "
    info = info & "?? " & ThisWorkbook.Name
    GetSystemInfo = info
End Function

Private Function GetBackupStatus() As String
    Dim daysSinceBackup As Long
    daysSinceBackup = Date - lastBackupDate
    
    If daysSinceBackup = 0 Then
        GetBackupStatus = "?? Резервная копия: сегодня"
    ElseIf daysSinceBackup = 1 Then
        GetBackupStatus = "?? Резервная копия: вчера"
    Else
        GetBackupStatus = "?? Резервная копия: " & daysSinceBackup & " дней назад"
    End If
End Function

Private Function GetBackupColor() As Long
    Dim daysSinceBackup As Long
    daysSinceBackup = Date - lastBackupDate
    
    If daysSinceBackup <= 1 Then
        GetBackupColor = RGB(0, 150, 0) ' Зеленый
    ElseIf daysSinceBackup <= 3 Then
        GetBackupColor = RGB(200, 150, 0) ' Желтый
    Else
        GetBackupColor = RGB(200, 0, 0) ' Красный
    End If
End Function

Private Sub UpdateStatus(message As String)
    lblStatus.Caption = "? " & message & " | Операций сегодня: " & operationCount
    DoEvents
End Sub

Private Sub IncrementOperationCount()
    operationCount = operationCount + 1
    lblStatus.Caption = "? Готово | Последняя операция: " & Format(Now, "hh:mm:ss") & _
                       " | Операций сегодня: " & operationCount
End Sub

Private Sub CheckSystemState()
    ' Проверка состояния системы при запуске
    If Date - lastBackupDate > 7 Then
        MsgBox "?? ВНИМАНИЕ!" & vbNewLine & vbNewLine & _
               "Резервная копия не создавалась более недели!" & vbNewLine & _
               "Рекомендуется создать резервную копию.", _
               vbExclamation, "Предупреждение"
    End If
End Sub

Private Sub ShowHelp()
    Dim helpText As String
    helpText = "?? СПРАВКА ПО СИСТЕМЕ" & vbNewLine & vbNewLine
    helpText = helpText & "ОСНОВНЫЕ ОПЕРАЦИИ:" & vbNewLine
    helpText = helpText & "• Массовый ввод - ввод с указанием координат" & vbNewLine
    helpText = helpText & "• Умное размещение - автоматический поиск места" & vbNewLine
    helpText = helpText & "• Поиск - найти товар на складе" & vbNewLine & vbNewLine
    helpText = helpText & "БЫСТРЫЕ ОПЕРАЦИИ:" & vbNewLine
    helpText = helpText & "• Быстрое добавление/списание" & vbNewLine
    helpText = helpText & "• Инвентаризация и проверка партий" & vbNewLine & vbNewLine
    helpText = helpText & "ГОРЯЧИЕ КЛАВИШИ:" & vbNewLine
    helpText = helpText & "• Ctrl+S - Создать резервную копию" & vbNewLine
    helpText = helpText & "• Ctrl+F - Поиск товара" & vbNewLine
    helpText = helpText & "• F5 - Обновить данные" & vbNewLine
    
    MsgBox helpText, vbInformation, "Справка"
End Sub

Private Sub ShowAbout()
    Dim aboutText As String
    aboutText = "?? СИСТЕМА УПРАВЛЕНИЯ СКЛАДОМ v2.0" & vbNewLine & vbNewLine
    aboutText = aboutText & "Разработано для управления складами 5-12" & vbNewLine & vbNewLine
    aboutText = aboutText & "ВОЗМОЖНОСТИ:" & vbNewLine
    aboutText = aboutText & "? Умное размещение товаров" & vbNewLine
    aboutText = aboutText & "? Автоматический учет лимитов" & vbNewLine
    aboutText = aboutText & "? Массовый ввод данных" & vbNewLine
    aboutText = aboutText & "? Полная отчетность" & vbNewLine
    aboutText = aboutText & "? Резервное копирование" & vbNewLine & vbNewLine
    aboutText = aboutText & "© 2024 Warehouse Management System"
    
    MsgBox aboutText, vbInformation, "О программе"
End Sub

' ===== ГОРЯЧИЕ КЛАВИШИ =====
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyS
            If Shift = 2 Then ' Ctrl+S
                btnBackup_Click
            End If
        Case vbKeyF
            If Shift = 2 Then ' Ctrl+F
                btnSearchProduct_Click
            End If
        Case vbKeyF5 ' F5
            CheckSystemState
            UpdateStatus "? Данные обновлены"
    End Select
End Sub

